import { strFromU8, unzipSync } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { Node as PMNode, type NodeType } from 'prosemirror-model';
import { docxSchema } from './schema.js';

/**
 * fast-xml-parser를 preserveOrder 모드로 쓴다 — OOXML은 같은 태그(w:p, w:tbl 등)가
 * 반복되며 형제 순서가 의미를 가지므로(문단과 표가 번갈아 나오는 문서 등),
 * 태그별로 묶어버리는 기본 모드는 순서 정보를 잃어버려 부적합하다.
 */
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  textNodeName: '#text',
  trimValues: false,
});

type XmlNode = Record<string, unknown>;

function tagOf(node: XmlNode): string | undefined {
  return Object.keys(node).find((key) => key !== ':@');
}

function attrsOf(node: XmlNode): Record<string, string> {
  return (node[':@'] as Record<string, string> | undefined) ?? {};
}

function childrenOf(node: XmlNode): XmlNode[] {
  const tag = tagOf(node);
  if (!tag) return [];
  const value = node[tag];
  return Array.isArray(value) ? (value as XmlNode[]) : [];
}

function findAll(list: readonly XmlNode[], tag: string): XmlNode[] {
  return list.filter((node) => tagOf(node) === tag);
}

function findFirst(list: readonly XmlNode[], tag: string): XmlNode | undefined {
  return list.find((node) => tagOf(node) === tag);
}

function textOf(list: readonly XmlNode[]): string {
  return list
    .filter((node) => '#text' in node)
    .map((node) => String((node as { '#text': unknown })['#text']))
    .join('');
}

/** w:pStyle val="Heading1".."Heading6" → 제목 레벨. 그 외(본문 등)는 문단으로 취급한다. */
function headingLevel(pPrChildren: readonly XmlNode[]): number | null {
  const pStyle = findFirst(pPrChildren, 'w:pStyle');
  if (!pStyle) return null;
  const val = attrsOf(pStyle)['@_w:val'] ?? '';
  const match = /^Heading([1-6])$/i.exec(val);
  return match ? Number(match[1]) : null;
}

const JC_TO_ALIGN: Record<string, string> = { center: 'center', right: 'right', both: 'justify' };

/** w:jc val="center"/"right"/"both" → ProseMirror align 속성. left/미지정은 null(기본값)로 둔다. */
function alignOf(pPrChildren: readonly XmlNode[]): string | null {
  const jc = findFirst(pPrChildren, 'w:jc');
  if (!jc) return null;
  const val = attrsOf(jc)['@_w:val'] ?? '';
  return JC_TO_ALIGN[val] ?? null;
}

interface RunMarks {
  readonly strong: boolean;
  readonly em: boolean;
  readonly underline: boolean;
  readonly strike: boolean;
  readonly color: string | null;
  readonly fontSizePt: number | null;
}

function runMarksOf(runChildren: readonly XmlNode[]): RunMarks {
  const rPr = findFirst(runChildren, 'w:rPr');
  const rPrChildren = rPr ? childrenOf(rPr) : [];
  const has = (tag: string) => findFirst(rPrChildren, tag) !== undefined;
  const underline = findFirst(rPrChildren, 'w:u');
  const color = findFirst(rPrChildren, 'w:color');
  const sz = findFirst(rPrChildren, 'w:sz');
  const colorVal = color ? attrsOf(color)['@_w:val'] : undefined;
  const szVal = sz ? attrsOf(sz)['@_w:val'] : undefined;
  return {
    strong: has('w:b'),
    em: has('w:i'),
    underline: underline !== undefined && attrsOf(underline)['@_w:val'] !== 'none',
    strike: has('w:strike'),
    color: colorVal && /^[0-9A-Fa-f]{6}$/.test(colorVal) ? `#${colorVal.toLowerCase()}` : null,
    fontSizePt: szVal && !Number.isNaN(Number(szVal)) ? Number(szVal) / 2 : null,
  };
}

/** 마크가 적용된 텍스트 조각을 ProseMirror text 노드로 변환한다 (빈 문자열은 제외). */
function textNodeWithMarks(text: string, marks: RunMarks): PMNode | null {
  if (text === '') return null;
  const markList = [
    marks.strong ? docxSchema.marks.strong.create() : null,
    marks.em ? docxSchema.marks.em.create() : null,
    marks.underline ? docxSchema.marks.underline.create() : null,
    marks.strike ? docxSchema.marks.strike.create() : null,
    marks.color ? docxSchema.marks.textColor.create({ color: marks.color }) : null,
    marks.fontSizePt ? docxSchema.marks.fontSize.create({ size: marks.fontSizePt }) : null,
  ].filter((mark): mark is NonNullable<typeof mark> => mark !== null);
  return docxSchema.text(text, markList);
}

/**
 * w:r의 자식들 중 w:t 요소(들)의 텍스트를 이어붙인다.
 * findAll이 반환하는 것은 w:t "래퍼" 노드(예: { 'w:t': [...], ':@': {...} })이므로,
 * 텍스트는 그 안의 childrenOf(wt)를 다시 한 번 textOf로 펼쳐야 나온다.
 */
function runTextOf(runChildren: readonly XmlNode[]): string {
  return findAll(runChildren, 'w:t')
    .map((wt) => textOf(childrenOf(wt)))
    .join('');
}

/** w:p의 자식을 순회하며 인라인(텍스트) 콘텐츠를 뽑아낸다. 이미지(w:drawing)는 건너뛴다(v1 미지원). */
function inlineContentOf(pChildren: readonly XmlNode[]): PMNode[] {
  const nodes: PMNode[] = [];
  for (const run of findAll(pChildren, 'w:r')) {
    const runChildren = childrenOf(run);
    const node = textNodeWithMarks(runTextOf(runChildren), runMarksOf(runChildren));
    if (node) nodes.push(node);
  }
  return nodes;
}

function paragraphFrom(pNode: XmlNode): { node: PMNode; isListItem: boolean } {
  const pChildren = childrenOf(pNode);
  const pPr = findFirst(pChildren, 'w:pPr');
  const pPrChildren = pPr ? childrenOf(pPr) : [];
  const level = headingLevel(pPrChildren);
  const align = alignOf(pPrChildren);
  const isListItem = findFirst(pPrChildren, 'w:numPr') !== undefined;
  const inline = inlineContentOf(pChildren);
  const node =
    level !== null
      ? docxSchema.nodes.heading.create({ level, align }, inline)
      : docxSchema.nodes.paragraph.create({ align }, inline);
  return { node, isListItem };
}

function tableCellFrom(tcNode: XmlNode, colspan: number, rowspan: number): PMNode {
  const tcChildren = childrenOf(tcNode);
  const paragraphs = findAll(tcChildren, 'w:p').map((p) => paragraphFrom(p).node);
  const content = paragraphs.length > 0 ? paragraphs : [docxSchema.nodes.paragraph.create()];
  return docxSchema.nodes.table_cell.create({ colspan, rowspan }, content);
}

function tcPrChildrenOf(tcNode: XmlNode): XmlNode[] {
  const tcPr = findFirst(childrenOf(tcNode), 'w:tcPr');
  return tcPr ? childrenOf(tcPr) : [];
}

/** w:tcPr/w:gridSpan@w:val — 가로 병합 칸 수. 없으면 1(병합 없음). */
function gridSpanOf(tcPrChildren: readonly XmlNode[]): number {
  const gridSpan = findFirst(tcPrChildren, 'w:gridSpan');
  if (!gridSpan) return 1;
  const val = Number(attrsOf(gridSpan)['@_w:val']);
  return Number.isFinite(val) && val > 0 ? val : 1;
}

/**
 * w:tcPr/w:vMerge — 세로 병합 상태. 'restart'면 이 칸이 병합의 시작(눈에 보이는 셀),
 * @_w:val이 없거나 "continue"면 위 칸에 이어지는 자리(빈 자리표시자, 실제 셀로 만들지 않는다).
 * 태그 자체가 없으면 세로 병합과 무관한 일반 칸이다.
 */
function vMergeOf(tcPrChildren: readonly XmlNode[]): 'restart' | 'continue' | null {
  const vMerge = findFirst(tcPrChildren, 'w:vMerge');
  if (!vMerge) return null;
  return attrsOf(vMerge)['@_w:val'] === 'restart' ? 'restart' : 'continue';
}

interface CellPlacement {
  readonly tc: XmlNode;
  readonly colspan: number;
  rowspan: number;
}

/**
 * OOXML의 표 병합(w:gridSpan=가로, w:vMerge=세로)을 ProseMirror의 colspan/rowspan 모델로
 * 옮긴다. ProseMirror(prosemirror-tables)는 세로 병합을 "이어지는 칸"을 아예 만들지 않고
 * 시작 칸의 rowspan만 늘리는 방식으로 표현하므로, w:vMerge가 "continue"인 <w:tc>는 새 셀을
 * 만들지 않고 같은 열의 시작 칸(originByCol) rowspan만 1 늘린 뒤 건너뛴다.
 */
function tableFrom(tblNode: XmlNode): PMNode {
  const tblChildren = childrenOf(tblNode);
  const trNodes = findAll(tblChildren, 'w:tr');

  const rowPlacements: CellPlacement[][] = [];
  const originByCol = new Map<number, CellPlacement>();

  for (const tr of trNodes) {
    const placements: CellPlacement[] = [];
    let col = 0;
    for (const tc of findAll(childrenOf(tr), 'w:tc')) {
      const tcPrChildren = tcPrChildrenOf(tc);
      const gridSpan = gridSpanOf(tcPrChildren);
      const vMerge = vMergeOf(tcPrChildren);
      if (vMerge === 'continue') {
        const origin = originByCol.get(col);
        if (origin) origin.rowspan += 1;
        col += gridSpan;
        continue;
      }
      const placement: CellPlacement = { tc, colspan: gridSpan, rowspan: 1 };
      placements.push(placement);
      originByCol.set(col, placement);
      col += gridSpan;
    }
    rowPlacements.push(placements);
  }

  const rows = rowPlacements.map((placements) =>
    docxSchema.nodes.table_row.create(
      null,
      placements.map((p) => tableCellFrom(p.tc, p.colspan, p.rowspan)),
    ),
  );
  return docxSchema.nodes.table.create(null, rows);
}

/**
 * 연속된 목록 문단(w:numPr 보유)을 bullet_list로 묶는다.
 * 원본이 번호 매기기였는지 글머리 기호였는지는(numbering.xml 미해석) 구분하지 않고
 * 전부 글머리 기호 목록으로 가져온다 — v1 스코프의 알려진 단순화.
 */
function groupListItems(blocks: ReadonlyArray<{ node: PMNode; isListItem: boolean }>): PMNode[] {
  const result: PMNode[] = [];
  let pendingItems: PMNode[] = [];
  const flush = () => {
    if (pendingItems.length === 0) return;
    const items = pendingItems.map((paragraph) => docxSchema.nodes.list_item.create(null, paragraph));
    result.push(docxSchema.nodes.bullet_list.create(null, items));
    pendingItems = [];
  };
  for (const block of blocks) {
    if (block.isListItem) {
      pendingItems.push(block.node);
    } else {
      flush();
      result.push(block.node);
    }
  }
  flush();
  return result;
}

export interface DocxImportResult {
  readonly doc: PMNode;
  /** 이미지 등 v1 미지원 요소가 포함되어 편집·재저장 시 유실될 수 있는지 */
  readonly hasUnsupportedContent: boolean;
}

/**
 * 이미지/OLE 개체 등 v1 미지원 요소가 있는지만 빠르게 검사한다(전체 파싱 없이).
 * 편집기 진입 전 "저장하면 이미지가 사라집니다" 경고를 띄우는 용도(src/routes/viewer.ts).
 */
export function docxHasUnsupportedContent(docxBytes: Uint8Array): boolean {
  const entries = unzipSync(docxBytes);
  const documentXmlBytes = entries['word/document.xml'];
  if (!documentXmlBytes) return false;
  return /<w:drawing\b|<w:object\b|<w:pict\b/.test(strFromU8(documentXmlBytes));
}

/** docx(zip) 바이트 → ProseMirror 문서. 알려지지 않은 요소(이미지·머리글 등)는 건너뛴다. */
export function importDocx(docxBytes: Uint8Array): DocxImportResult {
  const entries = unzipSync(docxBytes);
  const documentXmlBytes = entries['word/document.xml'];
  if (!documentXmlBytes) {
    throw new Error('word/document.xml을 찾을 수 없습니다 — 올바른 docx 파일이 아닙니다.');
  }
  const documentXml = strFromU8(documentXmlBytes);
  const hasUnsupportedContent = /<w:drawing\b|<w:object\b|<w:pict\b/.test(documentXml);

  const parsed = parser.parse(documentXml) as XmlNode[];
  const documentEl = findFirst(parsed, 'w:document');
  const bodyEl = documentEl && findFirst(childrenOf(documentEl), 'w:body');
  const bodyChildren = bodyEl ? childrenOf(bodyEl) : [];

  const blocks: PMNode[] = [];
  const paragraphRun: Array<{ node: PMNode; isListItem: boolean }> = [];
  const flushParagraphRun = () => {
    blocks.push(...groupListItems(paragraphRun));
    paragraphRun.length = 0;
  };
  for (const child of bodyChildren) {
    const tag = tagOf(child);
    if (tag === 'w:p') {
      paragraphRun.push(paragraphFrom(child));
    } else if (tag === 'w:tbl') {
      flushParagraphRun();
      blocks.push(tableFrom(child));
    }
    // w:sectPr 등 그 외 body 직속 요소는 무시한다 (v1 미지원 영역)
  }
  flushParagraphRun();

  const content = blocks.length > 0 ? blocks : [docxSchema.nodes.paragraph.create()];
  const doc = (docxSchema.nodes.doc as NodeType).create(null, content);
  return { doc, hasUnsupportedContent };
}
