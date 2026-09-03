import { strFromU8, unzipSync } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { Node as PMNode, type NodeType } from 'prosemirror-model';
import {
  docxSchema,
  type CellBorderSide,
  type CellBorders,
  type CellPadding,
  type PageSetup,
  type RowHeight,
} from './schema.js';

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
 * w:r 하나의 자식들을 등장 순서대로 훑어 텍스트(w:t)와 줄바꿈(w:br)을 PM 노드로 만든다.
 * w:br을 무시하고 w:t만 이어붙이면 "첫 줄"과 "둘째 줄"이 구분 없이 "첫 줄둘째 줄"처럼
 * 뭉쳐버리므로, 각 w:t는 자기 위치 그대로 별개 text 노드로, w:br은 hard_break 노드로 만든다.
 * findAll이 반환하는 w:t는 "래퍼" 노드(예: { 'w:t': [...], ':@': {...} })라 텍스트는 그 안의
 * childrenOf(wt)를 다시 한 번 textOf로 펼쳐야 나온다.
 */
function runInlineNodesOf(runChildren: readonly XmlNode[], marks: RunMarks): PMNode[] {
  const nodes: PMNode[] = [];
  for (const child of runChildren) {
    const tag = tagOf(child);
    if (tag === 'w:t') {
      const node = textNodeWithMarks(textOf(childrenOf(child)), marks);
      if (node) nodes.push(node);
    } else if (tag === 'w:br') {
      nodes.push(docxSchema.nodes.hard_break.create());
    }
  }
  return nodes;
}

/** w:p의 자식을 순회하며 인라인(텍스트·줄바꿈) 콘텐츠를 뽑아낸다. 이미지(w:drawing)는 건너뛴다(v1 미지원). */
function inlineContentOf(pChildren: readonly XmlNode[]): PMNode[] {
  const nodes: PMNode[] = [];
  for (const run of findAll(pChildren, 'w:r')) {
    const runChildren = childrenOf(run);
    nodes.push(...runInlineNodesOf(runChildren, runMarksOf(runChildren)));
  }
  return nodes;
}

function paragraphFrom(pNode: XmlNode, defaultAlign: string | null): { node: PMNode; isListItem: boolean } {
  const pChildren = childrenOf(pNode);
  const pPr = findFirst(pChildren, 'w:pPr');
  const pPrChildren = pPr ? childrenOf(pPr) : [];
  const level = headingLevel(pPrChildren);
  const hasPStyle = findFirst(pPrChildren, 'w:pStyle') !== undefined;
  // 문단 자신의 w:jc가 없고, 다른 스타일도 참조하지 않으면(=암묵적으로 기본 스타일을 쓰면)
  // 그 기본 스타일 자체의 정렬(예: w:jc="both")을 물려받는다 — 제목(heading)은 별도 스타일
  // 캐스케이드를 아직 풀지 않으므로 이 폴백에서 제외한다(잘못 끼워맞추지 않기 위해).
  const align = alignOf(pPrChildren) ?? (level === null && !hasPStyle ? defaultAlign : null);
  const isListItem = findFirst(pPrChildren, 'w:numPr') !== undefined;
  const inline = inlineContentOf(pChildren);
  const node =
    level !== null
      ? docxSchema.nodes.heading.create({ level, align }, inline)
      : docxSchema.nodes.paragraph.create({ align }, inline);
  return { node, isListItem };
}

/** CSS px 기준 dxa(=twip, 1/20pt) 환산 — 96px = 1440dxa(1인치) */
const DXA_PER_PX = 15;

function dxaToPx(dxa: number): number {
  return Math.round(dxa / DXA_PER_PX);
}

/** w:tcPr/w:shd@w:fill — 셀 음영 색. "auto"나 6자리 hex가 아니면(무늬 채움 등) 음영 없음으로 취급한다. */
function backgroundOf(tcPrChildren: readonly XmlNode[]): string | null {
  const shd = findFirst(tcPrChildren, 'w:shd');
  if (!shd) return null;
  const fill = attrsOf(shd)['@_w:fill'];
  return fill && /^[0-9A-Fa-f]{6}$/.test(fill) ? `#${fill.toLowerCase()}` : null;
}

function tableCellFrom(
  tcNode: XmlNode,
  colspan: number,
  rowspan: number,
  colwidth: readonly number[] | null,
  borders: CellBorders,
  defaultAlign: string | null,
  cellPadding: CellPadding,
): PMNode {
  const tcChildren = childrenOf(tcNode);
  const paragraphs = findAll(tcChildren, 'w:p').map((p) => paragraphFrom(p, defaultAlign).node);
  const content = paragraphs.length > 0 ? paragraphs : [docxSchema.nodes.paragraph.create()];
  const tcPrChildren = tcPrChildrenOf(tcNode);
  const background = backgroundOf(tcPrChildren);
  const verticalAlign = verticalAlignOf(tcPrChildren);
  return docxSchema.nodes.table_cell.create(
    { colspan, rowspan, colwidth: colwidth ? [...colwidth] : null, background, borders, verticalAlign, cellPadding },
    content,
  );
}

function tcPrChildrenOf(tcNode: XmlNode): XmlNode[] {
  const tcPr = findFirst(childrenOf(tcNode), 'w:tcPr');
  return tcPr ? childrenOf(tcPr) : [];
}

interface CellMarginSides {
  readonly top: number | undefined;
  readonly right: number | undefined;
  readonly bottom: number | undefined;
  readonly left: number | undefined;
}

/** w:tblCellMar 또는 w:tcMar의 자식들에서 4변을 읽는다(dxa, 없는 변은 undefined). */
function cellMarginSidesOf(marChildren: readonly XmlNode[]): CellMarginSides {
  const sideOf = (tag: string): number | undefined => {
    const el = findFirst(marChildren, tag);
    if (!el) return undefined;
    const val = Number(attrsOf(el)['@_w:w']);
    return Number.isFinite(val) && val >= 0 ? val : undefined;
  };
  return { top: sideOf('w:top'), right: sideOf('w:right'), bottom: sideOf('w:bottom'), left: sideOf('w:left') };
}

/** Word 자체의 표준 기본 셀 여백(관례상의 "Normal Table" 기본값) — 위/아래 0, 좌/우 108dxa(0.075in). */
const STANDARD_CELL_MARGIN_DXA: CellMarginSides = { top: 0, right: 108, bottom: 0, left: 108 };

/** w:tblPr/w:tblCellMar — 이 표 자체가 지정한 기본 셀 여백. 없으면 null. */
function tableCellMarginOf(tblChildren: readonly XmlNode[]): CellMarginSides | null {
  const tblPr = findFirst(tblChildren, 'w:tblPr');
  const tblPrChildren = tblPr ? childrenOf(tblPr) : [];
  const tblCellMar = findFirst(tblPrChildren, 'w:tblCellMar');
  return tblCellMar ? cellMarginSidesOf(childrenOf(tblCellMar)) : null;
}

/**
 * "기본" 표 스타일(w:type="table" w:default="1") 자체의 w:tblCellMar — w:tblStyle을 지정하지
 * 않은 모든 표가 암묵적으로 상속받는다(문단의 "기본" 스타일과 같은 구조). 실제 문서에서
 * 이 스타일이 좌우 108dxa 같은 표준값을 명시하는 경우가 흔하다.
 */
function defaultTableStyleCellMarginOf(stylesChildren: readonly XmlNode[]): CellMarginSides | null {
  const style = findAll(stylesChildren, 'w:style').find((s) => {
    const attrs = attrsOf(s);
    return attrs['@_w:type'] === 'table' && attrs['@_w:default'] === '1';
  });
  const tblPr = style && findFirst(childrenOf(style), 'w:tblPr');
  const tblCellMar = tblPr && findFirst(childrenOf(tblPr), 'w:tblCellMar');
  return tblCellMar ? cellMarginSidesOf(childrenOf(tblCellMar)) : null;
}

/** w:tcPr/w:tcMar — 이 셀만의 여백 재정의(있는 변만 표 기본값을 덮어쓴다). */
function cellMarginOverrideOf(tcPrChildren: readonly XmlNode[]): CellMarginSides | null {
  const tcMar = findFirst(tcPrChildren, 'w:tcMar');
  return tcMar ? cellMarginSidesOf(childrenOf(tcMar)) : null;
}

/**
 * 표 자체 지정 > 기본 표 스타일 > Word 표준 기본값 순으로 표 여백을 정하고, 그 위에
 * 이 셀만의 재정의를 얹어 최종 4변 여백(px)을 확정한다. 여백은 테두리와 달리 "0"과
 * "지정 안 함"을 구분할 필요가 없어(0도 유효한 여백값) ?? 체인으로 충분하다.
 *
 * dxaToPx(정수로 반올림)를 쓰지 않고 소수점까지 그대로 둔다 — 이 여백이 딱 맞는 셀 안에서
 * 텍스트가 원본처럼 한 줄에 들어가느냐 두 줄로 줄바꿈되느냐를 가르는 값이라, 반올림으로
 * 잃는 0.x px가 곧 줄바꿈 여부를 뒤집을 수 있다(예: 108dxa는 정확히 7.2px, 7px로
 * 반올림하면 그만큼 텍스트 폭이 넓어져 원본과 다시 어긋난다).
 */
function resolveCellMargin(
  tableDefault: CellMarginSides | null,
  styleDefault: CellMarginSides | null,
  cellOverride: CellMarginSides | null,
): CellPadding {
  const pick = (side: keyof CellMarginSides): number =>
    (cellOverride?.[side] ?? tableDefault?.[side] ?? styleDefault?.[side] ?? STANDARD_CELL_MARGIN_DXA[side]!) /
    DXA_PER_PX;
  return { topPx: pick('top'), rightPx: pick('right'), bottomPx: pick('bottom'), leftPx: pick('left') };
}

/** w:tcPr/w:vAlign — 셀 안 내용의 세로 정렬. docx는 top/center/bottom, CSS vertical-align은 top/middle/bottom. */
function verticalAlignOf(tcPrChildren: readonly XmlNode[]): string | null {
  const vAlign = findFirst(tcPrChildren, 'w:vAlign');
  if (!vAlign) return null;
  const val = attrsOf(vAlign)['@_w:val'];
  if (val === 'center') return 'middle';
  if (val === 'bottom' || val === 'top') return val;
  return null;
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

/** w:tblGrid/w:gridCol@w:w 목록(dxa) — 표 전체의 열별 너비. 없으면 빈 배열. */
function gridColWidthsOf(tblChildren: readonly XmlNode[]): number[] {
  const tblGrid = findFirst(tblChildren, 'w:tblGrid');
  if (!tblGrid) return [];
  return findAll(childrenOf(tblGrid), 'w:gridCol').map((col) => Number(attrsOf(col)['@_w:w']) || 0);
}

/** docx 테두리 스타일(w:val) → CSS border-style. 알 수 없는 값은 실선으로 대체한다. */
const OOXML_BORDER_STYLE_TO_CSS: Record<string, string> = {
  double: 'double',
  dashed: 'dashed',
  dashSmallGap: 'dashed',
  dotted: 'dotted',
};

/** 한 변(w:top/w:bottom/w:left/w:right/w:insideH/w:insideV) 하나를 CellBorderSide로 변환한다. */
/**
 * undefined(이 변 요소 자체가 XML에 없음 — 상위 기본값을 따라야 한다)와 null(요소는 있지만
 * "테두리 없음"으로 명시됨 — 상위 기본값을 따르면 안 되고 실제로 안 그려야 한다)을 구분해야
 * w:tcBorders가 특정 변만 "nil"로 재정의했을 때 표 기본 테두리로 잘못 되돌아가지 않는다.
 */
function borderSideOf(el: XmlNode | undefined): CellBorderSide | null | undefined {
  if (!el) return undefined;
  const attrs = attrsOf(el);
  const val = attrs['@_w:val'];
  if (!val || val === 'none' || val === 'nil') return null;
  const sz = Number(attrs['@_w:sz']);
  // w:sz는 1/8pt 단위. px = pt * (96/72) = pt * 4/3, pt = sz/8 → px = sz/6.
  // 0.5px 단위로 반올림하고 최소 1px은 되게 한다 — Word의 표준 얇은 테두리(sz=4, 0.5pt≈0.67px)가
  // 그대로면 브라우저 화면에서 거의 안 보일 만큼 얇아지기 때문이다.
  const widthPx = Number.isFinite(sz) && sz > 0 ? Math.max(1, Math.round((sz / 6) * 2) / 2) : 1;
  const colorRaw = attrs['@_w:color'];
  const color = colorRaw && /^[0-9A-Fa-f]{6}$/.test(colorRaw) ? `#${colorRaw.toLowerCase()}` : '#000000';
  return { style: OOXML_BORDER_STYLE_TO_CSS[val] ?? 'solid', widthPx, color };
}

interface BorderSet {
  readonly top: CellBorderSide | null | undefined;
  readonly right: CellBorderSide | null | undefined;
  readonly bottom: CellBorderSide | null | undefined;
  readonly left: CellBorderSide | null | undefined;
  readonly insideH: CellBorderSide | null | undefined;
  readonly insideV: CellBorderSide | null | undefined;
}

const EMPTY_BORDER_SET: BorderSet = { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };

/** w:tblBorders 또는 w:tcBorders의 자식들에서 6개 변을 모두 읽는다(없는 변은 null). */
function borderSetOf(borderChildren: readonly XmlNode[]): BorderSet {
  return {
    top: borderSideOf(findFirst(borderChildren, 'w:top')),
    right: borderSideOf(findFirst(borderChildren, 'w:right') ?? findFirst(borderChildren, 'w:end')),
    bottom: borderSideOf(findFirst(borderChildren, 'w:bottom')),
    left: borderSideOf(findFirst(borderChildren, 'w:left') ?? findFirst(borderChildren, 'w:start')),
    insideH: borderSideOf(findFirst(borderChildren, 'w:insideH')),
    insideV: borderSideOf(findFirst(borderChildren, 'w:insideV')),
  };
}

/** w:tblPr/w:tblBorders — 표 전체에 적용되는 기본 테두리(개별 셀이 따로 지정하지 않으면 이걸 쓴다). */
function tableBorderDefaultsOf(tblChildren: readonly XmlNode[]): BorderSet {
  const tblPr = findFirst(tblChildren, 'w:tblPr');
  const tblPrChildren = tblPr ? childrenOf(tblPr) : [];
  const tblBorders = findFirst(tblPrChildren, 'w:tblBorders');
  return tblBorders ? borderSetOf(childrenOf(tblBorders)) : EMPTY_BORDER_SET;
}

/** w:tcPr/w:tcBorders — 이 셀만의 테두리 재정의(있는 변만 표 기본값을 덮어쓴다). */
function cellBorderOverrideOf(tcPrChildren: readonly XmlNode[]): BorderSet | null {
  const tcBorders = findFirst(tcPrChildren, 'w:tcBorders');
  return tcBorders ? borderSetOf(childrenOf(tcBorders)) : null;
}

/**
 * 표 기본 테두리(바깥 변 top/right/bottom/left, 안쪽 grid선 insideH/insideV)와 이 셀만의
 * 재정의를 합쳐, 이 셀이 화면에 그려야 할 4변을 확정한다. 셀이 표의 가장자리에 있으면
 * 바깥 변(top/right/bottom/left)을, 안쪽에 있으면 insideH/insideV를 기본값으로 쓴다.
 * (Word의 실제 인접 셀 간 우선순위 비교 알고리즘까지는 재현하지 않는 v1 단순화 — 이 셀
 * 재정의가 있으면 무조건 그걸 쓴다.)
 */
function resolveCellBorders(
  tableDefault: BorderSet,
  cellOverride: BorderSet | null,
  pos: { row: number; col: number; colspan: number; endRow: number; totalRows: number; totalCols: number },
): CellBorders {
  const isFirstRow = pos.row === 0;
  const isLastRow = pos.endRow >= pos.totalRows - 1;
  const isFirstCol = pos.col === 0;
  const isLastCol = pos.col + pos.colspan >= pos.totalCols;
  // override가 undefined(그 변 자체를 안 건드림)일 때만 기본값으로 넘어간다 — null(명시적
  // "테두리 없음")은 그대로 존중해야 하므로 ?? 를 쓰면 안 된다(null도 "값 없음"으로 취급해버림).
  const pick = (
    overrideVal: CellBorderSide | null | undefined,
    fallback: CellBorderSide | null | undefined,
  ): CellBorderSide | null => (overrideVal !== undefined ? overrideVal : fallback ?? null);
  return {
    top: pick(cellOverride?.top, isFirstRow ? tableDefault.top : tableDefault.insideH),
    right: pick(cellOverride?.right, isLastCol ? tableDefault.right : tableDefault.insideV),
    bottom: pick(cellOverride?.bottom, isLastRow ? tableDefault.bottom : tableDefault.insideH),
    left: pick(cellOverride?.left, isFirstCol ? tableDefault.left : tableDefault.insideV),
  };
}

interface CellPlacement {
  readonly tc: XmlNode;
  readonly row: number;
  readonly col: number;
  readonly colspan: number;
  rowspan: number;
}

/**
 * OOXML의 표 병합(w:gridSpan=가로, w:vMerge=세로)을 ProseMirror의 colspan/rowspan 모델로
 * 옮긴다. ProseMirror(prosemirror-tables)는 세로 병합을 "이어지는 칸"을 아예 만들지 않고
 * 시작 칸의 rowspan만 늘리는 방식으로 표현하므로, w:vMerge가 "continue"인 <w:tc>는 새 셀을
 * 만들지 않고 같은 열의 시작 칸(originByCol) rowspan만 1 늘린 뒤 건너뛴다.
 */
/** w:tblPr/w:jc — 표 자체가 페이지 안에서 왼쪽/가운데/오른쪽 중 어디 놓이는지(셀 안 문단 정렬과 무관). */
function tableAlignOf(tblChildren: readonly XmlNode[]): string | null {
  const tblPr = findFirst(tblChildren, 'w:tblPr');
  const tblPrChildren = tblPr ? childrenOf(tblPr) : [];
  const jc = findFirst(tblPrChildren, 'w:jc');
  if (!jc) return null;
  const val = attrsOf(jc)['@_w:val'];
  return val === 'center' || val === 'right' ? val : null;
}

/** w:trPr/w:trHeight — 행 높이. hRule="auto"(또는 요소 자체가 없음)는 "지정 없음"과 동일하게 null. */
function rowHeightOf(trChildren: readonly XmlNode[]): RowHeight | null {
  const trPr = findFirst(trChildren, 'w:trPr');
  const trPrChildren = trPr ? childrenOf(trPr) : [];
  const trHeight = findFirst(trPrChildren, 'w:trHeight');
  if (!trHeight) return null;
  const attrs = attrsOf(trHeight);
  const val = Number(attrs['@_w:val']);
  if (!Number.isFinite(val) || val <= 0) return null;
  const hRule = attrs['@_w:hRule'];
  if (hRule === 'auto') return null;
  return { rule: hRule === 'exact' ? 'exact' : 'atLeast', px: dxaToPx(val) };
}

/**
 * w:tblGrid 열 폭 합이 페이지 내용 폭(contentWidthPx)보다 넓으면, Word가 실제로 그러듯
 * (AutoFit) 비율을 유지한 채 내용 폭에 맞춰 축소한다. contentWidthPx를 모르거나(sectPr 없음)
 * 이미 내용 폭 이내면 그대로 둔다 — 원본보다 좁은 표를 억지로 늘리지는 않는다(과거에 고친
 * width:100% 강제 확장 버그의 재발 방지).
 */
function fitGridColWidthsPx(gridColWidthsPx: readonly number[], contentWidthPx: number | null): number[] {
  const total = gridColWidthsPx.reduce((sum, w) => sum + w, 0);
  if (contentWidthPx === null || total <= contentWidthPx || total <= 0) return [...gridColWidthsPx];
  const scale = contentWidthPx / total;
  return gridColWidthsPx.map((w) => Math.max(1, Math.round(w * scale)));
}

function tableFrom(
  tblNode: XmlNode,
  contentWidthPx: number | null,
  defaultAlign: string | null,
  styleTableCellMargin: CellMarginSides | null,
): PMNode {
  const tblChildren = childrenOf(tblNode);
  const trNodes = findAll(tblChildren, 'w:tr');
  const gridColWidths = fitGridColWidthsPx(gridColWidthsOf(tblChildren).map(dxaToPx), contentWidthPx);
  const tableBorderDefaults = tableBorderDefaultsOf(tblChildren);
  const tableCellMarginDefault = tableCellMarginOf(tblChildren);
  const tableAlign = tableAlignOf(tblChildren);
  const rowHeights: Array<RowHeight | null> = [];

  const rowPlacements: CellPlacement[][] = [];
  const originByCol = new Map<number, CellPlacement>();

  trNodes.forEach((tr, row) => {
    rowHeights.push(rowHeightOf(childrenOf(tr)));
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
      const placement: CellPlacement = { tc, row, col, colspan: gridSpan, rowspan: 1 };
      placements.push(placement);
      originByCol.set(col, placement);
      col += gridSpan;
    }
    rowPlacements.push(placements);
  });

  const totalRows = rowPlacements.length;
  const totalCols = Math.max(
    0,
    ...rowPlacements.map((placements) => placements.reduce((sum, p) => sum + p.colspan, 0)),
  );

  const rows = rowPlacements.map((placements, row) =>
    docxSchema.nodes.table_row.create(
      { height: rowHeights[row] },
      placements.map((p) => {
        const widths = gridColWidths.slice(p.col, p.col + p.colspan);
        const colwidth = widths.length === p.colspan ? widths : null;
        const tcPrChildren = tcPrChildrenOf(p.tc);
        const cellOverride = cellBorderOverrideOf(tcPrChildren);
        const borders = resolveCellBorders(tableBorderDefaults, cellOverride, {
          row: p.row,
          col: p.col,
          colspan: p.colspan,
          endRow: p.row + p.rowspan - 1,
          totalRows,
          totalCols,
        });
        const cellPadding = resolveCellMargin(
          tableCellMarginDefault,
          styleTableCellMargin,
          cellMarginOverrideOf(tcPrChildren),
        );
        return tableCellFrom(p.tc, p.colspan, p.rowspan, colwidth, borders, defaultAlign, cellPadding);
      }),
    ),
  );
  return docxSchema.nodes.table.create({ align: tableAlign }, rows);
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
  /** 페이지 크기·여백(w:sectPr/w:pgSz·w:pgMar) — 화면에 실제 Word 페이지 비율로 그리기 위함. 없으면 null. */
  readonly pageSetup: PageSetup | null;
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

/** w:sectPr(파싱된 트리)에서 페이지 크기·여백을 읽는다. w:pgSz가 없으면(비정상 문서) null.
 * 여백이 명시 안 된 변은 Word 기본값(1인치=1440dxa)으로 채운다. */
function pageSetupOf(sectPr: XmlNode | undefined): PageSetup | null {
  if (!sectPr) return null;
  const sectPrChildren = childrenOf(sectPr);
  const pgSz = findFirst(sectPrChildren, 'w:pgSz');
  if (!pgSz) return null;
  const pgSzAttrs = attrsOf(pgSz);
  const w = Number(pgSzAttrs['@_w:w']);
  const h = Number(pgSzAttrs['@_w:h']);
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) return null;
  const pgMar = findFirst(sectPrChildren, 'w:pgMar');
  const pgMarAttrs = pgMar ? attrsOf(pgMar) : {};
  const marginOf = (key: string): number => {
    const v = Number(pgMarAttrs[key]);
    return dxaToPx(Number.isFinite(v) && v >= 0 ? v : 1440);
  };
  return {
    widthPx: dxaToPx(w),
    heightPx: dxaToPx(h),
    marginTopPx: marginOf('@_w:top'),
    marginBottomPx: marginOf('@_w:bottom'),
    marginLeftPx: marginOf('@_w:left'),
    marginRightPx: marginOf('@_w:right'),
    orientation: pgSzAttrs['@_w:orient'] === 'landscape' ? 'landscape' : 'portrait',
  };
}

/**
 * word/document.xml 원본 문자열에서 (마지막) w:sectPr 요소를 그대로 문자열로 뽑아낸다 — 우리가
 * 모델링하지 않는 세부 설정(헤더/푸터 참조, 단 나누기 등)까지 손실 없이 내보내기 위해, 파싱 결과를
 * 재조립하는 대신 원본 텍스트를 그대로 들고 있다가 export.ts가 그대로 되돌려 쓴다.
 */
export function extractSectPrXml(documentXml: string): string | null {
  const matches = documentXml.match(/<w:sectPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:sectPr>)/g);
  return matches && matches.length > 0 ? matches[matches.length - 1] : null;
}

/** Word가 실제로 쓰는 기본 폰트 크기가 없을 때의 대체값(10pt) — Word 새 문서의 기본값도 이와 같다. */
const FALLBACK_DEFAULT_FONT_SIZE_PT = 10;

/** w:docDefaults/w:rPrDefault/w:rPr/w:sz — 직접 서식도, 스타일도 지정하지 않은 실행문이 실제로
 * 쓰는 폰트 크기(half-point). 없으면 10pt로 대체한다. */
function defaultFontSizePtOf(docDefaults: XmlNode | undefined): number {
  if (!docDefaults) return FALLBACK_DEFAULT_FONT_SIZE_PT;
  const rPrDefault = findFirst(childrenOf(docDefaults), 'w:rPrDefault');
  const rPr = rPrDefault && findFirst(childrenOf(rPrDefault), 'w:rPr');
  const sz = rPr && findFirst(childrenOf(rPr), 'w:sz');
  const val = sz ? Number(attrsOf(sz)['@_w:val']) : NaN;
  return Number.isFinite(val) && val > 0 ? val / 2 : FALLBACK_DEFAULT_FONT_SIZE_PT;
}

/** w:docDefaults/w:rPrDefault/w:rPr/w:rFonts 요소 자체(있으면) — docDefaultsRFontsOf에서 쓴다. */
function docDefaultsRFontsOf(stylesChildren: readonly XmlNode[]): XmlNode | undefined {
  const docDefaults = findFirst(stylesChildren, 'w:docDefaults');
  const rPrDefault = docDefaults && findFirst(childrenOf(docDefaults), 'w:rPrDefault');
  const rPr = rPrDefault && findFirst(childrenOf(rPrDefault), 'w:rPr');
  return rPr && findFirst(childrenOf(rPr), 'w:rFonts');
}

/**
 * "기본" 문단 스타일(w:type="paragraph" w:default="1") — pStyle을 지정하지 않은 모든 문단이
 * 암묵적으로 쓰는 스타일이다. 실제 문서에서 이 스타일 자체가 docDefaults와 다른 폰트 크기·
 * 서체·정렬을 직접 지정하는 경우가 흔해서(예: docDefaults는 테마 맑은 고딕을 가리키지만
 * 기본 스타일 자체는 리터럴로 굴림을 지정), 그 경우 실제로 적용되는 값은 이 스타일 쪽이다.
 */
function findDefaultParagraphStyle(stylesChildren: readonly XmlNode[]): XmlNode | undefined {
  return findAll(stylesChildren, 'w:style').find((style) => {
    const attrs = attrsOf(style);
    return attrs['@_w:type'] === 'paragraph' && attrs['@_w:default'] === '1';
  });
}

interface DefaultParagraphStyleInfo {
  /** 스타일 자체가 지정한 폰트 크기(pt). 없으면 null → docDefaults로 폴백해야 한다. */
  readonly sizePt: number | null;
  /** 스타일 자체의 w:rFonts 요소. 없으면 undefined → docDefaults로 폴백해야 한다. */
  readonly rFonts: XmlNode | undefined;
  /** 스타일 자체의 문단 정렬(ProseMirror align 값). 없으면 null. */
  readonly align: string | null;
}

function defaultParagraphStyleInfoOf(stylesChildren: readonly XmlNode[]): DefaultParagraphStyleInfo {
  const style = findDefaultParagraphStyle(stylesChildren);
  const styleChildren = style ? childrenOf(style) : [];
  const rPr = findFirst(styleChildren, 'w:rPr');
  const rPrChildren = rPr ? childrenOf(rPr) : [];
  const sz = findFirst(rPrChildren, 'w:sz');
  const szVal = sz ? Number(attrsOf(sz)['@_w:val']) : NaN;
  const pPr = findFirst(styleChildren, 'w:pPr');
  return {
    sizePt: Number.isFinite(szVal) && szVal > 0 ? szVal / 2 : null,
    rFonts: findFirst(rPrChildren, 'w:rFonts'),
    align: alignOf(pPr ? childrenOf(pPr) : []),
  };
}

function stylesChildrenOf(stylesXml: string): XmlNode[] {
  const parsed = parser.parse(stylesXml) as XmlNode[];
  const stylesEl = findFirst(parsed, 'w:styles');
  return stylesEl ? childrenOf(stylesEl) : [];
}

/**
 * word/styles.xml의 문서 기본 폰트 크기를 뽑아낸다 — 직접 서식(w:rPr/w:sz)도 스타일도 없는
 * 실행문은 이 값을 상속받아 렌더링해야 브라우저 기본값(16px)이 아니라 Word가 실제로 보여주는
 * 크기와 같아진다(client/editor.ts가 .ProseMirror 기본 font-size로 적용). 기본 문단 스타일
 * 자체가 크기를 지정하면 그 값이 docDefaults보다 우선한다(실제 Word의 상속 순서).
 */
export function getDocxDefaultFontSizePt(docxBytes: Uint8Array): number {
  const entries = unzipSync(docxBytes);
  const stylesXmlBytes = entries['word/styles.xml'];
  if (!stylesXmlBytes) return FALLBACK_DEFAULT_FONT_SIZE_PT;
  const stylesChildren = stylesChildrenOf(strFromU8(stylesXmlBytes));
  const styleSizePt = defaultParagraphStyleInfoOf(stylesChildren).sizePt;
  return styleSizePt ?? defaultFontSizePtOf(findFirst(stylesChildren, 'w:docDefaults'));
}

/** a:fontScheme/a:minorFont(또는 a:majorFont) 안에서 라틴 기본 서체와 한글(Hang) 서체를 뽑는다. */
function themeTypefaceOf(themeXml: string, slot: 'minorFont' | 'majorFont'): { latin: string | null; hang: string | null } {
  const slotMatch = new RegExp(`<a:${slot}>([\\s\\S]*?)</a:${slot}>`).exec(themeXml);
  if (!slotMatch) return { latin: null, hang: null };
  const body = slotMatch[1];
  const latin = /<a:latin typeface="([^"]*)"/.exec(body)?.[1];
  const hang = /<a:font script="Hang" typeface="([^"]*)"/.exec(body)?.[1];
  return { latin: latin || null, hang: hang || null };
}

/**
 * w:rFonts의 한 서체 지정(예: w:eastAsia/w:eastAsiaTheme, w:ascii/w:asciiTheme)을 실제 서체
 * 이름으로 푼다. 직접 이름(literal)이 있으면 그걸 그대로 쓰고, 테마 참조(예: "majorEastAsia")만
 * 있으면 word/theme/theme1.xml의 해당 슬롯(major/minor)에서 script에 맞는 서체(한글은 Hang,
 * 라틴은 latin)를 찾는다.
 */
function resolveFontRef(
  literal: string | undefined,
  themeRef: string | undefined,
  themeXml: string | undefined,
  script: 'hang' | 'latin',
): string | null {
  if (literal) return literal;
  if (themeRef && themeXml) {
    const slot = /major/i.test(themeRef) ? 'majorFont' : 'minorFont';
    const typefaces = themeTypefaceOf(themeXml, slot);
    return typefaces[script];
  }
  return null;
}

/** w:rFonts 요소 하나를 CSS font-family 값으로 푼다(테마 참조 포함). 없으면 null. */
function resolveFontFamilyFromRFonts(rFonts: XmlNode | undefined, themeXml: string | undefined): string | null {
  if (!rFonts) return null;
  const attrs = attrsOf(rFonts);
  const eastAsia = resolveFontRef(attrs['@_w:eastAsia'], attrs['@_w:eastAsiaTheme'], themeXml, 'hang');
  const ascii = resolveFontRef(attrs['@_w:ascii'], attrs['@_w:asciiTheme'], themeXml, 'latin');
  const families = [eastAsia, ascii].filter(
    (family, index, all): family is string => family !== null && all.indexOf(family) === index,
  );
  return families.length > 0 ? families.map((family) => `"${family}"`).join(', ') : null;
}

/**
 * word/styles.xml의 문서 기본 서체를 CSS font-family 값으로 뽑아낸다 — 직접 서식이 없는
 * 실행문은 이 서체로 보여야 브라우저의 기본 UI 폰트가 아니라 Word가 실제로 쓰는 서체와
 * 같아진다. 기본 문단 스타일 자체가 서체를 지정하면(흔히 리터럴로) 그 값이 docDefaults(흔히
 * 테마 참조)보다 우선한다 — 실제 Word의 상속 순서와 같다. 지정된 서체가 전혀 없으면 null
 * (기존 CSS 폴백 유지).
 */
export function getDocxDefaultFontFamily(docxBytes: Uint8Array): string | null {
  const entries = unzipSync(docxBytes);
  const stylesXmlBytes = entries['word/styles.xml'];
  if (!stylesXmlBytes) return null;
  const stylesChildren = stylesChildrenOf(strFromU8(stylesXmlBytes));
  const themeXmlBytes = entries['word/theme/theme1.xml'];
  const themeXml = themeXmlBytes ? strFromU8(themeXmlBytes) : undefined;
  const styleRFonts = defaultParagraphStyleInfoOf(stylesChildren).rFonts;
  return (
    resolveFontFamilyFromRFonts(styleRFonts, themeXml) ??
    resolveFontFamilyFromRFonts(docDefaultsRFontsOf(stylesChildren), themeXml)
  );
}

/**
 * 페이지 크기·여백만 가볍게 뽑아낸다(표·문단 전체를 파싱하는 importDocx보다 훨씬 싸다) —
 * 뷰어가 "이 문서를 실제 Word 페이지 비율로 얼마나 넓게 그릴지" 계산하는 용도(src/routes/viewer.ts).
 */
export function getDocxPageSetup(docxBytes: Uint8Array): PageSetup | null {
  const entries = unzipSync(docxBytes);
  const documentXmlBytes = entries['word/document.xml'];
  if (!documentXmlBytes) return null;
  const sectPrXml = extractSectPrXml(strFromU8(documentXmlBytes));
  if (!sectPrXml) return null;
  const parsed = parser.parse(sectPrXml) as XmlNode[];
  const sectPr = findFirst(parsed, 'w:sectPr');
  return pageSetupOf(sectPr);
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
  const pageSetup = pageSetupOf(findFirst(bodyChildren, 'w:sectPr'));
  const contentWidthPx = pageSetup ? pageSetup.widthPx - pageSetup.marginLeftPx - pageSetup.marginRightPx : null;
  // pStyle 없이 암묵적으로 "기본" 문단/표 스타일을 쓰는 문단·표(대부분의 실제 문서가 이렇다)가
  // 실제로 상속받는 정렬·셀 여백 — w:jc가 없다고 곧 left, w:tblCellMar가 없다고 곧 우리 CSS
  // 기본값(0.4rem 0.6rem)인 게 아니라, "기본" 스타일 자체가 다른 값을 지정하고 있을 수 있다
  // (paragraphFrom/tableFrom 참고).
  const stylesXmlBytes = entries['word/styles.xml'];
  const stylesChildren = stylesXmlBytes ? stylesChildrenOf(strFromU8(stylesXmlBytes)) : [];
  const defaultAlign = stylesXmlBytes ? defaultParagraphStyleInfoOf(stylesChildren).align : null;
  const styleTableCellMargin = stylesXmlBytes ? defaultTableStyleCellMarginOf(stylesChildren) : null;

  const blocks: PMNode[] = [];
  const paragraphRun: Array<{ node: PMNode; isListItem: boolean }> = [];
  const flushParagraphRun = () => {
    blocks.push(...groupListItems(paragraphRun));
    paragraphRun.length = 0;
  };
  for (const child of bodyChildren) {
    const tag = tagOf(child);
    if (tag === 'w:p') {
      paragraphRun.push(paragraphFrom(child, defaultAlign));
    } else if (tag === 'w:tbl') {
      flushParagraphRun();
      blocks.push(tableFrom(child, contentWidthPx, defaultAlign, styleTableCellMargin));
    }
    // w:sectPr 등 그 외 body 직속 요소는 무시한다 (v1 미지원 영역)
  }
  flushParagraphRun();

  const content = blocks.length > 0 ? blocks : [docxSchema.nodes.paragraph.create()];
  const doc = (docxSchema.nodes.doc as NodeType).create(null, content);
  return { doc, hasUnsupportedContent, pageSetup };
}
