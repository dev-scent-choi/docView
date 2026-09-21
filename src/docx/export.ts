import { strFromU8, strToU8, zipSync } from 'fflate';
import { guardedUnzipSync } from '../zipGuard.js';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { extractSectPrXml } from './import.js';
import type { CellBorderSide, CellBorders, CellPadding, RowHeight } from './schema.js';

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function renderRun(text: string, marks: readonly Mark[]): string {
  const names = new Set(marks.map((mark) => mark.type.name));
  const rPrParts: string[] = [];
  if (names.has('strong')) rPrParts.push('<w:b/>');
  if (names.has('em')) rPrParts.push('<w:i/>');
  if (names.has('underline')) rPrParts.push('<w:u w:val="single"/>');
  if (names.has('strike')) rPrParts.push('<w:strike/>');
  const colorMark = marks.find((mark) => mark.type.name === 'textColor');
  if (colorMark) {
    const hex = String(colorMark.attrs.color).replace('#', '').toUpperCase();
    rPrParts.push(`<w:color w:val="${hex}"/>`);
  }
  const sizeMark = marks.find((mark) => mark.type.name === 'fontSize');
  if (sizeMark) {
    const halfPoints = Math.round(Number(sizeMark.attrs.size) * 2);
    rPrParts.push(`<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`);
  }
  const rPr = rPrParts.length > 0 ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

const ALIGN_TO_JC: Record<string, string> = { center: 'center', right: 'right', justify: 'both' };

function pPrXml(extra: string, align: unknown): string {
  const jc = typeof align === 'string' ? ALIGN_TO_JC[align] : undefined;
  const parts = [extra, jc ? `<w:jc w:val="${jc}"/>` : ''].filter(Boolean).join('');
  return parts ? `<w:pPr>${parts}</w:pPr>` : '';
}

function renderInline(node: PMNode): string {
  let xml = '';
  node.forEach((child) => {
    if (child.isText) {
      xml += renderRun(child.text ?? '', child.marks);
    } else if (child.type.name === 'hard_break') {
      xml += '<w:r><w:br/></w:r>';
    }
  });
  return xml;
}

function renderParagraph(node: PMNode, prefix = ''): string {
  const prefixRun = prefix ? renderRun(prefix, []) : '';
  const pPr = pPrXml('', node.attrs.align);
  return `<w:p>${pPr}${prefixRun}${renderInline(node)}</w:p>`;
}

function renderHeading(node: PMNode): string {
  const level = node.attrs.level as number;
  const pPr = pPrXml(`<w:pStyle w:val="Heading${level}"/>`, node.attrs.align);
  return `<w:p>${pPr}${renderInline(node)}</w:p>`;
}

/**
 * 목록은 word/numbering.xml 연동 없이(v1 스코프 단순화) "• "/"1. " 접두사가 붙은
 * 일반 문단으로 평탄화해 내보낸다 — Word에서 다시 열어도 텍스트는 그대로 보이지만
 * 실제 목록 서식(자동 번호 매기기 등)은 아니다.
 */
function renderList(node: PMNode, ordered: boolean): string {
  let xml = '';
  let index = 1;
  node.forEach((listItem) => {
    const prefix = ordered ? `${index}. ` : '• ';
    listItem.forEach((child) => {
      xml += child.type.name === 'paragraph' ? renderParagraph(child, prefix) : renderBlock(child);
    });
    index += 1;
  });
  return xml;
}

/** 셀에 colwidth(px)가 없을 때(예: 편집기에서 새로 만든 표)의 기본 셀 너비 */
const CELL_WIDTH_DXA = 2000;
/** CSS px 기준 dxa(=twip, 1/20pt) 환산 — 96px = 1440dxa(1인치). src/docx/import.ts와 대칭. */
const DXA_PER_PX = 15;

function pxToDxa(px: number): number {
  return Math.round(px * DXA_PER_PX);
}

/** 편집기에서 새로 만든 표처럼 셀에 borders 속성이 아예 없을 때만 쓰이는 표 기본 테두리 —
 * import.ts에서 가져온 셀은 항상 자기 borders를 명시적으로 갖고 있어 이 기본값을 안 쓴다. */
const BORDER = '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>';

const CSS_BORDER_STYLE_TO_OOXML: Record<string, string> = {
  solid: 'single',
  double: 'double',
  dashed: 'dashed',
  dotted: 'dotted',
};

/** CellBorderSide 한 변을 w:top/w:left/w:bottom/w:right 중 하나로 직렬화한다. null이면 "테두리 없음"(w:val="nil"). */
function renderBorderSideXml(tag: string, side: CellBorderSide | null): string {
  if (!side) return `<w:${tag} w:val="nil"/>`;
  const val = CSS_BORDER_STYLE_TO_OOXML[side.style] ?? 'single';
  // widthPx = sz/6 (import.ts와 대칭) → sz = widthPx*6, 최소 2(1/4pt)는 되게 한다.
  const sz = Math.max(2, Math.round(side.widthPx * 6));
  const color = side.color.replace('#', '').toUpperCase();
  return `<w:${tag} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
}

/** 셀의 borders 속성(4변 모두 확정된 값)을 <w:tcBorders>로 직렬화한다. 속성이 없으면 빈 문자열(표 기본 테두리를 그대로 따른다). */
function renderCellBordersXml(borders: unknown): string {
  if (!borders || typeof borders !== 'object') return '';
  const b = borders as CellBorders;
  return (
    '<w:tcBorders>' +
    renderBorderSideXml('top', b.top) +
    renderBorderSideXml('left', b.left) +
    renderBorderSideXml('bottom', b.bottom) +
    renderBorderSideXml('right', b.right) +
    '</w:tcBorders>'
  );
}

/** 셀의 cellPadding 속성(4변 모두 확정된 px 값)을 <w:tcMar>로 직렬화한다. 없으면 빈 문자열(표/스타일 기본 여백을 그대로 따른다). */
function renderCellMarginXml(padding: unknown): string {
  if (!padding || typeof padding !== 'object') return '';
  const p = padding as CellPadding;
  const side = (tag: string, px: number) => `<w:${tag} w:w="${pxToDxa(px)}" w:type="dxa"/>`;
  return `<w:tcMar>${side('top', p.topPx)}${side('left', p.leftPx)}${side('bottom', p.bottomPx)}${side('right', p.rightPx)}</w:tcMar>`;
}

/** 세로 병합 중(pending) 열의 상태 — 몇 행 더 w:vMerge 이어짐 칸을 채워야 하는지, 그 칸의 가로폭(colspan)·너비(dxa) */
interface PendingVMerge {
  remaining: number;
  readonly colspan: number;
  readonly widthDxa: number;
}

/**
 * ProseMirror 표 한 행을 docx 행으로 직렬화한다. prosemirror-tables는 세로 병합을 "이어지는
 * 칸은 그 행의 셀 목록에 아예 없다"는 방식으로 표현하므로(src/docx/import.ts의 tableFrom과
 * 동일한 모델), 이 행의 실제 셀들 사이사이에 이전 행에서 시작된 세로 병합(pendingVMerge)이
 * 남아 있는 열이 있으면 그 자리에 w:vMerge 이어짐 칸을 끼워 넣어 열 정렬을 맞춘다. 실제 셀과
 * pendingVMerge 칸은 열 위치가 절대 겹치지 않으므로(겹치면 애초에 이 행에 그 열의 실제 셀이
 * 없었을 것이다) 열 번호를 따로 계산할 필요 없이 앞에서부터 순서대로 배치하면 된다.
 */
/** table_row의 height 속성(atLeast/exact, px)을 <w:trPr><w:trHeight .../></w:trPr>로 직렬화한다. 없으면 빈 문자열(자동 높이). */
function renderRowHeightXml(height: unknown): string {
  if (!height || typeof height !== 'object') return '';
  const h = height as RowHeight;
  const hRule = h.rule === 'exact' ? 'exact' : 'atLeast';
  return `<w:trPr><w:trHeight w:val="${pxToDxa(h.px)}" w:hRule="${hRule}"/></w:trPr>`;
}

function renderTableRow(
  row: PMNode,
  pendingVMerge: Map<number, PendingVMerge>,
  columnWidthsDxa: Map<number, number>,
): string {
  const cells: PMNode[] = [];
  row.forEach((cell) => cells.push(cell));

  const trPrXml = renderRowHeightXml(row.attrs.height);
  let cellsXml = '';
  let col = 0;
  let cellIndex = 0;

  while (cellIndex < cells.length || pendingVMerge.has(col)) {
    const pending = pendingVMerge.get(col);
    if (pending) {
      const gridSpanXml = pending.colspan > 1 ? `<w:gridSpan w:val="${pending.colspan}"/>` : '';
      cellsXml +=
        `<w:tc><w:tcPr><w:tcW w:w="${pending.widthDxa}" w:type="dxa"/>` +
        `${gridSpanXml}<w:vMerge/></w:tcPr><w:p/></w:tc>`;
      pending.remaining -= 1;
      if (pending.remaining === 0) pendingVMerge.delete(col);
      col += pending.colspan;
      continue;
    }

    const cell = cells[cellIndex];
    cellIndex += 1;
    const colspan = Number(cell.attrs.colspan) || 1;
    const rowspan = Number(cell.attrs.rowspan) || 1;
    const colwidthPx = Array.isArray(cell.attrs.colwidth) ? (cell.attrs.colwidth as number[]) : null;
    // 열별 너비(colwidth)가 있으면 그대로 쓰고(열마다 tblGrid에 반영), 없으면(예: 편집기에서
    // 새로 만든 셀) 기본 너비를 colspan칸에 균등 배분한다.
    const perColumnDxa = colwidthPx
      ? colwidthPx.map(pxToDxa)
      : Array.from({ length: colspan }, () => CELL_WIDTH_DXA);
    const totalWidthDxa = perColumnDxa.reduce((sum, w) => sum + w, 0);
    perColumnDxa.forEach((w, i) => columnWidthsDxa.set(col + i, w));

    let cellBody = '';
    cell.forEach((block) => {
      cellBody += renderBlock(block);
    });
    const gridSpanXml = colspan > 1 ? `<w:gridSpan w:val="${colspan}"/>` : '';
    const vMergeXml = rowspan > 1 ? '<w:vMerge w:val="restart"/>' : '';
    const background = typeof cell.attrs.background === 'string' ? (cell.attrs.background as string) : null;
    const shdXml = background
      ? `<w:shd w:val="clear" w:color="auto" w:fill="${background.replace('#', '').toUpperCase()}"/>`
      : '';
    const bordersXml = renderCellBordersXml(cell.attrs.borders);
    const verticalAlign = cell.attrs.verticalAlign;
    // docx는 top/center/bottom, 스키마·CSS 쪽은 top/middle/bottom을 쓰므로 여기서 되돌린다.
    const vAlignXml =
      verticalAlign === 'middle'
        ? '<w:vAlign w:val="center"/>'
        : verticalAlign === 'top' || verticalAlign === 'bottom'
          ? `<w:vAlign w:val="${verticalAlign}"/>`
          : '';
    const cellMarginXml = renderCellMarginXml(cell.attrs.cellPadding);
    cellsXml +=
      `<w:tc><w:tcPr><w:tcW w:w="${totalWidthDxa}" w:type="dxa"/>` +
      `${gridSpanXml}${vMergeXml}${bordersXml}${shdXml}${vAlignXml}${cellMarginXml}</w:tcPr>${cellBody}</w:tc>`;
    if (rowspan > 1) pendingVMerge.set(col, { remaining: rowspan - 1, colspan, widthDxa: totalWidthDxa });
    col += colspan;
  }

  return `<w:tr>${trPrXml}${cellsXml}</w:tr>`;
}

function renderTable(node: PMNode): string {
  // 표 전체 열 개수 — 세로 병합으로 뒤 행은 실제 셀이 더 적을 수 있어 모든 행 중 최댓값을 쓴다.
  let colCount = 0;
  node.forEach((row) => {
    let width = 0;
    row.forEach((cell) => {
      width += Number(cell.attrs.colspan) || 1;
    });
    colCount = Math.max(colCount, width);
  });

  const pendingVMerge = new Map<number, PendingVMerge>();
  const columnWidthsDxa = new Map<number, number>();
  let rowsXml = '';
  node.forEach((row) => {
    rowsXml += renderTableRow(row, pendingVMerge, columnWidthsDxa);
  });

  const gridCols = Array.from(
    { length: colCount },
    (_, i) => `<w:gridCol w:w="${columnWidthsDxa.get(i) ?? CELL_WIDTH_DXA}"/>`,
  ).join('');

  // 표 자체가 페이지 안에서 어디 놓이는지(셀 안 문단 정렬과는 별개) — 없으면(기본 왼쪽) 생략한다.
  const align = node.attrs.align;
  const jcXml = align === 'center' || align === 'right' ? `<w:jc w:val="${align}"/>` : '';

  return (
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${jcXml}<w:tblBorders>${BORDER}</w:tblBorders></w:tblPr>` +
    `<w:tblGrid>${gridCols}</w:tblGrid>${rowsXml}</w:tbl>`
  );
}

function renderBlock(node: PMNode): string {
  switch (node.type.name) {
    case 'paragraph':
      return renderParagraph(node);
    case 'heading':
      return renderHeading(node);
    case 'bullet_list':
      return renderList(node, false);
    case 'ordered_list':
      return renderList(node, true);
    case 'table':
      return renderTable(node);
    default:
      return '';
  }
}

const DOCUMENT_XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * sectPrXml을 생략하면(주로 테스트에서) 빈 <w:sectPr/>을 쓴다 — 실제 저장 경로(exportDocx)는
 * 항상 원본 문서에서 뽑아낸 실제 sectPrXml을 넘겨 페이지 크기·여백 등을 보존한다.
 */
export function renderDocumentXml(doc: PMNode, sectPrXml: string = '<w:sectPr/>'): string {
  let body = '';
  doc.forEach((node) => {
    body += renderBlock(node);
  });
  return (
    `${DOCUMENT_XML_HEADER}<w:document xmlns:w="${NS_W}"><w:body>${body}${sectPrXml}</w:body></w:document>`
  );
}

/**
 * ProseMirror 문서를 docx로 직렬화한다. 원본 zip의 다른 파트(styles.xml, media 등)는
 * 그대로 보존하고 word/document.xml만 교체한다 — 단, 이 함수가 만드는 document.xml은
 * 스타일 정의를 참조만 할 뿐(Heading1 등) 재정의하지 않으므로, 원본에 해당 스타일이
 * 없으면(예: 빈 스타일시트로 새로 만든 docx) 기본 서식으로 표시될 수 있다.
 *
 * w:sectPr(페이지 크기·여백 등)은 이 스키마가 전혀 모델링하지 않으므로, 매번 새로 만드는 대신
 * 원본 word/document.xml에서 그대로 뽑아 재사용한다 — 그래야 편집·재저장을 거쳐도 페이지
 * 설정이 유지된다(과거엔 항상 빈 <w:sectPr/>로 덮어써서 편집 후 페이지 크기가 사라졌었다).
 */
export function exportDocx(originalDocxBytes: Uint8Array, doc: PMNode): Uint8Array {
  const entries = guardedUnzipSync(originalDocxBytes, 'exportDocx');
  const originalDocumentXml = entries['word/document.xml'];
  const sectPrXml = (originalDocumentXml && extractSectPrXml(strFromU8(originalDocumentXml))) ?? '<w:sectPr/>';
  entries['word/document.xml'] = strToU8(renderDocumentXml(doc, sectPrXml));
  return zipSync(entries);
}
