import { strToU8, unzipSync, zipSync } from 'fflate';
import type { Mark, Node as PMNode } from 'prosemirror-model';

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

const CELL_WIDTH_DXA = 2000;
const BORDER = '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>';

/** 세로 병합 중(pending) 열의 상태 — 몇 행 더 w:vMerge 이어짐 칸을 채워야 하는지, 그 칸의 가로폭(colspan) */
interface PendingVMerge {
  remaining: number;
  readonly colspan: number;
}

/**
 * ProseMirror 표 한 행을 docx 행으로 직렬화한다. prosemirror-tables는 세로 병합을 "이어지는
 * 칸은 그 행의 셀 목록에 아예 없다"는 방식으로 표현하므로(src/docx/import.ts의 tableFrom과
 * 동일한 모델), 이 행의 실제 셀들 사이사이에 이전 행에서 시작된 세로 병합(pendingVMerge)이
 * 남아 있는 열이 있으면 그 자리에 w:vMerge 이어짐 칸을 끼워 넣어 열 정렬을 맞춘다. 실제 셀과
 * pendingVMerge 칸은 열 위치가 절대 겹치지 않으므로(겹치면 애초에 이 행에 그 열의 실제 셀이
 * 없었을 것이다) 열 번호를 따로 계산할 필요 없이 앞에서부터 순서대로 배치하면 된다.
 */
function renderTableRow(row: PMNode, pendingVMerge: Map<number, PendingVMerge>): string {
  const cells: PMNode[] = [];
  row.forEach((cell) => cells.push(cell));

  let cellsXml = '';
  let col = 0;
  let cellIndex = 0;

  while (cellIndex < cells.length || pendingVMerge.has(col)) {
    const pending = pendingVMerge.get(col);
    if (pending) {
      const gridSpanXml = pending.colspan > 1 ? `<w:gridSpan w:val="${pending.colspan}"/>` : '';
      cellsXml +=
        `<w:tc><w:tcPr><w:tcW w:w="${CELL_WIDTH_DXA * pending.colspan}" w:type="dxa"/>` +
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
    let cellBody = '';
    cell.forEach((block) => {
      cellBody += renderBlock(block);
    });
    const gridSpanXml = colspan > 1 ? `<w:gridSpan w:val="${colspan}"/>` : '';
    const vMergeXml = rowspan > 1 ? '<w:vMerge w:val="restart"/>' : '';
    cellsXml +=
      `<w:tc><w:tcPr><w:tcW w:w="${CELL_WIDTH_DXA * colspan}" w:type="dxa"/>` +
      `${gridSpanXml}${vMergeXml}</w:tcPr>${cellBody}</w:tc>`;
    if (rowspan > 1) pendingVMerge.set(col, { remaining: rowspan - 1, colspan });
    col += colspan;
  }

  return `<w:tr>${cellsXml}</w:tr>`;
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
  const gridCols = Array.from({ length: colCount }, () => `<w:gridCol w:w="${CELL_WIDTH_DXA}"/>`).join('');

  let rowsXml = '';
  const pendingVMerge = new Map<number, PendingVMerge>();
  node.forEach((row) => {
    rowsXml += renderTableRow(row, pendingVMerge);
  });

  return (
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${BORDER}</w:tblBorders></w:tblPr>` +
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

export function renderDocumentXml(doc: PMNode): string {
  let body = '';
  doc.forEach((node) => {
    body += renderBlock(node);
  });
  return (
    `${DOCUMENT_XML_HEADER}<w:document xmlns:w="${NS_W}"><w:body>${body}<w:sectPr/></w:body></w:document>`
  );
}

/**
 * ProseMirror 문서를 docx로 직렬화한다. 원본 zip의 다른 파트(styles.xml, media 등)는
 * 그대로 보존하고 word/document.xml만 교체한다 — 단, 이 함수가 만드는 document.xml은
 * 스타일 정의를 참조만 할 뿐(Heading1 등) 재정의하지 않으므로, 원본에 해당 스타일이
 * 없으면(예: 빈 스타일시트로 새로 만든 docx) 기본 서식으로 표시될 수 있다.
 */
export function exportDocx(originalDocxBytes: Uint8Array, doc: PMNode): Uint8Array {
  const entries = unzipSync(originalDocxBytes);
  entries['word/document.xml'] = strToU8(renderDocumentXml(doc));
  return zipSync(entries);
}
