import ExcelJS from 'exceljs';
import * as Y from 'yjs';
import {
  GRID_BASELINE_MAP,
  GRID_CELLS_MAP,
  GRID_MERGES_MAP,
  GRID_STRUCTURE_MAP,
  GRID_STYLES_BASELINE_MAP,
  GRID_STYLES_MAP,
  gridCellKey,
  parseGridCellKey,
  type CellStyleAttrs,
  type GridMergeShape,
  type SheetStructure,
} from '../collab/gridShared.js';
import type { GridDocument, GridStyle } from '../spreadsheet.js';
import { parseXlsxToGrid } from '../spreadsheet.js';

function toCellStyleAttrs(style: GridStyle): CellStyleAttrs {
  const attrs: { -readonly [K in keyof CellStyleAttrs]?: CellStyleAttrs[K] } = {};
  if (style.b) attrs.b = 1;
  if (style.i) attrs.i = 1;
  if (style.u) attrs.u = 1;
  if (style.st) attrs.st = 1;
  if (style.col) attrs.col = style.col;
  if (style.bg) attrs.bg = style.bg;
  if (style.a) attrs.a = style.a;
  if (style.va) attrs.va = style.va;
  if (style.w) attrs.wrap = 1;
  if (style.ff) attrs.ff = style.ff;
  if (style.fs) attrs.fs = style.fs;
  if (style.nf) attrs.nf = style.nf;
  if (style.nd) attrs.nd = style.nd;
  if (style.lk) attrs.lk = 1;
  if (style.nu) attrs.nu = 1;
  if (style.bt) attrs.bt = style.bt;
  if (style.br) attrs.br = style.br;
  if (style.bb) attrs.bb = style.bb;
  if (style.bl) attrs.bl = style.bl;
  return attrs;
}

/** 병합 범위(1-based, {r,c,rs,cs})에서 "앵커가 아닌, 병합에 덮인" 칸의 (0-based r:c) 키 집합을 만든다.
 * 이런 칸은 ExcelJS가 앵커와 동일한 값을 그대로 돌려주는 것뿐(진짜 값이 아님)이라 시딩에서 제외해야
 * 병합이 "14칸에 같은 제목이 중복 입력된 것"이 아니라 진짜 병합으로 남는다(그리드 에디터가 렌더링
 * 시 앵커에만 colspan/rowspan을 주고 이 칸들은 아예 그리지 않는다 — viewer.js의 처리와 동일한 원리). */
function mergeCoveredCells(merges: readonly GridMergeShape[]): Set<string> {
  const covered = new Set<string>();
  for (const m of merges) {
    const r0 = m.r - 1;
    const c0 = m.c - 1;
    for (let r = r0; r < r0 + m.rs; r += 1) {
      for (let c = c0; c < c0 + m.cs; c += 1) {
        if (r === r0 && c === c0) continue;
        covered.add(`${r}:${c}`);
      }
    }
  }
  return covered;
}

/** parseXlsxToGrid가 만든 표시용 텍스트(예: "2,287")를 협업 문서의 실제 셀 값으로 그대로 심으면
 * 안 된다 — Number("2,287")는 NaN이라, 이 값을 참조하는 다른 수식(SUM 등)이 이 셀을 숫자가 아닌
 * 것으로 취급해 조용히 0으로 계산해버린다(실제 재현된 버그: 합계 열이 항상 0으로 보임). 쉼표 구분
 * 표시 형식(number/currency)이 걸린 칸만, 화면표시용 텍스트에서 쉼표를 제거해 원래 숫자 문자열로
 * 되돌린 뒤 심는다 — "화면에 어떻게 보일지"는 style.nf/nd로 이미 별도로 심어지므로 여기서는 순수
 * 값만 복원하면 된다. */
function seedableCellValue(text: string, style: GridStyle | undefined): string {
  if (!style?.nf || (style.nf !== 'number' && style.nf !== 'currency')) return text;
  const stripped = text.replace(/[₩,\s]/g, '');
  return Number.isFinite(Number(stripped)) && stripped !== '' ? stripped : text;
}

/** 읽기전용 그리드 뷰(parseXlsxToGrid의 결과)를 협업용 Y.Doc에 시딩한다. */
export function seedGridFromXlsx(ydoc: Y.Doc, grid: GridDocument): void {
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const baseline = ydoc.getMap<string>(GRID_BASELINE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const stylesBaseline = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_BASELINE_MAP);
  const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  const merges = ydoc.getMap<readonly GridMergeShape[]>(GRID_MERGES_MAP);
  grid.sheets.forEach((sheet, sheetIdx) => {
    structure.set(String(sheetIdx), {
      name: sheet.name,
      rowCount: sheet.rowCount,
      colCount: sheet.colCount,
      freezeRows: sheet.freezeRows,
      freezeCols: sheet.freezeCols,
      colWidths: sheet.colWidths,
    });
    if (sheet.merges.length > 0) merges.set(String(sheetIdx), sheet.merges);
    const covered = mergeCoveredCells(sheet.merges);
    sheet.rows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (cell === 0) return;
        if (covered.has(`${r}:${c}`)) return; // 병합에 덮인 칸 — 앵커의 중복 값이라 시딩하지 않는다
        const key = gridCellKey(sheetIdx, r, c);
        const style = cell.s !== undefined ? grid.styles[cell.s] : undefined;
        if (cell.v !== '') {
          const value = seedableCellValue(cell.v, style);
          cells.set(key, value);
          baseline.set(key, value);
        }
        if (style) {
          const attrs = toCellStyleAttrs(style);
          if (Object.keys(attrs).length > 0) {
            styles.set(key, attrs);
            stylesBaseline.set(key, attrs);
          }
        }
      });
    });
  });
}

/** '='로 시작하면 수식으로 저장한다(계산은 하지 않음 — 실제 결과는 저장 후 엑셀/뷰어에서 재계산된다). */
function coerceCellValue(raw: string): string | number | { formula: string } {
  const trimmed = raw.trim();
  if (trimmed.startsWith('=') && trimmed.length > 1) {
    return { formula: trimmed.slice(1) };
  }
  if (trimmed !== '' && /^-?\d+(\.\d+)?$/.test(trimmed)) {
    const num = Number(trimmed);
    if (Number.isFinite(num)) return num;
  }
  return raw;
}

const ALIGN_TO_EXCEL: Record<'l' | 'c' | 'r', ExcelJS.Alignment['horizontal']> = {
  l: 'left',
  c: 'center',
  r: 'right',
};
const VALIGN_TO_EXCEL: Record<'t' | 'm' | 'b', ExcelJS.Alignment['vertical']> = {
  t: 'top',
  m: 'middle',
  b: 'bottom',
};
const BORDER_WEIGHT_TO_EXCEL: Record<1 | 2, ExcelJS.BorderStyle> = {
  1: 'thin',
  2: 'medium',
};
/** spreadsheet.ts의 PT_TO_PX(96/72)와 정반대 변환 — 저장은 항상 px 기준(CellStyleAttrs.fs)이라
 * 실제 엑셀 파일에 쓸 때만 pt로 되돌린다. */
const PX_TO_PT = 72 / 96;

function numFmtStringFor(style: CellStyleAttrs): string | undefined {
  if (!style.nf) return undefined;
  const decimals = style.nd ? `.${'0'.repeat(style.nd)}` : '';
  switch (style.nf) {
    case 'number': return `#,##0${decimals}`;
    case 'currency': return `"₩"#,##0${decimals}`;
    case 'percent': return `0${decimals}%`;
    default: return undefined;
  }
}

function applyStyleToCell(cell: ExcelJS.Cell, style: CellStyleAttrs | undefined): void {
  if (!style) {
    cell.font = {};
    cell.fill = { type: 'pattern', pattern: 'none' };
    cell.alignment = {};
    cell.border = {};
    cell.numFmt = 'General';
    return;
  }
  cell.font = {
    bold: style.b === 1,
    italic: style.i === 1,
    underline: style.u === 1,
    strike: style.st === 1,
    color: style.col ? { argb: `FF${style.col.replace('#', '').toUpperCase()}` } : undefined,
    name: style.ff,
    size: style.fs ? Math.round(style.fs * PX_TO_PT) : undefined,
  };
  cell.fill = style.bg
    ? { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${style.bg.replace('#', '').toUpperCase()}` } }
    : { type: 'pattern', pattern: 'none' };
  cell.alignment = {
    horizontal: style.a ? ALIGN_TO_EXCEL[style.a] : undefined,
    vertical: style.va ? VALIGN_TO_EXCEL[style.va] : undefined,
    wrapText: style.wrap === 1,
  };
  cell.border = {
    top: style.bt ? { style: BORDER_WEIGHT_TO_EXCEL[style.bt] } : undefined,
    right: style.br ? { style: BORDER_WEIGHT_TO_EXCEL[style.br] } : undefined,
    bottom: style.bb ? { style: BORDER_WEIGHT_TO_EXCEL[style.bb] } : undefined,
    left: style.bl ? { style: BORDER_WEIGHT_TO_EXCEL[style.bl] } : undefined,
  };
  cell.numFmt = numFmtStringFor(style) ?? 'General';
}

/**
 * 원본 xlsx(수식·서식 등 우리가 편집하지 않는 모든 것)를 그대로 로드한 뒤,
 * 실제로 사용자가 바꾼 셀(값·서식 각각)만 덮어써 재직렬화한다 — docx export와 같은
 * "바뀐 부분만 교체" 원칙. baseline과 값이 같은 셀(한 번도 편집되지 않은 셀 — 수식 셀 포함)은
 * 건드리지 않는다: 수식 셀은 시딩 시 계산된 결괏값 문자열로만 채워지므로, 건드리지 않았는데도
 * `cells`에 있다고 전부 써버리면 수식 자체가 정적 값으로 깨진다.
 * 행/열 삭제(구조 축소)는 시트 끝에서부터만 지원한다(v1 스코프 단순화).
 */
export async function exportXlsxFromGrid(originalXlsxBytes: Uint8Array, ydoc: Y.Doc): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalXlsxBytes as unknown as ArrayBuffer);
  const visibleSheets = workbook.worksheets.filter((sheet) => !sheet.state || sheet.state === 'visible');

  // ExcelJS는 시트 보호(sheetProtection)를 읽어오긴 하지만, 아무 것도 안 건드리고 그대로
  // writeBuffer()만 하면 그 정보를 재직렬화하지 않고 조용히 잃어버린다(ExcelJS 자체의 한계로 보임 —
  // protect()를 이번 세션에서 다시 명시적으로 불러야만 저장 시 살아남는다). 그래서 로드 직후
  // 상태를 기억해뒀다가 쓰기 직전에 다시 걸어준다 — 안 그러면 편집기 저장을 한 번이라도 거치고
  // 나면(사용자가 아무것도 안 건드려도 유휴 자동저장으로 매번 일어남) 잠금이 전부 풀려버린다.
  const wasProtected = visibleSheets.map((sheet) =>
    Boolean((sheet as unknown as { sheetProtection?: unknown }).sheetProtection),
  );

  const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const baseline = ydoc.getMap<string>(GRID_BASELINE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const stylesBaseline = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_BASELINE_MAP);

  for (const [key, target] of structure) {
    const sheetIdx = Number(key);
    const sheet = visibleSheets[sheetIdx];
    if (!sheet) continue;

    if (sheet.rowCount > target.rowCount) {
      sheet.spliceRows(target.rowCount + 1, sheet.rowCount - target.rowCount);
    }
    if (sheet.columnCount > target.colCount) {
      sheet.spliceColumns(target.colCount + 1, sheet.columnCount - target.colCount);
    }
  }

  const touchedValueKeys = new Set<string>([...baseline.keys(), ...cells.keys()]);
  for (const key of touchedValueKeys) {
    const before = baseline.get(key);
    const after = cells.get(key);
    if (before === after) continue; // 편집되지 않음 (수식 셀 포함) — 원본 그대로 둔다

    const parsed = parseGridCellKey(key);
    if (!parsed) continue;
    const target = structure.get(String(parsed.sheetIdx));
    const sheet = visibleSheets[parsed.sheetIdx];
    if (!target || !sheet) continue;
    if (parsed.row >= target.rowCount || parsed.col >= target.colCount) continue;
    sheet.getCell(parsed.row + 1, parsed.col + 1).value =
      after === undefined || after === '' ? null : coerceCellValue(after);
  }

  const touchedStyleKeys = new Set<string>([...stylesBaseline.keys(), ...styles.keys()]);
  for (const key of touchedStyleKeys) {
    const before = stylesBaseline.get(key);
    const after = styles.get(key);
    if (before === after) continue;
    if (before && after && JSON.stringify(before) === JSON.stringify(after)) continue;

    const parsed = parseGridCellKey(key);
    if (!parsed) continue;
    const target = structure.get(String(parsed.sheetIdx));
    const sheet = visibleSheets[parsed.sheetIdx];
    if (!target || !sheet) continue;
    if (parsed.row >= target.rowCount || parsed.col >= target.colCount) continue;
    applyStyleToCell(sheet.getCell(parsed.row + 1, parsed.col + 1), after);
  }

  for (let i = 0; i < visibleSheets.length; i += 1) {
    if (wasProtected[i]) {
      await visibleSheets[i].protect('', {});
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/** 저장 직후 뷰어용 그리드(읽기전용 열람 캐시)를 새 xlsx 바이트에서 다시 뽑아낸다. */
export function reparseGridForViewer(newXlsxBytes: Uint8Array): Promise<GridDocument> {
  return parseXlsxToGrid(Buffer.from(newXlsxBytes));
}
