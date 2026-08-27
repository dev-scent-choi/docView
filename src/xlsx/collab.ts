import ExcelJS from 'exceljs';
import * as Y from 'yjs';
import {
  GRID_BASELINE_MAP,
  GRID_CELLS_MAP,
  GRID_STRUCTURE_MAP,
  GRID_STYLES_BASELINE_MAP,
  GRID_STYLES_MAP,
  gridCellKey,
  parseGridCellKey,
  type CellStyleAttrs,
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
  return attrs;
}

/** 읽기전용 그리드 뷰(parseXlsxToGrid의 결과)를 협업용 Y.Doc에 시딩한다. */
export function seedGridFromXlsx(ydoc: Y.Doc, grid: GridDocument): void {
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const baseline = ydoc.getMap<string>(GRID_BASELINE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const stylesBaseline = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_BASELINE_MAP);
  const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  grid.sheets.forEach((sheet, sheetIdx) => {
    structure.set(String(sheetIdx), {
      name: sheet.name,
      rowCount: sheet.rowCount,
      colCount: sheet.colCount,
    });
    sheet.rows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (cell === 0) return;
        const key = gridCellKey(sheetIdx, r, c);
        if (cell.v !== '') {
          cells.set(key, cell.v);
          baseline.set(key, cell.v);
        }
        if (cell.s !== undefined) {
          const style = grid.styles[cell.s];
          if (style) {
            const attrs = toCellStyleAttrs(style);
            if (Object.keys(attrs).length > 0) {
              styles.set(key, attrs);
              stylesBaseline.set(key, attrs);
            }
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

function applyStyleToCell(cell: ExcelJS.Cell, style: CellStyleAttrs | undefined): void {
  if (!style) {
    cell.font = {};
    cell.fill = { type: 'pattern', pattern: 'none' };
    cell.alignment = {};
    return;
  }
  cell.font = {
    bold: style.b === 1,
    italic: style.i === 1,
    underline: style.u === 1,
    strike: style.st === 1,
    color: style.col ? { argb: `FF${style.col.replace('#', '').toUpperCase()}` } : undefined,
  };
  cell.fill = style.bg
    ? { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${style.bg.replace('#', '').toUpperCase()}` } }
    : { type: 'pattern', pattern: 'none' };
  cell.alignment = style.a ? { horizontal: ALIGN_TO_EXCEL[style.a] } : {};
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

  const buffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/** 저장 직후 뷰어용 그리드(읽기전용 열람 캐시)를 새 xlsx 바이트에서 다시 뽑아낸다. */
export function reparseGridForViewer(newXlsxBytes: Uint8Array): Promise<GridDocument> {
  return parseXlsxToGrid(Buffer.from(newXlsxBytes));
}
