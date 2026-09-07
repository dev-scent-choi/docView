import FormulaParser, { FormulaError, type CellRef, type RangeRef } from 'fast-formula-parser';
import { formulaFunctions, formulaFunctionsNeedContext } from '../client/formulaOverrides.js';
import type { GridDocument } from '../spreadsheet.js';

/** 셀 하나의 원본 텍스트를 canonical 그리드에서 읽는다 — computeGridCellDisplay의 재귀
 * 참조(onCell/onRange)와 최초 호출이 공유하는 저수준 접근자. */
function rawCellText(grid: GridDocument, sheetIdx: number, row: number, col: number): string {
  const sheet = grid.sheets[sheetIdx];
  if (!sheet) return '';
  const cell = sheet.rows[row]?.[col];
  return cell === 0 || cell === undefined ? '' : cell.v;
}

/**
 * 셀 하나의 "표시용" 값을 계산한다 — 클라이언트의 옛 computeCellDisplay(grid-editor.ts)와
 * 동일한 로직이지만, Yjs cells 맵 대신 canonical GridDocument를 직접 읽으므로 화면 밖(로드
 * 안 된 페이지) 참조도 항상 정확하다. 순환 참조 가드도 호출마다 새로 만든다(재계산이 매번
 * 독립적이어야 하므로, 클라이언트처럼 클로저에 고정된 evalGuard를 쓰지 않는다).
 */
export function computeGridCellDisplay(
  grid: GridDocument,
  sheetIdx: number,
  row: number,
  col: number,
  evalGuard: Set<string> = new Set(),
): unknown {
  const key = `${sheetIdx}:${row}:${col}`;
  const raw = rawCellText(grid, sheetIdx, row, col);
  if (!raw.startsWith('=')) {
    if (raw.trim() === '') return '';
    const num = Number(raw);
    return Number.isFinite(num) ? num : raw;
  }
  if (evalGuard.has(key)) return FormulaError.REF.name;
  evalGuard.add(key);
  try {
    const parser = new FormulaParser({
      functions: formulaFunctions,
      functionsNeedContext: formulaFunctionsNeedContext,
      onCell: (ref: CellRef) => computeGridCellDisplay(grid, sheetIdx, ref.row - 1, ref.col - 1, evalGuard),
      onRange: (ref: RangeRef) => {
        const rows: unknown[][] = [];
        for (let r = ref.from.row; r <= ref.to.row; r += 1) {
          const rowVals: unknown[] = [];
          for (let c = ref.from.col; c <= ref.to.col; c += 1) {
            rowVals.push(computeGridCellDisplay(grid, sheetIdx, r - 1, c - 1, evalGuard));
          }
          rows.push(rowVals);
        }
        return rows;
      },
    });
    return parser.parse(raw.slice(1), { row: row + 1, col: col + 1, sheet: 'Sheet1' });
  } catch (error) {
    return error instanceof FormulaError ? error.name : '#ERROR!';
  } finally {
    evalGuard.delete(key);
  }
}

/** computeGridCellDisplay의 결과를 GRID_COMPUTED_MAP에 저장할 문자열로 바꾼다 — 서식(쉼표·
 * 통화 등)은 클라이언트가 스타일 정보로 따로 입히므로 여기서는 순수 값만 문자열화한다. */
export function formatComputedValueForStorage(value: unknown): string {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number') return value === null || value === undefined ? '' : String(value);
  return String(Math.round(value * 1e10) / 1e10);
}
