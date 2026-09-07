import { describe, expect, test } from 'vitest';
import { computeGridCellDisplay, formatComputedValueForStorage } from '../src/xlsx/formulas.js';
import type { GridDocument } from '../src/spreadsheet.js';

function makeGrid(rows: Array<Array<{ v: string } | 0>>): GridDocument {
  return {
    styles: [],
    sheets: [
      {
        name: 'Sheet1',
        rowCount: rows.length,
        colCount: Math.max(...rows.map((r) => r.length)),
        colWidths: [],
        merges: [],
        freezeRows: 0,
        freezeCols: 0,
        rows,
      },
    ],
  };
}

describe('xlsx/formulas.ts: computeGridCellDisplay', () => {
  test('a plain value cell returns its own text (numbers coerced to number)', () => {
    const grid = makeGrid([[{ v: '10' }, { v: 'hello' }]]);
    expect(computeGridCellDisplay(grid, 0, 0, 0)).toBe(10);
    expect(computeGridCellDisplay(grid, 0, 0, 1)).toBe('hello');
  });

  test('SUM formula referencing cells on the same visible page computes correctly', () => {
    const grid = makeGrid([
      [{ v: '10' }],
      [{ v: '20' }],
      [{ v: '=SUM(A1:A2)' }],
    ]);
    expect(computeGridCellDisplay(grid, 0, 2, 0)).toBe(30);
  });

  test('a formula referencing a row far outside any 500-row page still computes correctly', () => {
    // 화면 밖(다른 페이지) 참조가 서버 canonical 기준으로는 항상 정확해야 한다는 게 이 설계의
    // 핵심 — 5000번째 행(10번째 페이지)을 0번째 행에서 참조.
    const rows: Array<Array<{ v: string } | 0>> = Array.from({ length: 5001 }, () => [0]);
    rows[4999] = [{ v: '42' }]; // 0-based 행 4999 = 5000번째 행
    rows[0] = [{ v: '=A5000' }];
    const grid = makeGrid(rows);
    expect(computeGridCellDisplay(grid, 0, 0, 0)).toBe(42);
  });

  test('circular reference returns the REF error name, not an infinite loop', () => {
    const grid = makeGrid([[{ v: '=A1' }]]);
    expect(computeGridCellDisplay(grid, 0, 0, 0)).toBe('#REF!');
  });

  test('MAX (a fast-formula-parser stub overridden in formulaOverrides) works correctly', () => {
    const grid = makeGrid([[{ v: '3' }, { v: '7' }, { v: '=MAX(A1:B1)' }]]);
    expect(computeGridCellDisplay(grid, 0, 0, 2)).toBe(7);
  });
});

describe('xlsx/formulas.ts: formatComputedValueForStorage', () => {
  test('formats a number, trimming floating point noise', () => {
    expect(formatComputedValueForStorage(0.1 + 0.2)).toBe('0.3');
  });
  test('formats a boolean as TRUE/FALSE', () => {
    expect(formatComputedValueForStorage(true)).toBe('TRUE');
  });
  test('formats an error name as-is', () => {
    expect(formatComputedValueForStorage('#DIV/0!')).toBe('#DIV/0!');
  });
});
