import { describe, expect, test } from 'vitest';
import { computeGridCellDisplay, formatComputedValueForStorage } from '../src/xlsx/formulas.js';
import type { GridDocument } from '../src/spreadsheet.js';
import { EDIT_SOFT_ROW_LIMIT, EDIT_SOFT_COL_LIMIT } from '../src/collab/gridShared.js';

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

  test('a range formula whose cell count exceeds EDIT_SOFT_ROW_LIMIT * EDIT_SOFT_COL_LIMIT returns #REF! without iterating the range', () => {
    // 범위가 (EDIT_SOFT_ROW_LIMIT + 1)행 x EDIT_SOFT_COL_LIMIT열 = 30060칸으로 상한(30000)을
    // 넘는다. 그리드 자체는 아주 작게(빈 셀 1개) 둬서, 상한 체크가 실제로 범위를 순회하기 전에
    // 걸린다는 것을 확인한다 — 큰 픽스처를 만들 필요가 없다는 게 핵심.
    // 엑셀 열 문자는 26진법(A-Z)이라 60번째 열은 한 글자로 안 끝난다(A=1..Z=26, AA=27..BH=60).
    function colToLetters(n: number): string {
      let result = '';
      let x = n;
      while (x > 0) {
        const rem = (x - 1) % 26;
        result = String.fromCharCode(65 + rem) + result;
        x = Math.floor((x - 1) / 26);
      }
      return result;
    }
    const lastCol = colToLetters(EDIT_SOFT_COL_LIMIT); // 60번째 열 = "BH"
    const formula = `=SUM(A1:${lastCol}${EDIT_SOFT_ROW_LIMIT + 1})`;
    const gridWithFormula = makeGrid([[{ v: formula }]]);
    expect(computeGridCellDisplay(gridWithFormula, 0, 0, 0)).toBe('#REF!');
  });

  test('a range formula just under the cap still computes normally (regression guard)', () => {
    // 50행 x 10열 = 500칸, 상한(30000)에 한참 못 미침 — 정상적으로 계산돼야 한다.
    const rows: Array<Array<{ v: string } | 0>> = Array.from({ length: 50 }, () =>
      Array.from({ length: 10 }, () => ({ v: '1' })),
    );
    rows.push([{ v: '=SUM(A1:J50)' }]);
    const grid = makeGrid(rows);
    expect(computeGridCellDisplay(grid, 0, 50, 0)).toBe(500);
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
