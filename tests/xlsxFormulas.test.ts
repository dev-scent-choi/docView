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

  // 성능 수정(체인형 수식의 O(N²) 재계산을 막는 memo)이 기존 계산 결과를 바꾸지 않는지 확인하는
  // 회귀 테스트 — memo를 명시적으로 공유해도, 아예 안 써도(기본값) 같은 값이 나와야 한다.
  describe('memo (performance fix) does not change any computed value', () => {
    test('a chain of formula cells each referencing the previous one (A2=A1+1, A3=A2+1, ...) computes ' +
      'the mathematically correct value at the end of the chain, with or without a shared memo', () => {
      const N = 30;
      const rows: Array<Array<{ v: string } | 0>> = [[{ v: '1' }]];
      for (let r = 2; r <= N; r += 1) rows.push([{ v: `=A${r - 1}+1` }]);
      const grid = makeGrid(rows);

      // memo 없이(기본값 — 매 호출마다 새 Map): 재귀적으로 A1까지 완전히 다시 계산한다.
      expect(computeGridCellDisplay(grid, 0, N - 1, 0)).toBe(N);

      // memo를 명시적으로 공유해도 정확히 같은 값이어야 한다(체인 끝까지 재귀로 한 번에 계산되므로
      // 이 특정 호출 하나만으로는 memo가 실제로 캐시 적중을 만들어내지 않지만, 결과가 동일해야
      // 한다는 게 핵심 — memo 도입이 계산 로직 자체를 바꾸지 않았다는 증거).
      const memo = new Map<string, unknown>();
      expect(computeGridCellDisplay(grid, 0, N - 1, 0, new Set(), memo)).toBe(N);
    });

    test('two sibling formula cells that both depend on the same third formula cell (a "diamond" ' +
      'dependency: D references B and C, both of which reference A) get the correct value whether ' +
      'or not they share a memo across separate top-level calls', () => {
      // A1 = 10 (수식), B1 = A1*2 (수식), C1 = A1+5 (수식), D1 = B1+C1 (수식) — B1과 C1이 각각
      // 독립적으로 A1을 참조하는 재귀(재진입) 평가를 만든다. D1 = 20+15 = 35.
      const grid = makeGrid([[{ v: '=5+5' }, { v: '=A1*2' }, { v: '=A1+5' }, { v: '=B1+C1' }]]);

      // 각 셀을 memo 없이 독립적으로 호출해도 정확해야 한다(기존 동작 그대로).
      expect(computeGridCellDisplay(grid, 0, 0, 0)).toBe(10); // A1
      expect(computeGridCellDisplay(grid, 0, 0, 1)).toBe(20); // B1
      expect(computeGridCellDisplay(grid, 0, 0, 2)).toBe(15); // C1
      expect(computeGridCellDisplay(grid, 0, 0, 3)).toBe(35); // D1

      // 하나의 memo를 여러 "최상위" 호출에 걸쳐 공유해도(recomputeFormulasForPage가 페이지 하나의
      // 모든 수식 셀에 걸쳐 하나의 memo를 쓰는 것과 같은 패턴) 같은 결과가 나와야 한다 — A1이 B1과
      // C1 양쪽에서 재참조되므로, 이게 바로 memo가 실제로 적중하는(그리고 여전히 정확해야 하는)
      // 경우다.
      const memo = new Map<string, unknown>();
      expect(computeGridCellDisplay(grid, 0, 0, 0, new Set(), memo)).toBe(10);
      expect(computeGridCellDisplay(grid, 0, 0, 1, new Set(), memo)).toBe(20);
      expect(computeGridCellDisplay(grid, 0, 0, 2, new Set(), memo)).toBe(15);
      expect(computeGridCellDisplay(grid, 0, 0, 3, new Set(), memo)).toBe(35);
    });

    test('a genuine circular reference (A1=B1, B1=A1) still returns the REF error name, not a stale ' +
      'or incorrect memoized value, even when a memo is shared across the two top-level calls', () => {
      const grid = makeGrid([[{ v: '=B1' }, { v: '=A1' }]]);
      const memo = new Map<string, unknown>();
      expect(computeGridCellDisplay(grid, 0, 0, 0, new Set(), memo)).toBe('#REF!');
      expect(computeGridCellDisplay(grid, 0, 0, 1, new Set(), memo)).toBe('#REF!');
    });
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
