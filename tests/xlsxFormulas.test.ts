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
      'or incorrect memoized value, even when a memo is shared across the two top-level calls, ' +
      'regardless of which cell is queried first', () => {
      const memo = new Map<string, unknown>();
      const grid = makeGrid([[{ v: '=B1' }, { v: '=A1' }]]);
      expect(computeGridCellDisplay(grid, 0, 0, 0, new Set(), memo)).toBe('#REF!');
      expect(computeGridCellDisplay(grid, 0, 0, 1, new Set(), memo)).toBe('#REF!');

      // 순서를 뒤집어도(B1을 먼저 물어도) 같아야 한다 — 새 memo/그리드로, 이번엔 B1(col 1)을 먼저.
      const memo2 = new Map<string, unknown>();
      const grid2 = makeGrid([[{ v: '=B1' }, { v: '=A1' }]]);
      expect(computeGridCellDisplay(grid2, 0, 0, 1, new Set(), memo2)).toBe('#REF!');
      expect(computeGridCellDisplay(grid2, 0, 0, 0, new Set(), memo2)).toBe('#REF!');

      // 화이트박스 확인: 순환에 걸린 두 칸 다 memo에 영구 캐시되지 않아야 한다(taint) — 그래야
      // 이후 다른 패스에서 다시 물었을 때도 항상 새로 계산해 정확한 값을 준다.
      expect(memo.has('0:0:0')).toBe(false);
      expect(memo.has('0:0:1')).toBe(false);
    });

    // Finding 2의 정확한 재현 시나리오: A1=1(값), B1=SUM(A1:A2)(수식), A2=B1(수식) — B1과 A2가 서로를
    // 참조하는 "비대칭" 순환(B1은 직접 A2를 참조하지 않고 범위 SUM(A1:A2)를 통해 간접 참조, A2는
    // B1을 직접 참조). d2bff59의 무조건 memo.set은 이 경우 "먼저 계산된 칸이 어떤 하위 호출에서
    // 중간값으로 A2를 계산했는지"에 따라 A2의 캐시된 값이 달라지는 순서 의존 버그를 만들었다
    // (리뷰에서 실측: B1을 먼저 물으면 그 안에서 재귀적으로 계산된 A2의 중간값('#REF!', 사이클
    // 감지로 인한 값)이 그대로 A2의 최종 캐시로 굳어버려, 이후 A2를 단독으로 물어도 잘못된
    // '#REF!'가 캐시에서 튀어나온다 — A2를 먼저 물으면 다른 경로로 계산돼 우연히 정확한 값이
    // 나온다). tainted 수정은 사이클에 걸린 두 칸(B1, A2) 모두 캐시를 건너뛰므로, 어느 순서로
    // 묻든, 그리고 하나의 memo를 공유하든 각자 별도의 memo(별도 패스)를 쓰든 항상 같은(정확한)
    // 값을 줘야 한다.
    test('an asymmetric circular reference (A1=1, B1=SUM(A1:A2), A2=B1) returns the SAME value for ' +
      'A2 regardless of whether B1 or A2 is evaluated first within one pass, and regardless of ' +
      'whether they are split across two separate passes/memos (the exact scenario the taint fix ' +
      'addresses)', () => {
      const makeAsymmetricCycleGrid = () =>
        makeGrid([
          [{ v: '1' }, { v: '=SUM(A1:A2)' }],
          [{ v: '=B1' }, 0],
        ]);

      // Case 1: 한 memo를 공유하며 B1을 먼저 묻고, 그다음 A2를 묻는다.
      const memoCase1 = new Map<string, unknown>();
      const gridCase1 = makeAsymmetricCycleGrid();
      const b1First = computeGridCellDisplay(gridCase1, 0, 0, 1, new Set(), memoCase1);
      const a2AfterB1 = computeGridCellDisplay(gridCase1, 0, 1, 0, new Set(), memoCase1);

      // Case 2: 같은 그리드, 이번엔 A2를 먼저 묻고, 그다음 B1을 묻는다(하나의 memo 공유).
      const memoCase2 = new Map<string, unknown>();
      const gridCase2 = makeAsymmetricCycleGrid();
      const a2First = computeGridCellDisplay(gridCase2, 0, 1, 0, new Set(), memoCase2);
      const b1AfterA2 = computeGridCellDisplay(gridCase2, 0, 0, 1, new Set(), memoCase2);

      // 순서를 바꿔도 두 칸의 값이 똑같아야 한다 — d2bff59의 버그(위 문서 주석)는 정확히 이 지점에서
      // A2가 B1을 먼저 물었을 때와 A2를 먼저 물었을 때 서로 다른 값을 냈다.
      expect(a2AfterB1).toBe(a2First);
      expect(b1First).toBe(b1AfterA2);
      expect(a2AfterB1).toBe(b1First); // 이 특정 수식 구조에서는 B1과 A2가 같은 값으로 수렴해야 한다.

      // Case 3: 서로 완전히 다른 두 "패스"(별도 memo, 별도 evalGuard)에서 하나씩만 물어도 — 마치
      // B1은 페이지 1의 재계산에서, A2는 페이지 2의 재계산(또는 다음 재계산 패스)에서 계산되는
      // 것처럼 — 여전히 같은 값이어야 한다.
      const b1SplitPass = computeGridCellDisplay(makeAsymmetricCycleGrid(), 0, 0, 1, new Set(), new Map());
      const a2SplitPass = computeGridCellDisplay(makeAsymmetricCycleGrid(), 0, 1, 0, new Set(), new Map());
      expect(a2SplitPass).toBe(a2AfterB1);
      expect(b1SplitPass).toBe(b1First);

      // 화이트박스 확인: 사이클에 걸린 두 칸 다 어느 케이스에서도 영구 캐시되지 않아야 한다.
      expect(memoCase1.has('0:0:1')).toBe(false); // B1
      expect(memoCase1.has('0:1:0')).toBe(false); // A2
      expect(memoCase2.has('0:0:1')).toBe(false);
      expect(memoCase2.has('0:1:0')).toBe(false);
    });
  });
});

describe('xlsx/formulas.ts: FormulaParser instance pool (performance fix)', () => {
  // 풀이 없던 시절(d2bff59)에는 수식 셀 하나마다 `new FormulaParser({...})`를 새로 만들었다 — 그
  // 생성자 하나의 비용(~1.9ms, 리뷰 실측)이 계산 전체 비용의 99%였다. memo는 체인형(cascading)
  // 수식에서만 도움이 되고, 서로 참조하지 않는 "독립된" 수식 칸(예: 아래 SUM(B1)+C1 같은 단발성
  // 계산이 수천 번 반복되는 실제 사용자 시나리오)에는 각 칸이 정확히 한 번씩만 계산되므로 memo가
  // 전혀 도움이 안 됐다. 아래 테스트들은 새 풀 기반 구현이 (a) 기존과 동일한 값을 내고(정확성),
  // (b) 독립 수식 케이스에서도 실제로 빨라졌는지(성능) 확인한다.

  // "nested-function-with-trailing-tokens": 중첩된 수식 참조(A1)가 함수 인자 목록 "안에서"
  // 등장하고, 그 뒤로 그 함수를 닫는 토큰과 바깥 수식의 나머지 토큰("+B1")이 더 남아있는 모양 —
  // 안쪽 parse()가 풀에서 별도 인스턴스를 받아 자기 토큰 스트림만 건드리고, 바깥 parse()가 자신의
  // 인스턴스에서 남은 토큰("+B1")을 정확히 이어서 읽는지 검증한다(공유 인스턴스였다면 바로 이
  // 지점에서 토큰 스트림이 오염된다 — 문서 주석 참고).
  test('a nested formula reference inside a function call, followed by trailing tokens in the outer ' +
    'formula (ROUND(A1,2)+B1 where A1 is itself a formula), computes correctly', () => {
    const grid = makeGrid([[{ v: '=10/3' }, { v: '1' }, { v: '=ROUND(A1,2)+B1' }]]);
    expect(computeGridCellDisplay(grid, 0, 0, 2)).toBeCloseTo(4.33, 5);
  });

  // 재진입(reentrancy) 위험 표 — 인스턴스 하나를 공유했다면(리뷰가 실측한 대로) 11/14 케이스가
  // 조용히 잘못된 #ERROR!를 냈을 모양들. B1은 그 자체로 수식(=2+3=5)이라 이 셀들을 계산하는
  // 도중 반드시 "바깥 parse()가 안 끝난 채로 안쪽 parse()가 시작"되는 중첩이 실제로 일어난다.
  describe('reentrancy-hazard formula shapes (B1 is itself a formula, referenced from another formula)', () => {
    function makeReentrancyGrid(outerFormula: string): GridDocument {
      // A(placeholder)=0, B1='=2+3'(=5, 수식), C1='10'(값), D1=outerFormula.
      return makeGrid([[0, { v: '=2+3' }, { v: '10' }, { v: outerFormula }]]);
    }

    test('=SUM(B1)+C1 → SUM(5)+10 = 15', () => {
      const grid = makeReentrancyGrid('=SUM(B1)+C1');
      expect(computeGridCellDisplay(grid, 0, 0, 1)).toBe(5); // B1 자체도 정확해야 함
      expect(computeGridCellDisplay(grid, 0, 0, 3)).toBe(15);
    });

    test('=IF(B1>0,1,2)+C1 → IF(5>0,1,2)+10 = 11', () => {
      const grid = makeReentrancyGrid('=IF(B1>0,1,2)+C1');
      expect(computeGridCellDisplay(grid, 0, 0, 3)).toBe(11);
    });

    test('=CONCATENATE(B1,"x")&"y" → "5" + "x" + "y" = "5xy"', () => {
      const grid = makeReentrancyGrid('=CONCATENATE(B1,"x")&"y"');
      expect(computeGridCellDisplay(grid, 0, 0, 3)).toBe('5xy');
    });
  });

  // 40행 이상의 깊은 체인 — 재귀 한 단계마다 풀에서 인스턴스를 하나씩 더 체크아웃하므로(중첩된
  // parse() 호출이 동시에 여러 개 "미완료" 상태로 스택에 쌓임), 풀이 free-list 관리를 깊은 재귀
  // 깊이에서도 안전하게 처리하는지 확인한다.
  test('a 45-deep chain of formula cells (A2=A1+1, A3=A2+1, ..., A45=A44+1) computes the ' +
    'mathematically correct value at the end of the chain', () => {
    const N = 45;
    const rows: Array<Array<{ v: string } | 0>> = [[{ v: '1' }]];
    for (let r = 2; r <= N; r += 1) rows.push([{ v: `=A${r - 1}+1` }]);
    const grid = makeGrid(rows);
    expect(computeGridCellDisplay(grid, 0, N - 1, 0)).toBe(N);
  });

  // 성능(이 태스크의 실제 목적): 서로 참조하지 않는 "독립된" 수식 칸 500행 x 2열(총 1000칸) —
  // memo가 도움이 안 되는 바로 그 케이스(브리프의 실제 사용자 불만 시나리오). 리뷰의 프로토타입은
  // 이 케이스에서 ~1947ms → ~5ms(약 370배)를 측정했고, 이 구현으로 직접 재현한 격리 벤치마크
  // (scratch-bench, 이 테스트 파일 밖)에서도 naive 1897ms → pooled 6ms(약 316배)를 확인했다.
  // 여기서는 타이트한 임계값 대신 넉넉하지만 의미 있는 상한(수백 배 개선이 아니면 넘길 수 없는
  // 수준)만 건다 — CI 환경의 CPU 경합을 고려해 플레이크를 피하면서도, 풀 없이 되돌아가면(칸당
  // ~1.9ms × 1000칸 ≈ 1.9초) 절대 넘길 수 없는 상한이라는 게 핵심.
  test('500 rows of INDEPENDENT formula cells (no cell references another formula cell) compute ' +
    'quickly — the case d2bff59\'s memo gave zero measured improvement for', () => {
    const ROWS = 500;
    const rows: Array<Array<{ v: string } | 0>> = [];
    for (let r = 1; r <= ROWS; r += 1) {
      rows.push([{ v: String(r) }, { v: String(r * 2) }, { v: `=A${r}*B${r}` }, { v: `=A${r}+B${r}` }]);
    }
    const grid = makeGrid(rows);

    const start = Date.now();
    for (let r = 0; r < ROWS; r += 1) {
      computeGridCellDisplay(grid, 0, r, 2);
      computeGridCellDisplay(grid, 0, r, 3);
    }
    const elapsedMs = Date.now() - start;

    // 정확성도 함께 확인한다(성능만 보고 값이 깨진 걸 놓치지 않도록).
    expect(computeGridCellDisplay(grid, 0, 0, 2)).toBe(2); // C1 = A1*B1 = 1*2
    expect(computeGridCellDisplay(grid, 0, 0, 3)).toBe(3); // D1 = A1+B1 = 1+2
    expect(computeGridCellDisplay(grid, 0, ROWS - 1, 2)).toBe(ROWS * (ROWS * 2)); // C500 = 500*1000
    expect(computeGridCellDisplay(grid, 0, ROWS - 1, 3)).toBe(ROWS + ROWS * 2); // D500 = 500+1000

    // 넉넉한 상한 — 풀 없이(칸당 new FormulaParser)는 1000칸이면 실측 ~1.9초가 걸린다.
    expect(elapsedMs).toBeLessThan(300);
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
