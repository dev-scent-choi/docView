import FormulaParser, { type CellRef, type RangeRef } from 'fast-formula-parser';
const { FormulaError } = FormulaParser;
import { formulaFunctions, formulaFunctionsNeedContext } from '../client/formulaOverrides.js';
import type { GridDocument } from '../spreadsheet.js';
import { EDIT_SOFT_ROW_LIMIT, EDIT_SOFT_COL_LIMIT, gridCellKey } from '../collab/gridShared.js';
import type { LazyGridStore } from './lazyGridStore.js';
import { getRow } from './lazyGridStore.js';

/** onRange 콜백 하나가 순회할 수 있는 셀 수의 상한 — "가장 큰 상식적인 인터랙티브 시트 영역"이라는
 * 기존 전례(EDIT_SOFT_ROW_LIMIT/EDIT_SOFT_COL_LIMIT, gridShared.ts)를 그대로 재사용한다.
 * fast-formula-parser의 onRange는 본질적으로 동기 콜백이라 중간에 양보(await)할 수 없으므로,
 * =SUM(A1:A50000) 같은 수식이 이벤트 루프를 통째로 멈추는 걸 막으려면 "얼마나 순회했는지"가
 * 아니라 "순회를 시작하기도 전에" 범위 크기를 걸러야 한다. */
const MAX_RANGE_CELLS = EDIT_SOFT_ROW_LIMIT * EDIT_SOFT_COL_LIMIT;

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
 *
 * **memo(성능)**: 수식 셀 하나를 계산할 때마다 `new FormulaParser({...})`를 새로 만드는 비용
 * 자체는 여기서 없애지 않는다 — `fast-formula-parser`의 FormulaParser는 인스턴스 상태(`this.position`,
 * `this.parser.input`, 즉 렉서/파서가 물고 있는 토큰 스트림)를 `parse()` 호출마다 덮어쓰므로, 한
 * 인스턴스를 "중첩된"(재진입) `parse()` 호출에 공유하면 바깥 호출이 아직 끝나지 않은 채로 안쪽
 * 호출이 같은 인스턴스의 `this.position`/`this.parser.input`을 덮어써 버려 바깥 호출이 이어서 읽는
 * 토큰/좌표가 안쪽 호출의 것으로 오염된다(재귀적으로 다른 수식 셀을 참조하는 게 바로 이 상황을
 * 만든다 — onCell/onRange 콜백이 아직 안 끝난 parse() 호출 "안에서" 다시 다른 셀의 수식을 평가하기
 * 때문). 그래서 인스턴스 재사용 대신, 이미 계산된 셀의 **결과값**을 이 재계산 패스 동안만 캐시한다.
 * 이렇게 하면 체인형 수식(`C{r} = C{r-1} + B{r}`처럼 뒤 칸이 앞 칸을 참조하는 경우) 한 페이지를
 * 재계산할 때, 뒤에서부터 훑든 앞에서부터 훑든 각 칸이 "이 패스에서" 최대 한 번만 실제로 계산된다
 * — memo가 없으면 셀 N개가 사슬로 이어져 있을 때 총 계산 횟수가 O(N²)로 폭증한다(N번째 칸을 계산하며
 * 1..N-1번째를 전부 다시 계산하고, N-1번째 칸을 "그 자신도 이 페이지의 수식 셀이라" 별도로 다시
 * 계산할 때 또 1..N-2번째를 전부 다시 계산하는 식). 캐시 키가 좌표뿐이므로 이 memo는 반드시
 * "그리드 내용이 바뀌지 않는다고 보장되는 하나의 재계산 패스" 동안만(호출자가 매번 새 Map을 만들어)
 * 써야 한다 — 여러 패스에 걸쳐 재사용하면 그 사이에 바뀐 셀 값을 못 보고 낡은 값을 돌려주게 된다.
 */
export function computeGridCellDisplay(
  grid: GridDocument,
  sheetIdx: number,
  row: number,
  col: number,
  evalGuard: Set<string> = new Set(),
  memo: Map<string, unknown> = new Map(),
): unknown {
  const key = `${sheetIdx}:${row}:${col}`;
  const raw = rawCellText(grid, sheetIdx, row, col);
  if (!raw.startsWith('=')) {
    if (raw.trim() === '') return '';
    const num = Number(raw);
    return Number.isFinite(num) ? num : raw;
  }
  if (memo.has(key)) return memo.get(key);
  if (evalGuard.has(key)) return FormulaError.REF.name;
  evalGuard.add(key);
  // onRange가 너무 큰 범위를 만났음을 기록하는 플래그 — fast-formula-parser의 parse()는 평가 중
  // 동기적으로 던져진 에러를 전부 FormulaError.ERROR('#ERROR!')로 재포장해버려서(원본 에러의
  // name을 보존하지 않음, node_modules/fast-formula-parser/grammar/hooks.js의 parse() catch
  // 블록 참고), onRange 안에서 FormulaError.REF를 던져도 우리 쪽 catch까지 '#REF!'로 도달하지
  // 않는다. 그래서 던지는 대신 플래그로 표시해두고, parser.parse()가 끝난 뒤(무엇을 반환했든)
  // 그 결과를 덮어써서 '#REF!'를 돌려준다.
  let rangeTooLarge = false;
  let result: unknown;
  try {
    const parser = new FormulaParser({
      functions: formulaFunctions,
      functionsNeedContext: formulaFunctionsNeedContext,
      onCell: (ref: CellRef) => computeGridCellDisplay(grid, sheetIdx, ref.row - 1, ref.col - 1, evalGuard, memo),
      onRange: (ref: RangeRef) => {
        const cellCount = (ref.to.row - ref.from.row + 1) * (ref.to.col - ref.from.col + 1);
        if (cellCount > MAX_RANGE_CELLS) {
          // 순회를 시작하기 전에 걸러낸다 — "너무 크다는 걸 알아챌 때까지 순회"하면 이미
          // 이벤트 루프를 막아버린 뒤라 상한을 두는 의미가 없다. 빈 배열을 돌려줘 이 범위를
          // 소비하는 함수(SUM 등)가 아무 것도 순회하지 않게 하고, 최종 결과는 아래에서 덮어쓴다.
          rangeTooLarge = true;
          return [];
        }
        const rows: unknown[][] = [];
        for (let r = ref.from.row; r <= ref.to.row; r += 1) {
          const rowVals: unknown[] = [];
          for (let c = ref.from.col; c <= ref.to.col; c += 1) {
            rowVals.push(computeGridCellDisplay(grid, sheetIdx, r - 1, c - 1, evalGuard, memo));
          }
          rows.push(rowVals);
        }
        return rows;
      },
    });
    const parsed = parser.parse(raw.slice(1), { row: row + 1, col: col + 1, sheet: 'Sheet1' });
    result = rangeTooLarge ? FormulaError.REF.name : parsed;
  } catch (error) {
    result = rangeTooLarge ? FormulaError.REF.name : (error instanceof FormulaError ? error.name : '#ERROR!');
  } finally {
    evalGuard.delete(key);
  }
  memo.set(key, result);
  return result;
}

/** computeGridCellDisplay와 완전히 같은 평가 로직이지만, 데이터 소스가 GridDocument 동기
 * 인덱싱이 아니라 "이미 캐시에 채워져 있어야 하는" LazyGridStore 동기 조회다. 캐시에 없는
 * 좌표를 만나면 던지지 않고(fast-formula-parser가 동기 콜백 안의 예외를 재포장해버림 —
 * rangeTooLarge와 같은 이유) onMiss 콜백으로 알린다.
 *
 * memo는 computeGridCellDisplay의 것과 같은 목적(체인형 수식의 O(N²) 재계산 방지)이지만, 호출자
 * (computeGridCellDisplayAsync)가 "이번 시도(attempt)에서 새로 계산한 값"만 담은 임시 Map을
 * 넘긴다 — 시도 도중 캐시 미스가 하나라도 나면 그 시도에서 나온 값은(미스 난 칸을 빈 칸 취급한
 * 채로 계산됐을 수 있어) 통째로 버려야 하므로, 이 함수 자신은 "언제 값을 영구히 믿어도 되는지"를
 * 몰라도 된다 — 그 판단은 호출자의 몫이다. */
function computeGridCellDisplaySync(
  store: LazyGridStore,
  sheetIdx: number,
  row: number,
  col: number,
  evalGuard: Set<string>,
  memo: Map<string, unknown>,
  onMiss: (sheetIdx: number, row: number) => void,
): unknown {
  const rawCellTextSync = (r: number, c: number): string => {
    const originalCol = store.logicalToOriginalCol[sheetIdx]?.[c];
    const key = gridCellKey(sheetIdx, r, c);
    if (store.dirtyOverlay.has(key)) {
      const cell = store.dirtyOverlay.get(key);
      return cell === 0 || cell === undefined ? '' : cell.v;
    }
    if (originalCol === undefined || originalCol < 0) return '';
    const originalRow = store.logicalToOriginalRow[sheetIdx]?.[r];
    if (originalRow === undefined || originalRow < 0) return '';
    const cacheKeyStr = `${sheetIdx}:${originalRow}`;
    const cachedRow = store.readCache.get(cacheKeyStr);
    if (!cachedRow) {
      onMiss(sheetIdx, r);
      return '';
    }
    const cell = cachedRow[originalCol];
    return cell === 0 || cell === undefined ? '' : cell.v;
  };
  const key = `${sheetIdx}:${row}:${col}`;
  const raw = rawCellTextSync(row, col);
  if (!raw.startsWith('=')) {
    if (raw.trim() === '') return '';
    const num = Number(raw);
    return Number.isFinite(num) ? num : raw;
  }
  if (memo.has(key)) return memo.get(key);
  if (evalGuard.has(key)) return FormulaError.REF.name;
  evalGuard.add(key);
  // onRange가 너무 큰 범위를 만났음을 기록하는 플래그 — fast-formula-parser의 parse()는 평가 중
  // 동기적으로 던져진 에러를 전부 FormulaError.ERROR('#ERROR!')로 재포장해버려서(원본 에러의
  // name을 보존하지 않음, node_modules/fast-formula-parser/grammar/hooks.js의 parse() catch
  // 블록 참고), onRange 안에서 FormulaError.REF를 던져도 우리 쪽 catch까지 '#REF!'로 도달하지
  // 않는다. 그래서 던지는 대신 플래그로 표시해두고, parser.parse()가 끝난 뒤(무엇을 반환했든)
  // 그 결과를 덮어써서 '#REF!'를 돌려준다.
  let rangeTooLarge = false;
  let result: unknown;
  try {
    const parser = new FormulaParser({
      functions: formulaFunctions,
      functionsNeedContext: formulaFunctionsNeedContext,
      onCell: (ref: CellRef) => computeGridCellDisplaySync(store, sheetIdx, ref.row - 1, ref.col - 1, evalGuard, memo, onMiss),
      onRange: (ref: RangeRef) => {
        const cellCount = (ref.to.row - ref.from.row + 1) * (ref.to.col - ref.from.col + 1);
        if (cellCount > MAX_RANGE_CELLS) {
          // 순회를 시작하기 전에 걸러낸다 — "너무 크다는 걸 알아챌 때까지 순회"하면 이미
          // 이벤트 루프를 막아버린 뒤라 상한을 두는 의미가 없다. 빈 배열을 돌려줘 이 범위를
          // 소비하는 함수(SUM 등)가 아무 것도 순회하지 않게 하고, 최종 결과는 아래에서 덮어쓴다.
          rangeTooLarge = true;
          return [];
        }
        const rows: unknown[][] = [];
        for (let r = ref.from.row; r <= ref.to.row; r += 1) {
          const rowVals: unknown[] = [];
          for (let c = ref.from.col; c <= ref.to.col; c += 1) {
            rowVals.push(computeGridCellDisplaySync(store, sheetIdx, r - 1, c - 1, evalGuard, memo, onMiss));
          }
          rows.push(rowVals);
        }
        return rows;
      },
    });
    const parsed = parser.parse(raw.slice(1), { row: row + 1, col: col + 1, sheet: 'Sheet1' });
    result = rangeTooLarge ? FormulaError.REF.name : parsed;
  } catch (error) {
    result = rangeTooLarge ? FormulaError.REF.name : (error instanceof FormulaError ? error.name : '#ERROR!');
  } finally {
    evalGuard.delete(key);
  }
  memo.set(key, result);
  return result;
}

/**
 * computeGridCellDisplay의 LazyGridStore 버전 — 캐시 미스가 나면(computeGridCellDisplaySync가
 * onMiss로 알려주면) 그 행을 getRow로 비동기로 채운 뒤 같은 수식 평가를 처음부터 다시 시도한다
 * (스펙 문서 "수식 계산" 절 참고). 매 시도마다 최소 한 행이 새로 캐시에 채워지므로 이 루프는
 * 반드시 끝난다.
 *
 * `memo`는 호출자(예: room.ts의 recomputeFormulasForPage)가 "재계산 패스 하나"(예: 페이지 하나)
 * 동안 여러 수식 셀에 걸쳐 공유하는 Map이다 — 체인형 수식(뒤 칸이 앞 칸을 참조)에서 앞 칸이 이미
 * 이 패스 안에서 계산됐다면 다시 계산하지 않는다(computeGridCellDisplay 문서 주석 참고, O(N²) 방지).
 * 다만 이 공유 memo에는 **미스 없이 끝난 시도의 결과만** 들어가야 한다 — 시도 도중 어딘가에서
 * 캐시 미스가 나면 그 시도 동안 계산된 값들은(미스 난 칸을 빈 칸으로 취급한 채 계산됐을 수 있어)
 * 신뢰할 수 없으므로, 그런 "더러운" 시도의 결과는 공유 memo에 절대 섞이지 않고 통째로 버려진다
 * (매 시도는 공유 memo의 스냅숏 복사본에만 쓰고, 그 시도가 깨끗이 끝났을 때만 원본에 합친다).
 */
export async function computeGridCellDisplayAsync(
  store: LazyGridStore,
  sheetIdx: number,
  row: number,
  col: number,
  evalGuard: Set<string> = new Set(),
  memo: Map<string, unknown> = new Map(),
): Promise<unknown> {
  const MAX_RETRIES = 10_000; // 한 수식이 참조할 수 있는 서로 다른 행 수의 넉넉한 상한(안전망)
  for (let attempt = 0; attempt < MAX_RETRIES; attempt += 1) {
    let missed: { sheetIdx: number; row: number } | null = null;
    // 공유 memo의 스냅숏 복사본에서 시작한다 — 이미 다른(깨끗했던) 시도/셀에서 검증된 값은 그대로
    // 재사용하고, 이번 시도에서 새로 계산되는 값은 이 복사본에만 쌓인다.
    const attemptMemo = new Map(memo);
    const result = computeGridCellDisplaySync(store, sheetIdx, row, col, new Set(evalGuard), attemptMemo, (missSheet, missRow) => {
      if (!missed) missed = { sheetIdx: missSheet, row: missRow }; // 이번 시도의 첫 미스만 채우고 재시도
    });
    if (!missed) {
      for (const [k, v] of attemptMemo) memo.set(k, v); // 깨끗이 끝났으니 새로 계산된 값들을 공유 memo에 합친다
      return result;
    }
    const toFill: { sheetIdx: number; row: number } = missed;
    await getRow(store, toFill.sheetIdx, toFill.row); // 캐시를 채운다
  }
  throw new Error(`수식 계산이 ${MAX_RETRIES}번 재시도 후에도 끝나지 않음 (문서 손상 의심)`);
}

/** computeGridCellDisplay의 결과를 GRID_COMPUTED_MAP에 저장할 문자열로 바꾼다 — 서식(쉼표·
 * 통화 등)은 클라이언트가 스타일 정보로 따로 입히므로 여기서는 순수 값만 문자열화한다. */
export function formatComputedValueForStorage(value: unknown): string {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number') return value === null || value === undefined ? '' : String(value);
  return String(Math.round(value * 1e10) / 1e10);
}
