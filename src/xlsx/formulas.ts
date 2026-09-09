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
  // onRange가 너무 큰 범위를 만났음을 기록하는 플래그 — fast-formula-parser의 parse()는 평가 중
  // 동기적으로 던져진 에러를 전부 FormulaError.ERROR('#ERROR!')로 재포장해버려서(원본 에러의
  // name을 보존하지 않음, node_modules/fast-formula-parser/grammar/hooks.js의 parse() catch
  // 블록 참고), onRange 안에서 FormulaError.REF를 던져도 우리 쪽 catch까지 '#REF!'로 도달하지
  // 않는다. 그래서 던지는 대신 플래그로 표시해두고, parser.parse()가 끝난 뒤(무엇을 반환했든)
  // 그 결과를 덮어써서 '#REF!'를 돌려준다.
  let rangeTooLarge = false;
  try {
    const parser = new FormulaParser({
      functions: formulaFunctions,
      functionsNeedContext: formulaFunctionsNeedContext,
      onCell: (ref: CellRef) => computeGridCellDisplay(grid, sheetIdx, ref.row - 1, ref.col - 1, evalGuard),
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
            rowVals.push(computeGridCellDisplay(grid, sheetIdx, r - 1, c - 1, evalGuard));
          }
          rows.push(rowVals);
        }
        return rows;
      },
    });
    const result = parser.parse(raw.slice(1), { row: row + 1, col: col + 1, sheet: 'Sheet1' });
    return rangeTooLarge ? FormulaError.REF.name : result;
  } catch (error) {
    if (rangeTooLarge) return FormulaError.REF.name;
    return error instanceof FormulaError ? error.name : '#ERROR!';
  } finally {
    evalGuard.delete(key);
  }
}

/** computeGridCellDisplay와 완전히 같은 평가 로직이지만, 데이터 소스가 GridDocument 동기
 * 인덱싱이 아니라 "이미 캐시에 채워져 있어야 하는" LazyGridStore 동기 조회다. 캐시에 없는
 * 좌표를 만나면 던지지 않고(fast-formula-parser가 동기 콜백 안의 예외를 재포장해버림 —
 * rangeTooLarge와 같은 이유) onMiss 콜백으로 알린다. */
function computeGridCellDisplaySync(
  store: LazyGridStore,
  sheetIdx: number,
  row: number,
  col: number,
  evalGuard: Set<string>,
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
  if (evalGuard.has(key)) return FormulaError.REF.name;
  evalGuard.add(key);
  // onRange가 너무 큰 범위를 만났음을 기록하는 플래그 — fast-formula-parser의 parse()는 평가 중
  // 동기적으로 던져진 에러를 전부 FormulaError.ERROR('#ERROR!')로 재포장해버려서(원본 에러의
  // name을 보존하지 않음, node_modules/fast-formula-parser/grammar/hooks.js의 parse() catch
  // 블록 참고), onRange 안에서 FormulaError.REF를 던져도 우리 쪽 catch까지 '#REF!'로 도달하지
  // 않는다. 그래서 던지는 대신 플래그로 표시해두고, parser.parse()가 끝난 뒤(무엇을 반환했든)
  // 그 결과를 덮어써서 '#REF!'를 돌려준다.
  let rangeTooLarge = false;
  try {
    const parser = new FormulaParser({
      functions: formulaFunctions,
      functionsNeedContext: formulaFunctionsNeedContext,
      onCell: (ref: CellRef) => computeGridCellDisplaySync(store, sheetIdx, ref.row - 1, ref.col - 1, evalGuard, onMiss),
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
            rowVals.push(computeGridCellDisplaySync(store, sheetIdx, r - 1, c - 1, evalGuard, onMiss));
          }
          rows.push(rowVals);
        }
        return rows;
      },
    });
    const result = parser.parse(raw.slice(1), { row: row + 1, col: col + 1, sheet: 'Sheet1' });
    return rangeTooLarge ? FormulaError.REF.name : result;
  } catch (error) {
    if (rangeTooLarge) return FormulaError.REF.name;
    return error instanceof FormulaError ? error.name : '#ERROR!';
  } finally {
    evalGuard.delete(key);
  }
}

/**
 * computeGridCellDisplay의 LazyGridStore 버전 — 캐시 미스가 나면(computeGridCellDisplaySync가
 * onMiss로 알려주면) 그 행을 getRow로 비동기로 채운 뒤 같은 수식 평가를 처음부터 다시 시도한다
 * (스펙 문서 "수식 계산" 절 참고). 매 시도마다 최소 한 행이 새로 캐시에 채워지므로 이 루프는
 * 반드시 끝난다.
 */
export async function computeGridCellDisplayAsync(
  store: LazyGridStore,
  sheetIdx: number,
  row: number,
  col: number,
  evalGuard: Set<string> = new Set(),
): Promise<unknown> {
  const MAX_RETRIES = 10_000; // 한 수식이 참조할 수 있는 서로 다른 행 수의 넉넉한 상한(안전망)
  for (let attempt = 0; attempt < MAX_RETRIES; attempt += 1) {
    let missed: { sheetIdx: number; row: number } | null = null;
    const result = computeGridCellDisplaySync(store, sheetIdx, row, col, new Set(evalGuard), (missSheet, missRow) => {
      if (!missed) missed = { sheetIdx: missSheet, row: missRow }; // 이번 시도의 첫 미스만 채우고 재시도
    });
    if (!missed) return result;
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
