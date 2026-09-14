import FormulaParser, { type CellRef, type RangeRef } from 'fast-formula-parser';
import { formulaFunctions, formulaFunctionsNeedContext } from '../client/formulaOverrides.js';

/**
 * `new FormulaParser({...})`를 셀 하나 계산할 때마다 새로 만드는 비용이 사실상 전부다 — 실측(리뷰
 * 참고) 결과 인스턴스 하나를 생성하는 데 걸리는 시간이 ~1.9ms이고, 이게 전체 계산 비용의 99%를
 * 차지한다(생성자가 ~10개의 함수 라이브러리를 `Object.assign`으로 합치고, `Object.keys()`를 여러 번
 * 돌려 조회 테이블을 다시 만들고, chevrotain 문법 파서를 새로 인스턴스화하기 때문 — 이 전부가 수식
 * 내용과 무관하게 매번 반복된다). 체인형 수식(뒤 칸이 앞 칸을 참조)에는 formulas.ts의 memo가
 * 이미 효과가 있지만(같은 칸을 두 번 계산 안 함), *독립된* 수식 셀(칸끼리 서로 참조하지 않는 경우 —
 * 예: `=A{r}*B{r}`을 수천 행에 반복)은 각 칸이 정확히 한 번씩만 계산되므로 memo가 전혀 도움이 안
 * 되고, 매 칸이 그대로 이 ~1.9ms 생성 비용을 다 문다.
 *
 * **왜 인스턴스를 공유하면 안 되는가(d2bff59가 이미 확인한 것)**: `parse()`는 인스턴스 상태
 * (`this.position`, `this.parser.input`)를 호출마다 덮어쓴다. 수식 셀이 다른 수식 셀을 참조하면
 * 바깥 `parse()` 호출이 끝나기 전에 안쪽 `parse()` 호출이 같은 인스턴스에서 시작되므로(재진입),
 * 인스턴스 하나를 공유하면 안쪽 호출이 바깥 호출의 `position`/토큰 스트림을 덮어써 버려 바깥
 * 호출이 이어서 읽는 좌표/토큰이 안쪽 호출의 것으로 오염된다.
 *
 * **풀(pool)이 안전한 이유**: 수식 셀 중첩(재귀 평가)은 반드시 동기적이고 LIFO다 — 안쪽 `parse()`가
 * 끝나야 바깥 `parse()`가 이어서 진행된다. 그래서 "중첩 단계마다 서로 다른 인스턴스를 하나씩 꺼내
 * 쓴다"만 지키면, 어느 두 `parse()` 호출도 "둘 다 아직 안 끝난 채로" 같은 인스턴스를 쓰는 일이
 * 없다. 이 모듈은 그 규칙을 강제하는 아주 얇은 계층이다:
 *
 * - 인스턴스는 모듈 레벨 free-list(빈 인스턴스 배열)에서 재사용되고, 생성자는 인스턴스 하나당 딱
 *   한 번만 불린다(매 호출마다가 아니라).
 * - 각 인스턴스는 "고정된(stable)" onCell/onRange 콜백(트램폴린)으로 만들어진다 — 이 콜백은 그
 *   인스턴스 자신의 실제 로직을 담지 않고, 그때그때 체크아웃된 "컨텍스트"(`entry.ctx`)로 그대로
 *   위임만 한다. `entry.ctx`는 `withPooledParser`가 매 호출 직전에 설정하고, 끝나면(성공이든 예외든)
 *   `finally`에서 지운다.
 * - "체크아웃" = free-list에서 하나 꺼내거나(없으면 새로 생성) → 컨텍스트 설정 → 호출자의 콜백
 *   실행 → (반드시 finally에서) 컨텍스트 해제 + free-list에 반납.
 * - 중첩된 `parse()` 호출(수식 셀이 다른 수식 셀을 참조해 onCell/onRange 콜백 "안에서" 다시
 *   `withPooledParser`를 부르는 경우)은 그 시점에 바깥 인스턴스가 아직 체크아웃된 채라 free-list에
 *   없으므로, 자연히 "다른" 인스턴스를 꺼내게 된다(재귀가 스택 프레임을 공유하지 않는 것과 같은
 *   이유) — 그래서 두 개의 "아직 안 끝난" `parse()` 호출이 같은 인스턴스나 그 가변 상태를 공유하는
 *   일이 원천적으로 없다.
 */

interface PooledParserCallbacks {
  onCell: (ref: CellRef) => unknown;
  onRange: (ref: RangeRef) => unknown[][];
}

interface PoolEntry {
  parser: FormulaParser;
  /** 현재 체크아웃 중인 호출의 실제 콜백 — 체크아웃돼 있지 않을 때는 반드시 null이어야 한다. */
  ctx: PooledParserCallbacks | null;
}

/** 반납된(현재 아무도 안 쓰는) 인스턴스들의 free-list. 상한 없음 — 동시에 체크아웃될 수 있는 최대
 * 개수는 "동시에 진행 중인 중첩 parse() 호출 깊이"뿐이라, 사실상 수식 하나의 참조 중첩 깊이로
 * 자연히 제한된다(수만 개가 쌓일 시나리오가 없다). */
const freeList: PoolEntry[] = [];

function createPoolEntry(): PoolEntry {
  // entry 자신을 클로저로 캡처해서, 트램폴린 콜백이 항상 "이 인스턴스에 지금 체크아웃돼 있는"
  // 컨텍스트로 위임하게 한다 — entry.parser가 아직 만들어지기 전에 콜백부터 정의해야 하므로
  // entry 객체를 먼저 만들고 나중에 parser 필드를 채운다.
  const entry: PoolEntry = { parser: null as unknown as FormulaParser, ctx: null };
  entry.parser = new FormulaParser({
    functions: formulaFunctions,
    functionsNeedContext: formulaFunctionsNeedContext,
    onCell: (ref: CellRef) => {
      if (!entry.ctx) {
        // 풀 사용 규칙(체크아웃 중에만 콜백이 불림)이 깨졌다는 뜻 — 수식 자체의 문제가 아니라
        // 이 모듈의 버그이므로, 조용히 이상한 값을 돌려주는 대신 바로 드러나게 던진다.
        throw new Error('pooled FormulaParser onCell invoked with no active checkout');
      }
      return entry.ctx.onCell(ref);
    },
    onRange: (ref: RangeRef) => {
      if (!entry.ctx) {
        throw new Error('pooled FormulaParser onRange invoked with no active checkout');
      }
      return entry.ctx.onRange(ref);
    },
  });
  return entry;
}

/**
 * 풀에서 인스턴스 하나를 체크아웃해 `fn`에 넘겨주고, `fn`이 끝나면(정상 반환이든 예외든) 반드시
 * free-list로 반납한다. 호출자는 풀 상태를 직접 다룰 필요 없이 이 함수만 부르면 된다 — `fn` 안에서
 * 던진 예외도 그대로 다시 던져지되(삼키지 않음), 반납은 항상 일어난다.
 *
 * `callbacks`는 이번 호출 한 번(과 그 안에서 동기적으로 끝나는 `parser.parse()` 한 번)에만 유효한
 * onCell/onRange다 — `fn` 밖에서 나중에(비동기로) 호출되면 안 된다. `parse()`는 항상 동기 함수이고
 * 이 프로젝트는 이 함수를 오직 동기 `parse()` 호출을 감싸는 데만 쓰므로 이 전제는 항상 성립한다.
 */
export function withPooledParser<T>(callbacks: PooledParserCallbacks, fn: (parser: FormulaParser) => T): T {
  const entry = freeList.pop() ?? createPoolEntry();
  entry.ctx = callbacks;
  try {
    return fn(entry.parser);
  } finally {
    entry.ctx = null;
    freeList.push(entry);
  }
}
