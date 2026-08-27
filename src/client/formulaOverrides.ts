import {
  Criteria,
  FormulaError,
  FormulaHelpers,
  Types,
  WildCard,
  type ContextFormulaFunction,
  type FormulaContext,
  type FormulaFunction,
  type ParsedCriteria,
} from 'fast-formula-parser';

/**
 * fast-formula-parser(MIT)는 VLOOKUP/HLOOKUP/INDEX는 실제로 구현돼 있지만, README의 "280개 함수
 * 지원" 목록과 달리 MAX/MIN/MATCH/CHOOSE/SWITCH/SUMIFS/AVERAGEIFS를 포함한 32개는 본문이 텅 빈
 * 스텁이다(node_modules 소스를 직접 열어 확인함). 그중 실무에서 자주 쓰는 것들만 여기서 직접
 * 구현해 `functions`/`functionsNeedContext` 옵션으로 라이브러리 내장 구현을 덮어쓴다.
 * (LARGE/SMALL/MEDIAN/MAXIFS/MINIFS/TEXTJOIN/SUBSTITUTE 등 나머지 스텁은 v1 범위 밖으로 남겨둔다.)
 */

const H = FormulaHelpers;

function accept(value: unknown, type?: unknown, defValue?: unknown, flat = true): any {
  return H.accept(value, type, defValue, flat);
}

/** context가 필요 없는(값 인자만 받는) 커스텀 함수들 — MAX/MIN/MATCH/CHOOSE/SWITCH */
export const formulaFunctions: Record<string, FormulaFunction> = {
  MAX: (...numbers: unknown[]) => {
    let max = -Infinity;
    H.flattenParams(numbers, Types.NUMBER, true, (item) => {
      if (typeof item === 'number' && item > max) max = item;
    });
    return max === -Infinity ? 0 : max;
  },

  MIN: (...numbers: unknown[]) => {
    let min = Infinity;
    H.flattenParams(numbers, Types.NUMBER, true, (item) => {
      if (typeof item === 'number' && item < min) min = item;
    });
    return min === Infinity ? 0 : min;
  },

  SWITCH: (expression: unknown, ...rest: unknown[]) => {
    const expr = accept(expression);
    let i = 0;
    for (; i + 1 < rest.length; i += 2) {
      if (accept(rest[i]) === expr) return accept(rest[i + 1]);
    }
    if (i < rest.length) return accept(rest[i]); // 짝 안 맞는 마지막 인자 = 기본값
    throw FormulaError.NA;
  },

  // Excel MATCH(lookup_value, lookup_array, [match_type]) — match_type: 0=정확히 일치(가장 흔한
  // 사용법, INDEX와 짝지어 씀), 1=오름차순 중 lookup_value 이하 최대값, -1=내림차순 중 이상 최소값.
  MATCH: (lookupValue: unknown, lookupArray: unknown, matchType?: unknown) => {
    const value = accept(lookupValue);
    let table: unknown[][];
    try {
      // flat=false — VLOOKUP과 동일하게 2차원 구조를 유지해야 아래 row.forEach가 동작한다
      table = accept(lookupArray, Types.ARRAY, undefined, false) as unknown[][];
    } catch (e) {
      if (e instanceof FormulaError) throw FormulaError.NA;
      throw e;
    }
    const type = matchType === undefined ? 1 : Math.trunc(accept(matchType, Types.NUMBER, 1));
    const flat: unknown[] = [];
    table.forEach((row) => row.forEach((cell) => flat.push(cell)));

    const lookupType = typeof value;
    if (type === 0) {
      let index = -1;
      if (WildCard.isWildCard(value)) {
        const re = WildCard.toRegex(value, 'i');
        index = flat.findIndex((item) => re.test(String(item)));
      } else {
        index = flat.findIndex((item) => item === value);
      }
      if (index === -1) throw FormulaError.NA;
      return index + 1;
    }
    let bestIndex = -1;
    for (let i = 0; i < flat.length; i += 1) {
      const curr = flat[i];
      if (typeof curr !== lookupType) continue;
      const qualifies = type > 0 ? (curr as number) <= (value as number) : (curr as number) >= (value as number);
      if (qualifies) bestIndex = i;
      else if (type > 0) break; // 오름차순 가정 — 기준을 벗어나면 더 볼 필요 없음
    }
    if (bestIndex === -1) throw FormulaError.NA;
    return bestIndex + 1;
  },
};

function matchesCriteria(value: unknown, criteria: ParsedCriteria): boolean {
  if (criteria.op === 'wc') return criteria.match === (criteria.value as RegExp).test(String(value ?? ''));
  const a = value;
  const b = criteria.value;
  switch (criteria.op) {
    case '=':
      return a === b;
    case '<>':
      return a !== b;
    case '>':
      return typeof a === 'number' && typeof b === 'number' && a > b;
    case '<':
      return typeof a === 'number' && typeof b === 'number' && a < b;
    case '>=':
      return typeof a === 'number' && typeof b === 'number' && a >= b;
    case '<=':
      return typeof a === 'number' && typeof b === 'number' && a <= b;
    default:
      return false;
  }
}

// SUMIFS/AVERAGEIFS는 라이브러리 내부 funsNeedContextAndNoDataRetrieve 목록에 없어(SUMIF/AVERAGEIF와
// 달리) 파서가 이미 참조를 실제 값으로 해석한 {value, isArray, ...} 형태로 넘겨준다 — H.retrieveArg로
// 한 번 더 해석하려 하면(raw ref를 기대하는 API라 .ref가 없어) 값이 깨진다. 그냥 accept()만 하면 된다.
function resolveRangeArg(_context: FormulaContext, arg: unknown): unknown[][] {
  // flat=false — SUMIFS/AVERAGEIFS는 행/열 좌표가 맞아떨어져야 하므로 2차원 구조를 유지해야 한다
  return accept(arg, Types.ARRAY, undefined, false) as unknown[][];
}

function parseCriteriaArg(_context: FormulaContext, arg: unknown): ParsedCriteria {
  return Criteria.parse(accept(arg));
}

function sumIfsLike(
  targetRange: unknown[][],
  criteriaGroups: readonly { range: unknown[][]; criteria: ParsedCriteria }[],
): { sum: number; count: number } {
  let sum = 0;
  let count = 0;
  targetRange.forEach((row, r) => {
    row.forEach((value, c) => {
      const matches = criteriaGroups.every((group) => matchesCriteria(group.range[r]?.[c], group.criteria));
      if (matches && typeof value === 'number') {
        sum += value;
        count += 1;
      }
    });
  });
  return { sum, count };
}

/** context(현재 셀 위치 등 파서 상태)가 필요한 커스텀 함수들 — SUMIFS/AVERAGEIFS(다중 조건 범위 상관) */
export const formulaFunctionsNeedContext: Record<string, ContextFormulaFunction> = {
  // CHOOSE는 라이브러리 내부에서 이름으로 하드코딩되어(hooks.js의 funsNeedContext 목록) 항상 context를
  // 첫 인자로 받는다 — functions에 등록하면 context가 indexNum 자리로 밀려 들어가 깨진다.
  CHOOSE: (_context: FormulaContext, indexNum: unknown, ...values: unknown[]) => {
    const idx = Math.trunc(accept(indexNum, Types.NUMBER));
    if (idx < 1 || idx > values.length) throw FormulaError.VALUE;
    return accept(values[idx - 1]);
  },

  SUMIFS: (context: FormulaContext, sumRange: unknown, ...criteriaPairs: unknown[]) => {
    const target = resolveRangeArg(context, sumRange);
    const groups = [];
    for (let i = 0; i + 1 < criteriaPairs.length; i += 2) {
      groups.push({ range: resolveRangeArg(context, criteriaPairs[i]), criteria: parseCriteriaArg(context, criteriaPairs[i + 1]) });
    }
    return sumIfsLike(target, groups).sum;
  },

  AVERAGEIFS: (context: FormulaContext, avgRange: unknown, ...criteriaPairs: unknown[]) => {
    const target = resolveRangeArg(context, avgRange);
    const groups = [];
    for (let i = 0; i + 1 < criteriaPairs.length; i += 2) {
      groups.push({ range: resolveRangeArg(context, criteriaPairs[i]), criteria: parseCriteriaArg(context, criteriaPairs[i + 1]) });
    }
    const { sum, count } = sumIfsLike(target, groups);
    if (count === 0) throw FormulaError.DIV0;
    return sum / count;
  },
};
