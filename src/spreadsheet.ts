import ExcelJS from 'exceljs';
import type { DocumentStore } from './storage.js';
import type { DocumentMeta } from './types.js';
import { buildSheetNameIndex, computeGridCellDisplay } from './xlsx/formulas.js';

/** 대형 시트를 훑는 동기 반복문 중간에 이벤트 루프에 제어권을 돌려준다 — 이게 없으면 큰 xlsx를
 * 열 때(파싱이든 동시편집 시딩이든) 그 몇 초~몇십 초 동안 Node가 단일 스레드라 그 문서를 연
 * 사용자뿐 아니라 그 순간 서버에 접속한 모든 사용자·모든 다른 문서 요청까지 함께 멈춘다. */
export function yieldToEventLoop(): Promise<void> {
  return new Promise((resolve) => setImmediate(resolve));
}

/** 이 개수의 행마다 한 번씩 yieldToEventLoop를 호출한다 — 너무 자주 양보하면(예: 매 행) setImmediate
 * 오버헤드 자체가 커지고, 너무 뜸하면 양보 효과가 없다. 2000행은 실측 없이 정한 보수적인 값이라
 * 필요하면 나중에 조정 가능하다. */
export const YIELD_EVERY_N_ROWS = 2000;

/** parseXlsxToGrid의 수식 계산 패스가 이 개수의 **수식 셀**마다 한 번씩 양보한다 — 행이 아니라
 * 셀 단위인 이유는, 그 패스가 도는 대상이 "모든 행"이 아니라 "수식이 들어 있는 칸만 모은 목록"이라
 * 행 기준 간격으로는 (수식이 한 행에 몰려 있을 때) 아무 효과가 없을 수 있기 때문이다. 수식 한 칸의
 * 계산 비용은 참조 범위 크기에 비례해 최대 MAX_RANGE_CELLS(xlsx/formulas.ts)까지 갈 수 있으므로,
 * 행 단위(2000)보다 훨씬 촘촘하게 잡아 대형 파일 업로드 한 번이 서버 전체를 멈추지 않게 한다
 * (room.ts의 RECOMPUTE_YIELD_EVERY_N_ROWS가 같은 이유로 페이지 크기의 1/4을 쓰는 것과 같은 취지). */
export const YIELD_EVERY_N_FORMULA_CELLS = 500;

/** 그리드 뷰 데이터 모델 — 원본 파일 대신 표시용 값·스타일만 담는다. */
export interface GridCell {
  readonly v: string;
  readonly s?: number;
}

export interface GridMerge {
  readonly r: number;
  readonly c: number;
  readonly rs: number;
  readonly cs: number;
}

export interface GridSheet {
  readonly name: string;
  readonly rowCount: number;
  readonly colCount: number;
  readonly colWidths: readonly number[];
  readonly merges: readonly GridMerge[];
  /** 엑셀 "틀 고정" — 위 N행/왼쪽 N열은 스크롤해도 계속 보인다. 없으면 둘 다 0. */
  readonly freezeRows: number;
  readonly freezeCols: number;
  readonly rows: ReadonlyArray<ReadonlyArray<GridCell | 0>>;
}

/** 테두리 굵기 — 1=가는 선, 2=굵은 선(엑셀 thin/medium을 단순화. 그 외 스타일은 가장 가까운 쪽으로 반올림). */
export type BorderWeight = 1 | 2;

/** 숫자 표시 형식 종류 — gridShared.ts의 CellStyleAttrs와 동일한 단순화(원본 서식 코드를 몇 가지
 * 자주 쓰는 범주로 분류). */
export type NumberFormatKind = 'number' | 'currency' | 'percent';

export interface GridStyle {
  readonly b?: 1;
  readonly i?: 1;
  readonly u?: 1;
  readonly st?: 1;
  /** 글자 크기(px). 기본(11pt≈15px)과 다를 때만 기록 */
  readonly fs?: number;
  /** 글꼴 이름. 기본 글꼴과 같으면 기록하지 않는다. */
  readonly ff?: string;
  readonly col?: string;
  readonly bg?: string;
  readonly a?: 'l' | 'c' | 'r';
  /** 세로 정렬 — 위/가운데/아래. */
  readonly va?: 't' | 'm' | 'b';
  readonly w?: 1;
  /** 숫자 표시 형식(쉼표/통화/백분율) + 소수 자릿수. */
  readonly nf?: NumberFormatKind;
  readonly nd?: number;
  /** 시트 보호가 걸려 있고 이 셀이 잠겨 있음(=편집 불가) */
  readonly lk?: 1;
  /** 엑셀 데이터 유효성 검사(정수/소수)가 걸려 있음 */
  readonly nu?: 1;
  readonly bt?: BorderWeight;
  readonly br?: BorderWeight;
  readonly bb?: BorderWeight;
  readonly bl?: BorderWeight;
}

export interface GridDocument {
  readonly sheets: readonly GridSheet[];
  readonly styles: readonly GridStyle[];
}

/**
 * 캐시(derived kind 'grid') 포맷 버전 — 그리드 "구조"(JSONL 줄 모양)가 아니라, 그 안의 셀 값을
 * "어떻게 계산했는지"가 바뀔 때마다 올린다. 지금 값(2)은 xlsx/formulas.ts의 시트 간 참조(cross-sheet
 * reference) 버그 수정에 대응한다 — 그 수정 전에 만들어진 캐시는(버전 필드가 아예 없다, 옛
 * serializeGrid가 안 써 넣었으므로) 시트를 넘나드는 수식이 조용히 틀린 값(또는 계산 전이라면
 * 원문 그대로)으로 저장돼 있을 수 있다. runGridExtraction(routes/documents.ts)이 이 값을 읽어
 * "다음 조회 시점에" 낡은 캐시를 감지해 원본에서 다시 계산하도록 하는 lazy self-heal의 기준선이다
 * (routes/viewer.ts의 loadGridUncached 참고) — 기존 문서를 한꺼번에 마이그레이션하지 않는다(그 자체가
 * 새로운 리스크이므로). 앞으로 계산 로직이 바뀔 때마다(예: 새로운 함수 지원, 다른 계산 버그 수정)
 * 이 값을 올리면 같은 메커니즘이 그대로 재사용된다. */
export const GRID_CACHE_FORMAT_VERSION = 2;

/** JSONL 헤더(1행): 셀 데이터를 제외한 시트 메타 정보 */
export interface GridHeader {
  /** GRID_CACHE_FORMAT_VERSION 문서 주석 참고. 이 필드가 없거나(구버전 캐시) 현재 버전과 다르면
   * "낡은 캐시"로 취급한다. */
  readonly formatVersion?: number;
  readonly styles: readonly GridStyle[];
  readonly sheets: ReadonlyArray<Omit<GridSheet, 'rows'>>;
}

/**
 * 그리드를 JSONL로 직렬화한다: 1행 = 헤더, 이후 시트 순서대로 행당 1줄.
 * 수십만 행 문서를 요청 범위만 파싱해 서빙할 수 있도록 하기 위함이다.
 */
export function serializeGrid(grid: GridDocument): Buffer {
  const header: GridHeader = {
    formatVersion: GRID_CACHE_FORMAT_VERSION,
    styles: grid.styles,
    sheets: grid.sheets.map(({ rows: _rows, ...meta }) => meta),
  };
  const parts: string[] = [JSON.stringify(header)];
  for (const sheet of grid.sheets) {
    for (const row of sheet.rows) {
      parts.push(JSON.stringify(row));
    }
  }
  return Buffer.from(parts.join('\n'), 'utf8');
}

/**
 * JSONL 버퍼에서 각 줄이 시작하는 바이트 오프셋 배열을 만든다(0번째는 항상 0 = 헤더 시작줄).
 * src/routes/viewer.ts의 그리드 조회(sliceRows/searchGrid)가 필요한 범위만 파싱할 수 있게
 * 해주는 인덱스 — 수십만 행이면 이 스캔 자체가 O(파일 크기)라, 그리드를 만드는 시점에 미리
 * 계산해 saveGridDerived로 함께 저장해 두면 조회 시점에 매번 다시 스캔하지 않아도 된다.
 */
export function computeLineOffsets(buffer: Buffer): readonly number[] {
  const lineOffsets: number[] = [0];
  for (let i = buffer.indexOf(10); i !== -1; i = buffer.indexOf(10, i + 1)) {
    lineOffsets.push(i + 1);
  }
  return lineOffsets;
}

/**
 * 그리드를 JSONL로 직렬화해 저장하고(derived kind 'grid'), 같은 바이트로 만든 줄 오프셋
 * 인덱스도 함께 저장한다(derived kind 'gridindex'). 업로드 시 최초 추출
 * (routes/documents.ts의 runGridExtraction)과 동시편집 저장 후 재생성(collab/room.ts)
 * 양쪽에서 재사용해, 그리드 내용이 바뀔 때마다 인덱스가 항상 최신 상태로 유지되게 한다
 * (그리드가 바뀌지 않는 한 인덱스도 따로 무효화할 필요가 없다).
 *
 * **알려진, 좁은 쓰기 경합(리뷰 지적, 구조적으로 고치지 않기로 함)**: 아래 두 `store.saveDerived`
 * 호출은 서로 독립적이고(문서 하나당 락이 없다), storage.ts의 saveDerived 자체도 saveOriginal과
 * 달리 temp-file+rename의 원자적 쓰기가 아니라 `fs.writeFile`로 그 경로에 직접 쓴다. 그래서
 * 같은 문서에 대해 이 함수가 **동시에 두 번** 호출되면(예: collab/room.ts의 flushToStore가 이
 * 문서를 유휴 자동저장하는 바로 그 순간, routes/viewer.ts의 regenerateStaleGridCache가 같은
 * 문서를 낡은 포맷 캐시로 감지해 재생성을 트리거하는 경우) 두 호출의 grid/gridindex 쓰기가
 * 시간순으로 뒤섞여, "한쪽의 grid + 다른 쪽의 gridindex"처럼 서로 짝이 안 맞는 조합이 잠깐 disk에
 * 남을 수 있다 — 이 조합이 우연히 routes/viewer.ts의 offsetsLookValid 같은 구조적 정합성 검사를
 * (진짜 내용은 안 맞아도 통과할 만큼) 우연히 통과하면, 다음 성공적인 저장 전까지 일시적으로 깨진
 * 파싱 결과를 서빙할 수 있다.
 *
 * 이 경합은 좁다 — "이 문서가 포맷이 낡은 grid 캐시를 갖고 있으면서, 동시에 바로 그 순간 실제로
 * 동시편집 중이어야" 하는 아주 구체적인 타이밍이 겹쳐야 하고, 어느 한쪽이든 다음에 성공적으로
 * (한쪽만) 저장하면 다시 짝이 맞는 grid/gridindex 쌍이 돼 저절로 회복된다(자가치유) — 영구적인
 * 데이터 손실이 아니다. 제대로 고치려면 문서 단위의 실제 프로세스 간/요청 간 쓰기 조정(락)이나
 * 원자적 다중 파일 쓰기 스킴이 필요한데, 이 정도 코드 변경은 이 함수의 범위를 넘어선다 — 지금은
 * 이 주석으로만 남겨 다음에 이 경로를 건드리는 사람이 바로 알 수 있게 해둔다.
 */
export async function saveGridDerived(
  store: Pick<DocumentStore, 'saveDerived'>,
  docId: string,
  grid: GridDocument,
): Promise<void> {
  const serialized = serializeGrid(grid);
  await store.saveDerived(docId, 'grid', serialized);
  await store.saveDerived(
    docId,
    'gridindex',
    Buffer.from(JSON.stringify({ lineOffsets: computeLineOffsets(serialized) })),
  );
}

/**
 * 문서 원본(xlsx/csv)을 읽어 GridDocument로 만든다 — routes/documents.ts의 최초 업로드 추출
 * (runGridExtraction)과 routes/viewer.ts의 낡은 캐시 self-heal(loadGridUncached) 양쪽이 공유하는
 * "확장자별로 어떤 파서를 쓸지"의 단일 소스. store를 인자로 받는 이유는 saveGridDerived와 같다 —
 * 이 모듈이 storage.ts를 값으로(순환 없이) 참조하지 않게 하기 위함. */
export async function buildGridFromOriginal(
  store: Pick<DocumentStore, 'readOriginal'>,
  meta: Pick<DocumentMeta, 'id' | 'ext'>,
): Promise<GridDocument> {
  const original = await store.readOriginal(meta.id, meta.ext);
  if (meta.ext === 'csv') {
    return parseCsvToGrid(original.toString('utf8'));
  }
  return parseXlsxToGrid(original);
}

/**
 * xlsx 파싱이 그리드로 받아들이는 최대 행/열 — 이 코드베이스가 "시트 하나의 현실적인 크기"로
 * 인정하는 유일한 기준선이다. 동시편집 서버(collab/room.ts)의 canonical 그리드 성장 상한도
 * 반드시 이 값을 그대로 써야 한다: 파서가 절대 만들어내지 못하는 크기를 canonical만 허용하면,
 * 조작된 소켓 메시지 하나로 파일에서는 있을 수 없는 크기의 배열을 서버에 할당시킬 수 있다
 * (YIELD_EVERY_N_ROWS를 여기서 export해 서버/클라이언트가 공유하는 것과 같은 이유).
 */
export const MAX_ROWS = 300_000;
export const MAX_COLS = 120;
const DEFAULT_COL_WIDTH_PX = 100;
const CHAR_TO_PX = 7.2;

function argbToCss(argb?: string): string | undefined {
  if (!argb) return undefined;
  // 대부분의 엑셀 파일(과 ExcelJS 자신)은 8자리 AARRGGBB로 쓰지만, POI는 알파 없이 6자리
  // RRGGBB만 쓰는 경우가 있다 — 두 형태 모두 지원해야 색이 조용히 누락되지 않는다.
  if (argb.length === 8) return `#${argb.slice(2)}`;
  if (argb.length === 6) return `#${argb}`;
  return undefined;
}

/**
 * 테마 색 인덱스 순서 (Excel은 dk1/lt1, dk2/lt2가 스왑된 관례를 사용한다)
 */
const THEME_COLOR_ORDER = [
  'lt1', 'dk1', 'lt2', 'dk2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
] as const;

/** 워크북 theme1.xml에서 테마 팔레트(RRGGBB)를 추출한다 */
export function parseThemeColors(themeXml: string | undefined): readonly string[] {
  if (!themeXml) return [];
  const scheme = themeXml.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/)?.[0];
  if (!scheme) return [];
  const byName = new Map<string, string>();
  for (const match of scheme.matchAll(
    /<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*lastClr="([0-9A-Fa-f]{6})")/g,
  )) {
    byName.set(match[1], (match[2] ?? match[3]).toUpperCase());
  }
  return THEME_COLOR_ORDER.map((name) => byName.get(name) ?? '000000');
}

/** Excel tint: HSL 명도(L)를 tint>0이면 흰색 쪽, tint<0이면 검정 쪽으로 보간한다 */
function applyTint(hex: string, tint: number): string {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  const l = (max + min) / 2;
  const d = max - min;
  let s = 0;
  if (d !== 0) {
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  const newL = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hueToRgb = (p: number, q: number, t: number) => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };
  let r2: number;
  let g2: number;
  let b2: number;
  if (s === 0) {
    r2 = g2 = b2 = newL;
  } else {
    const q = newL < 0.5 ? newL * (1 + s) : newL + s - newL * s;
    const p = 2 * newL - q;
    r2 = hueToRgb(p, q, h + 1 / 3);
    g2 = hueToRgb(p, q, h);
    b2 = hueToRgb(p, q, h - 1 / 3);
  }
  const toHex = (v: number) => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, '0');
  return `${toHex(r2)}${toHex(g2)}${toHex(b2)}`.toUpperCase();
}

/** 레거시 인덱스 팔레트 (표준 0~63) */
const INDEXED_COLORS: Record<number, string> = {
  0: '000000', 1: 'FFFFFF', 2: 'FF0000', 3: '00FF00', 4: '0000FF', 5: 'FFFF00',
  6: 'FF00FF', 7: '00FFFF', 8: '000000', 9: 'FFFFFF', 10: 'FF0000', 11: '00FF00',
  12: '0000FF', 13: 'FFFF00', 14: 'FF00FF', 15: '00FFFF', 16: '800000', 17: '008000',
  18: '000080', 19: '808000', 20: '800080', 21: '008080', 22: 'C0C0C0', 23: '808080',
  24: '9999FF', 25: '993366', 26: 'FFFFCC', 27: 'CCFFFF', 28: '660066', 29: 'FF8080',
  30: '0066CC', 31: 'CCCCFF', 32: '000080', 33: 'FF00FF', 34: 'FFFF00', 35: '00FFFF',
  36: '800080', 37: '800000', 38: '008080', 39: '0000FF', 40: '00CCFF', 41: 'CCFFFF',
  42: 'CCFFCC', 43: 'FFFF99', 44: '99CCFF', 45: 'FF99CC', 46: 'CC99FF', 47: 'FFCC99',
  48: '3366FF', 49: '33CCCC', 50: '99CC00', 51: 'FFCC00', 52: 'FF9900', 53: 'FF6600',
  54: '666699', 55: '969696', 56: '003366', 57: '339966', 58: '003300', 59: '333300',
  60: '993300', 61: '993366', 62: '333399', 63: '333333',
};

interface ExcelColorLike {
  readonly argb?: string;
  readonly theme?: number;
  readonly tint?: number;
  readonly indexed?: number;
}

/** argb / theme+tint / indexed 세 가지 색 표현을 모두 CSS 색으로 해석한다 */
function resolveColor(
  color: ExcelColorLike | undefined,
  themeColors: readonly string[],
): string | undefined {
  if (!color) return undefined;
  if (color.argb) return argbToCss(color.argb);
  if (color.theme !== undefined) {
    const base = themeColors[color.theme];
    if (!base) return undefined;
    const withTint = color.tint ? applyTint(base, color.tint) : base;
    return `#${withTint}`;
  }
  if (color.indexed !== undefined) {
    const hex = INDEXED_COLORS[color.indexed];
    return hex ? `#${hex}` : undefined;
  }
  return undefined;
}

function formatNumber(value: number, numFmt?: string): string {
  if (!Number.isFinite(value)) return String(value);
  if (numFmt && /[#0],/.test(numFmt)) {
    return value.toLocaleString('ko-KR', { maximumFractionDigits: 6 });
  }
  return String(value);
}

function formatDate(value: Date): string {
  const y = value.getUTCFullYear();
  const m = String(value.getUTCMonth() + 1).padStart(2, '0');
  const d = String(value.getUTCDate()).padStart(2, '0');
  const hh = value.getUTCHours();
  const mm = value.getUTCMinutes();
  return hh === 0 && mm === 0
    ? `${y}-${m}-${d}`
    : `${y}-${m}-${d} ${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}

export function displayValue(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return formatDate(value);
  // 리터럴(수식이 아닌) 논리값 칸 — typeof는 'object'가 아니라 'boolean'이라 아래 object 분기에
  // 걸리지 않고 이 체크가 없으면 맨 아래 String(value)로 떨어져 소문자 "true"/"false"가 된다.
  // formatFormulaResult(수식 계산 결과용, 이 파일 아래쪽)는 이미 대문자 "TRUE"/"FALSE"를 쓰고
  // 실제 엑셀도 논리값 칸을 대문자로 보여주므로, 여기서도 같은 규칙으로 맞춘다 — 그래야
  // isSpillEligible(viewer.js/grid-editor.ts)의 rawText==='TRUE'||'FALSE' 특례가 계산된 수식
  // 결과뿐 아니라 리터럴 논리값 칸에도 똑같이 걸린다(이전엔 대소문자가 달라 리터럴 칸에는
  // 전혀 걸리지 않았다).
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') return formatNumber(value, cell.numFmt);
  if (typeof value === 'object') {
    const rich = value as ExcelJS.CellRichTextValue;
    if (Array.isArray(rich.richText)) return rich.richText.map((run) => run.text).join('');
    const link = value as ExcelJS.CellHyperlinkValue;
    if (typeof link.text === 'string') return link.text;
    const formula = value as ExcelJS.CellFormulaValue;
    if ('formula' in formula) {
      const result = formula.result;
      // 캐시된 계산 결과가 없는 경우(우리 export가 저장한 수식은 결과를 캐시하지 않는다) —
      // 빈 값 대신 수식 원문을 그대로 돌려준다. 이 함수 단독으로는 여기서 끝이지만,
      // parseXlsxToGrid는 워크북을 다 읽은 뒤 2차 패스(applyComputedFormulaValues)에서 이렇게
      // 표시된 칸을 서버에서 직접 계산해 그 결과로 바꿔 넣는다 — 그래야 읽기전용 뷰어도
      // 동시편집기와 똑같은 값을 보여준다. 계산이 실패하면(#REF! 등 에러 값이 아니라 정말로
      // 값을 만들지 못하면) 이 수식 원문이 그대로 남아, 최소한 "여기 수식이 있다"는 건 알 수 있다.
      if (result === null || result === undefined) return `=${formula.formula}`;
      if (result instanceof Date) return formatDate(result);
      // 엑셀이 이미 캐시해 둔 논리값 결과("=A1>0" 등) — 위 리터럴 논리값과 같은 이유로 대문자로
      // 맞춘다. 캐시가 없어 서버가 직접 계산하는 경로는 formatFormulaResult가 이미 대문자를 쓴다.
      if (typeof result === 'boolean') return result ? 'TRUE' : 'FALSE';
      if (typeof result === 'number') return formatNumber(result, cell.numFmt);
      return String(result);
    }
    return '';
  }
  return String(value);
}

/** displayValue가 만든 표시용 텍스트(예: "2,287")를 계산·편집에 쓸 수 있는 "값 텍스트"(예: "2287")로
 * 되돌린다. Number("2,287")은 NaN이라, 이 텍스트를 그대로 수식 엔진에 먹이면 그 셀을 숫자가 아닌
 * 것으로 취급해 조용히 0으로 계산해버린다(실제 재현된 버그: 합계 열이 항상 0). 쉼표 구분 표시
 * 형식(number/currency)이 걸린 칸만, 숫자로 되돌릴 수 있을 때에만 벗겨낸다 — 되돌린 결과가 유한한
 * 숫자가 아니면(예: 수식 원문 "=SUM(B3,M3)", 쉼표가 들어간 일반 텍스트) 원문을 그대로 돌려준다.
 * "화면에 어떻게 보일지"는 style.nf/nd로 이미 따로 전달되므로 여기서는 순수 값만 복원하면 된다.
 *
 * 동시편집 시딩(xlsx/collab.ts의 seedGridFromXlsx/seedGridPageFromXlsx)과 읽기전용 뷰어의 수식
 * 계산 패스(이 파일의 applyComputedFormulaValues)가 **같은 규칙**을 써야 하므로 여기서 한 번만
 * 정의한다 — 한쪽만 고치면 "편집기에서는 맞는데 뷰어에서는 0" 같은 불일치가 다시 생긴다. */
export function unformattedCellValue(text: string, style: GridStyle | undefined): string {
  if (!style?.nf || (style.nf !== 'number' && style.nf !== 'currency')) return text;
  const stripped = text.replace(/[₩,\s]/g, '');
  return Number.isFinite(Number(stripped)) && stripped !== '' ? stripped : text;
}

/** displayValue가 "수식 원문 그대로" 폴백을 탄 칸인지 가려내기 위한 판정 — 값이 ExcelJS의 수식
 * 값 객체여야 한다. 표시 텍스트가 '='로 시작한다는 것만으로는 부족하다: 사용자가 진짜로 "=abc"
 * 라는 **문자열**을 입력해 둔 칸도 똑같이 '='로 시작하는데, 그런 칸까지 계산해 '#ERROR!'로
 * 덮어써버리면 멀쩡히 보이던 텍스트가 깨진다. */
function isFormulaCellValue(value: ExcelJS.CellValue): boolean {
  return value !== null && typeof value === 'object' && !(value instanceof Date) && 'formula' in value;
}

/** computeGridCellDisplay가 돌려준 값(number | string | boolean | 에러 이름 문자열 ...)을 그리드에
 * 저장할 표시 텍스트로 만든다. 숫자는 displayValue가 일반 숫자 칸에 쓰는 것과 **같은**
 * formatNumber(value, cell.numFmt)를 그대로 태운다 — 계산된 합계가 엑셀이 캐시해 둔 결과와 같은
 * 자릿수·쉼표 서식으로 보여야 하기 때문이다. '#REF!'/'#ERROR!' 같은 에러 이름은 문자열이라
 * 그대로 통과해, 엑셀이 깨진 수식에 보여주는 것과 같은 모양으로 표시된다. */
function formatFormulaResult(value: unknown, numFmt?: string): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return formatDate(value);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') {
    // 부동소수 잡음(0.1+0.2=0.30000000000000004) 제거 — formatComputedValueForStorage
    // (xlsx/formulas.ts)가 동시편집기 쪽에 쓰는 것과 같은 반올림 규칙. 절댓값이 너무 커서
    // value*1e10이 Infinity가 되는 경우에는 반올림을 건너뛴다(그 자릿수에서는 잡음도 없다).
    const rounded = Math.round(value * 1e10) / 1e10;
    return formatNumber(Number.isFinite(rounded) ? rounded : value, numFmt);
  }
  return String(value);
}

/** 수식 계산 패스가 계산해 채워야 하는 칸 하나(= displayValue가 수식 원문을 그대로 남긴 칸). */
interface PendingFormulaCell {
  readonly sheetIdx: number;
  readonly row: number;
  readonly col: number;
  /** 원본 엑셀 서식 코드 — 계산 결과를 일반 숫자 칸과 같은 규칙으로 서식화하기 위해 들고 다닌다. */
  readonly numFmt?: string;
}

/** 수식 계산 패스가 읽을 입력 그리드를 만든다 — 표시용 그리드와 내용은 같지만, 쉼표 서식이 걸린
 * 숫자 칸의 텍스트만 unformattedCellValue로 값 텍스트로 되돌려 둔다(그래야 SUM 등이 그 칸을
 * 숫자로 본다).
 *
 * 메모리: 수식이 **하나도 없는** 시트는 원본 시트 객체를 그대로 재사용하고, 수식이 있는 시트
 * 안에서도 서식이 벗겨질 칸이 하나도 없는 행은 원본 행 배열을 그대로 공유한다 — 실제로 새로
 * 할당되는 건 "쉼표 서식 숫자가 한 칸이라도 있는 행"뿐이다. 수십만 행 문서에서 표시용 그리드를
 * 통째로 복제하지 않기 위한 조치다. */
async function buildFormulaInputGrid(
  sheets: readonly GridSheet[],
  styles: readonly GridStyle[],
  sheetsWithFormulas: ReadonlySet<number>,
): Promise<GridDocument> {
  const inputSheets: GridSheet[] = [];
  for (let sheetIdx = 0; sheetIdx < sheets.length; sheetIdx += 1) {
    const sheet = sheets[sheetIdx];
    if (!sheetsWithFormulas.has(sheetIdx)) {
      inputSheets.push(sheet);
      continue;
    }
    const inputRows: Array<ReadonlyArray<GridCell | 0>> = [];
    for (let r = 0; r < sheet.rows.length; r += 1) {
      const row = sheet.rows[r];
      let unformatted: Array<GridCell | 0> | null = null;
      for (let c = 0; c < row.length; c += 1) {
        const cell = row[c];
        if (cell === 0) continue;
        const raw = unformattedCellValue(cell.v, cell.s === undefined ? undefined : styles[cell.s]);
        if (raw === cell.v) continue;
        // 이 행에서 처음으로 서식을 벗길 칸을 만난 순간에만 행을 복제한다(copy-on-write).
        if (!unformatted) unformatted = row.slice();
        // 계산 엔진은 v만 읽으므로 스타일 참조는 싣지 않는다(이 그리드는 순수 값 전용).
        unformatted[c] = { v: raw };
      }
      inputRows.push(unformatted ?? row);
      if ((r + 1) % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
    }
    inputSheets.push({ ...sheet, rows: inputRows });
  }
  return { sheets: inputSheets, styles };
}

/**
 * 워크북을 다 읽은 뒤 도는 2차 패스 — displayValue가 "캐시된 결과가 없어" 수식 원문을 그대로
 * 남겨둔 칸들을 서버에서 직접 계산해 그 결과로 바꿔 넣는다. 읽기전용 뷰어가 동시편집기
 * (collab/room.ts의 recomputeFormulasForPage)와 **같은** 값을 보여주게 하기 위한 것으로, 계산
 * 자체도 같은 함수(xlsx/formulas.ts의 computeGridCellDisplay)를 쓴다.
 *
 * **전체 시트를 다 만든 뒤에 도는 이유**: 수식은 자기 시트의 어느 칸이든(창 밖이든 아래쪽이든)
 * 참조할 수 있으므로, 행을 채우는 도중에 계산하면 아직 읽지 않은 아래쪽 참조가 빈 칸으로 보인다.
 *
 * **memo/tainted를 패스 전체가 공유하는 이유**: 체인형 수식(뒤 칸이 앞 칸을 참조)에서 memo가
 * 없으면 총 계산 횟수가 O(N²)로 폭증한다(computeGridCellDisplay 문서 주석 참고). 그래서 이
 * 함수 호출 하나 = 재계산 패스 하나로 보고 memo/tainted를 **여기서 한 번만** 만든다 — 셀마다·
 * 행마다 새로 만들면 안 된다.
 *
 * **sheetNameIndex도 마찬가지로 여기서 한 번만 만드는 이유**: computeGridCellDisplay가 자신에게
 * sheetNameIndex를 안 넘겨주면 호출마다(=여기 아래 루프의 pending 셀 하나마다) 새로 만든다 —
 * 재귀 안에서는 재사용되지만, 이 루프가 도는 pending.length번의 top-level 호출 각각은 서로 다시
 * 만든다는 뜻이다. memo/tainted와 똑같이 재계산 패스(=이 함수 호출) 하나에 시트 구성은 바뀌지
 * 않으므로 여기서 한 번만 만들어 모든 pending 셀에 걸쳐 재사용한다(buildSheetNameIndex 문서
 * 주석 참고).
 *
 * **계산이 다 끝난 뒤에야 표시용 그리드에 써 넣는 이유(중요)**: buildFormulaInputGrid는 메모리를
 * 아끼려고 "서식을 벗길 칸이 없는 행"의 배열을 표시용 그리드와 **그대로 공유**한다. 그래서 계산
 * 도중에 결과를 표시용 그리드에 바로 써 넣으면 그게 입력 그리드까지 같이 바꿔버린다 — 예를 들어
 * 이 프로젝트의 실제 파일처럼 `B49=SUM(B3:B48)`(수식)이고 `N49=SUM(B49:M49)`(그 수식들의 합)인
 * 경우, 먼저 계산된 B49의 표시 텍스트("12,153")가 입력 그리드에 덮어씌워져 N49가 그 칸을 더 이상
 * 수식으로도 숫자로도 보지 못하고 0을 내놓는다(실제로 재현됨). 결과를 전부 모은 뒤 한 번에
 * 반영하면, 모든 계산이 "파일을 막 읽었을 때의 원본 상태"만 보게 되어 이 종류의 별칭(aliasing)
 * 문제가 원천적으로 사라진다. memo 캐시 키가 좌표뿐인 것도 이 불변식(패스 동안 입력이 바뀌지
 * 않음) 위에서만 안전하다.
 */
async function applyComputedFormulaValues(
  sheets: readonly GridSheet[],
  sheetRows: ReadonlyArray<Array<Array<GridCell | 0>>>,
  styles: readonly GridStyle[],
  pending: readonly PendingFormulaCell[],
): Promise<void> {
  if (pending.length === 0) return;
  const inputGrid = await buildFormulaInputGrid(
    sheets,
    styles,
    new Set(pending.map((p) => p.sheetIdx)),
  );
  const memo = new Map<string, unknown>();
  const tainted = new Set<string>();
  const sheetNameIndex = buildSheetNameIndex(inputGrid);
  const computedTexts: string[] = new Array(pending.length).fill('');
  for (let i = 0; i < pending.length; i += 1) {
    const { sheetIdx, row, col, numFmt } = pending[i];
    try {
      computedTexts[i] = formatFormulaResult(
        computeGridCellDisplay(inputGrid, sheetIdx, row, col, new Set(), memo, tainted, sheetNameIndex),
        numFmt,
      );
    } catch {
      // 계산이 예외로 끝나면(예: 지원하지 않는 구문으로 인한 깊은 재귀) 빈 문자열로 남겨 아래에서
      // 수식 원문을 그대로 두게 한다 — 업로드 자체가 실패하는 것보다 "여기 수식이 있다"고
      // 보여주는 편이 낫다(이 변경 이전의 동작으로 폴백).
      computedTexts[i] = '';
    }
    if ((i + 1) % YIELD_EVERY_N_FORMULA_CELLS === 0) await yieldToEventLoop();
  }

  for (let i = 0; i < pending.length; i += 1) {
    const text = computedTexts[i];
    // 빈 결과는 반영하지 않는다: 칸을 비워 버리면 "여기 수식이 있다"는 정보까지 사라지고,
    // 값 없는 꼬리 행 트리밍이 이미 끝난 뒤라 행 수 계산과도 어긋난다.
    if (text === '') continue;
    const { sheetIdx, row, col } = pending[i];
    const cells = sheetRows[sheetIdx]?.[row];
    const cell = cells?.[col];
    // 꼬리 행 트리밍으로 사라진 칸일 수 있으므로 존재 여부를 먼저 본다.
    if (!cells || cell === 0 || cell === undefined || cell.v === text) continue;
    cells[col] = cell.s === undefined ? { v: text } : { v: text, s: cell.s };
    if ((i + 1) % YIELD_EVERY_N_FORMULA_CELLS === 0) await yieldToEventLoop();
  }
}

export function parseMergeRange(range: string): GridMerge | null {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) return null;
  const colNum = (letters: string) =>
    [...letters].reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
  const c1 = colNum(match[1]);
  const r1 = Number(match[2]);
  const c2 = colNum(match[3]);
  const r2 = Number(match[4]);
  return { r: r1, c: c1, rs: r2 - r1 + 1, cs: c2 - c1 + 1 };
}

export class StylePalette {
  private readonly styles: GridStyle[] = [];
  private readonly index = new Map<string, number>();

  register(style: GridStyle): number | undefined {
    if (Object.keys(style).length === 0) return undefined;
    const key = JSON.stringify(style);
    const existing = this.index.get(key);
    if (existing !== undefined) return existing;
    this.styles.push(style);
    this.index.set(key, this.styles.length - 1);
    return this.styles.length - 1;
  }

  toArray(): GridStyle[] {
    return this.styles;
  }
}

const DEFAULT_FONT_SIZE_PT = 11;
const PT_TO_PX = 96 / 72;
const MAX_FONT_PX = 26;

/** 엑셀 테두리 스타일 이름 → 단순화된 굵기(1=가는 선, 2=굵은 선). 스타일이 없으면 undefined. */
function borderWeight(border?: Partial<ExcelJS.Border>): BorderWeight | undefined {
  if (!border?.style) return undefined;
  return border.style === 'medium' || border.style === 'thick' || border.style === 'double' ? 2 : 1;
}

/** 원본 엑셀 서식 코드(예: "#,##0.00", "0.0%", "\"₩\"#,##0")를 그대로 저장하지 않고, 툴바가 다루는
 * 몇 가지 자주 쓰는 범주(쉼표/통화/백분율)+소수 자릿수로 단순화해서 읽는다 — 서식 코드 전체를
 * 파싱하는 건 범위 밖이고, 어차피 우리 툴바도 이 세 가지 버튼만 제공한다.
 *
 * **날짜 서식은 의도적으로 분류하지 않는다**(리뷰에서 지적된 gap — 날짜 칸이 기본 왼쪽 정렬/스필
 * 대상으로 잘못 취급되는 문제 — 는 여기서 NumberFormatKind에 'date'를 추가하는 대신 isSpillEligible
 * 쪽에서 표시 텍스트 모양(YYYY-MM-DD)만으로 처리했다). 이유: NumberFormatKind/style.nf는 이 파일
 * 밖에서도 실제 값 서식(그리드 편집기의 쉼표/통화/백분율 표시 형식 드롭다운 cell-numfmt-select —
 * public/editor-grid.html·src/client/grid-editor.ts, xlsx로 재내보낼 때 서식 코드를 고르는
 * src/xlsx/collab.ts의 numFmtStringFor)를 좌우하는 공유 타입이다. 그 소비처들은 전부 이 세 범주
 * only 기준으로 쓰여 있어(드롭다운엔 애초에 "날짜" 옵션이 없고, numFmtStringFor의 switch도 세
 * 범주만 안다), 'date'를 새 멤버로 넣으면 날짜 칸의 서식 드롭다운이 어떤 옵션도 선택되지 않은 채
 * 빈 채로 뜨는 등 이 정렬/스필 버그와 무관한 곳에서 새로운(더 눈에 띄는) 부작용이 생긴다. 반면 이
 * 앱이 날짜를 실제로 그리는 경로는 formatDate(이 파일) 단 하나뿐이고, 원본 서식 코드가 무엇이든
 * 항상 "YYYY-MM-DD"나 "YYYY-MM-DD HH:MM" 모양으로만 렌더링하므로, 그 모양을 표시 텍스트에서 직접
 * 알아보는 쪽이 공유 타입을 건드리지 않고도 훨씬 좁고 안전하게 같은 문제를 고친다. */
function classifyNumFmt(numFmt?: string): { nf?: NumberFormatKind; nd?: number } {
  if (!numFmt || numFmt === 'General' || numFmt === '@') return {};
  const decimals = numFmt.match(/\.(0+)/)?.[1].length ?? 0;
  if (numFmt.includes('%')) return { nf: 'percent', nd: decimals };
  if (/[$₩¥€]|"[^"]+"/.test(numFmt)) return { nf: 'currency', nd: decimals };
  if (numFmt.includes('#,##0') || /^0(\.0+)?$/.test(numFmt.trim())) return { nf: 'number', nd: decimals };
  return {};
}

export function extractStyle(
  cell: ExcelJS.Cell,
  palette: StylePalette,
  themeColors: readonly string[],
  sheetProtected: boolean,
): number | undefined {
  const style: {
    b?: 1;
    i?: 1;
    u?: 1;
    st?: 1;
    fs?: number;
    ff?: string;
    col?: string;
    bg?: string;
    a?: 'l' | 'c' | 'r';
    va?: 't' | 'm' | 'b';
    w?: 1;
    nf?: NumberFormatKind;
    nd?: number;
    lk?: 1;
    nu?: 1;
    bt?: BorderWeight;
    br?: BorderWeight;
    bb?: BorderWeight;
    bl?: BorderWeight;
  } = {};
  const font = cell.font;
  if (font?.bold) style.b = 1;
  if (font?.italic) style.i = 1;
  if (font?.underline) style.u = 1;
  if (font?.strike) style.st = 1;
  if (font?.size && font.size !== DEFAULT_FONT_SIZE_PT) {
    style.fs = Math.min(MAX_FONT_PX, Math.round(font.size * PT_TO_PX));
  }
  if (font?.name && font.name !== 'Calibri') style.ff = font.name;
  const fontColor = resolveColor(font?.color, themeColors);
  if (fontColor && fontColor.toUpperCase() !== '#000000') style.col = fontColor;
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const bg = resolveColor(fill.fgColor, themeColors);
    // 흰색이어도 캡처한다 — 잠긴 칸은 기본적으로 회색으로 렌더링되는데, 셀에 명시적으로 흰 배경을
    // 지정해두면(이 커넥션-현황 시트의 1행처럼) 그 지정이 잠금 회색보다 우선해야 한다.
    if (bg) style.bg = bg;
  }
  const horizontal = cell.alignment?.horizontal;
  if (horizontal === 'center') style.a = 'c';
  else if (horizontal === 'right') style.a = 'r';
  // 명시적 왼쪽 정렬도 캡처한다 — 안 그러면 이미 소스 xlsx에서 명시적으로 왼쪽 정렬된 칸(예:
  // 숫자 서식이면서 왼쪽 정렬인 칸)이 style.a 미지정으로 저장돼, isSpillEligible(뷰어/편집기
  // 둘 다) 쪽에서 "명시적 style.a==='l'은 항상 넘친다" 분기를 절대 못 타고 style.nf 기반 기본
  // 오른쪽 정렬로 잘못 분류되어 스필이 되레 억제된다(엑셀 실제 동작과 반대인, 옆 빈 칸으로 안
  // 넘치는 방향의 사소한 표시 차이). 'l'이 이제 실제 파일에서도 나올 수 있게 됐지만, 이 값을
  // 읽는 다른 소비처(ALIGN_TO_CSS/ALIGN_TO_EXCEL, 뷰어 injectGridStyles)는 이미 'l'을 다루도록
  // 돼 있어 부작용이 없다(확인 완료).
  else if (horizontal === 'left') style.a = 'l';
  const vertical = cell.alignment?.vertical;
  if (vertical === 'top') style.va = 't';
  else if (vertical === 'middle') style.va = 'm';
  else if (vertical === 'bottom') style.va = 'b';
  if (cell.alignment?.wrapText) style.w = 1;
  const { nf, nd } = classifyNumFmt(cell.numFmt);
  if (nf) {
    style.nf = nf;
    if (nd) style.nd = nd;
  }
  // 시트 보호가 걸려 있을 때만 잠금 여부가 실제로 의미를 가진다(엑셀도 그렇게 동작함) — 그렇지
  // 않은 대부분의 일반 문서는 셀마다 protection.locked가 true(엑셀 기본값)라도 편집 가능해야 한다.
  if (sheetProtected && cell.protection?.locked !== false) style.lk = 1;
  // 엑셀의 실제 데이터 유효성 검사(정수/소수 제한)를 그대로 읽는다 — webviewer가 새로 지어낸
  // 개념이 아니라 어떤 xlsx든 이 규칙이 있으면 똑같이 동작한다.
  const validationType = cell.dataValidation?.type;
  if (validationType === 'whole' || validationType === 'decimal') style.nu = 1;
  const border = cell.border;
  const bt = borderWeight(border?.top);
  if (bt) style.bt = bt;
  const br = borderWeight(border?.right);
  if (br) style.br = br;
  const bb = borderWeight(border?.bottom);
  if (bb) style.bb = bb;
  const bl = borderWeight(border?.left);
  if (bl) style.bl = bl;
  return palette.register(style);
}

export async function parseXlsxToGrid(xlsxBytes: Buffer): Promise<GridDocument> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsxBytes as unknown as ArrayBuffer);
  const palette = new StylePalette();
  // exceljs 타입 선언(string[])과 달리 런타임은 { theme1: xml문자열 } 형태다
  const themes = (workbook.model as unknown as { themes?: Record<string, string> }).themes;
  const themeColors = parseThemeColors(themes?.theme1);
  const sheets: GridSheet[] = [];
  // sheets와 1:1로 대응하는 "쓰기 가능한" 행 배열 핸들 — GridSheet.rows는 readonly로 선언돼 있어
  // 수식 계산 패스가 결과를 써 넣을 수 없으므로, 같은 배열 객체를 여기에도 담아 둔다.
  const sheetRows: Array<Array<Array<GridCell | 0>>> = [];
  const pendingFormulas: PendingFormulaCell[] = [];

  // workbook.eachSheet/sheet.eachRow는 콜백형 동기 API라 중간에 await로 이벤트 루프에 양보할 수
  // 없다 — 배열(workbook.worksheets)을 직접 도는 for 루프로 바꿔, 대형 시트를 열 때 그 문서를 연
  // 사용자뿐 아니라 그 순간 서버에 접속한 다른 모든 사용자·다른 문서 요청까지 함께 멈추지 않도록
  // 주기적으로 양보한다.
  for (const sheet of workbook.worksheets) {
    if (sheet.state && sheet.state !== 'visible') continue;
    // sheet.actualRowCount는 ExcelJS 자체 문서상 "실제 값이 있는 행의 '개수'"라, 성긴 시트(예:
    // 1행과 350행에만 값이 있고 그 사이가 전부 비어 있는 시트)에서는 실제 최대 행 번호(350)보다
    // 훨씬 작은 값(예: 2)을 돌려준다 — worksheet.js의 실제 구현(eachRow가 row.hasValues인 행만
    // 세는 것)으로 직접 확인함. 이 값을 rowCount로 그대로 쓰면 뒤쪽의 진짜 데이터가 통째로 잘려
    // 나간다(동시편집기 lazyGridStore.ts의 ingestXlsxLazily에서 이미 한 번 겪고 고친 것과 같은
    // 부류의 버그 — 그쪽은 스트리밍 리더가 실제로 마주친 <row> 엘리먼트의 row.number를 그대로
    // 신뢰하는 방식으로 고쳤다). 여기서는 대신 sheet.rowCount(내부 _lastRowNumber — worksheet.js의
    // model setter가 `this._rows[row.number - 1] = row`로 채워서, 값이 있든 서식만 있든 실제로
    // <row> XML 엘리먼트가 존재했던 가장 큰 행 번호를 그대로 반영한다)를 쓴다. 값 없이 서식만
    // 남은 "유령 행"이 맨 끝에 더 포함되는 과다 계산 걱정은 없다 — 아래 lastWithValue 트리밍이
    // 어차피 그런 꼬리 행을 실제 데이터 기준으로 다시 잘라낸다.
    const rowCount = Math.min(sheet.rowCount || 0, MAX_ROWS);
    if (rowCount === 0) continue;
    // exceljs 타입 선언에는 없지만 런타임에는 시트 보호 시 채워진다(themes와 동일한 이유로 캐스팅)
    const sheetProtected = Boolean((sheet as unknown as { sheetProtection?: unknown }).sheetProtection);

    // actualColumnCount는 맨 오른쪽 열이 수식 셀만 있을 때 그 열을 세지 않는 경우가 있다(ExcelJS
    // 자체 버그로 보임 — 값이 아니라 수식만 있는 열은 "dimensions" 추적에서 빠지는 듯하다). 이 값만
    // 믿으면 예: 맨 오른쪽 합계(=SUM) 열이 통째로 누락되고, 그 결과 동시편집 저장 시
    // exportXlsxFromGrid의 "구조 밖 열 잘라내기" 로직이 그 열을 실제로 지워버린다. 그래서 아래
    // 본문 루프에서 실제로 쓰인 최대 열 번호(maxColUsed)도 함께 구해 더 큰 쪽을 쓴다 — 예전에는
    // 이걸 위해 전체 행을 한 번 더 훑는 별도 사전 스캔이 있었지만, 각 행의 칸 배열을 고정 폭이
    // 아니라 필요할 때마다 늘려가며 채우면(streaming 경로 ingestXlsxLazily와 같은 기법) 본문
    // 루프 한 번으로 maxColUsed까지 같이 구할 수 있어 ExcelJS 행/셀 순회를 두 번 하지 않는다.
    let maxColUsed = 0;
    const rows: Array<Array<GridCell | 0>> = [];
    // 이 시트가 sheets 배열에서 갖게 될 인덱스 — 아래 sheets.push가 이 반복의 마지막이고, 여기까지
    // 왔으면 더 이상 continue로 건너뛰지 않으므로 지금 값이 곧 최종 인덱스다(숨김/빈 시트가
    // 걸러지므로 워크북 상의 시트 번호와는 다를 수 있다).
    const sheetIdx = sheets.length;
    for (let r = 1; r <= rowCount; r += 1) {
      const row = sheet.getRow(r);
      const cells: Array<GridCell | 0> = [];
      // includeEmpty:true여야 한다 — 값은 없지만 스타일(잠금·숫자 전용 유효성 검사·테두리 등)만
      // 있는 칸도 빠짐없이 읽는다. false로 두면 "값이 없는 칸"은 아예 순회 대상에서 빠져, 예를 들어
      // 아직 아무도 입력하지 않은 숫자 전용 칸이 잠금·유효성 검사 정보 없이 렌더링되어 아무 텍스트나
      // 입력·붙여넣기해도 막히지 않는 문제가 있었다(실제로 재현된 버그).
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber > MAX_COLS) return;
        if (colNumber > maxColUsed) maxColUsed = colNumber;
        while (cells.length < colNumber) cells.push(0);
        const text = displayValue(cell);
        const styleRef = extractStyle(cell, palette, themeColors, sheetProtected);
        if (text === '' && styleRef === undefined) return;
        cells[colNumber - 1] = styleRef === undefined ? { v: text } : { v: text, s: styleRef };
        // displayValue가 캐시된 결과가 없어 수식 원문을 그대로 남긴 칸 — 워크북을 다 읽은 뒤
        // applyComputedFormulaValues가 직접 계산해 채운다(displayValue 문서 주석 참고).
        if (text.startsWith('=') && isFormulaCellValue(cell.value)) {
          pendingFormulas.push({ sheetIdx, row: r - 1, col: colNumber - 1, numFmt: cell.numFmt });
        }
      });
      rows.push(cells);
      if (r % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
    }
    const colCount = Math.min(Math.max(sheet.actualColumnCount || 0, maxColUsed), MAX_COLS);
    if (colCount === 0) continue;
    // 위 루프는 각 행을 그 행에서 실제로 마주친 마지막 칸까지만 채우므로(동적 성장), 시트 전체의
    // colCount보다 짧게 끝난 행이 있을 수 있다 — 고정 폭 배열이었던 예전 방식과 동일한 모양으로
    // 맞추기 위해 0으로 패딩한다.
    for (const cells of rows) {
      while (cells.length < colCount) cells.push(0);
    }

    const colWidths: number[] = [];
    for (let c = 1; c <= colCount; c += 1) {
      const width = sheet.getColumn(c).width;
      colWidths.push(width ? Math.round(width * CHAR_TO_PX) : DEFAULT_COL_WIDTH_PX);
    }

    // 서식만 남은 유령 행(값 없는 꼬리 행)을 잘라내 행 수를 실제 데이터 기준으로 맞춘다
    let lastWithValue = rows.length;
    while (lastWithValue > 1) {
      const row = rows[lastWithValue - 1];
      const hasValue = row.some((cell) => cell !== 0 && cell.v !== '');
      if (hasValue) break;
      lastWithValue -= 1;
    }
    rows.length = lastWithValue;

    const merges = (
      (sheet.model as { merges?: readonly string[] }).merges ?? []
    )
      .map(parseMergeRange)
      .filter(
        (merge): merge is GridMerge => merge !== null && merge.r <= rows.length && merge.c <= colCount,
      );

    const frozenView = sheet.views?.find((v) => (v as { state?: string }).state === 'frozen');
    const freezeRows = frozenView ? ((frozenView as { ySplit?: number }).ySplit ?? 0) : 0;
    const freezeCols = frozenView ? ((frozenView as { xSplit?: number }).xSplit ?? 0) : 0;

    sheets.push({ name: sheet.name, rowCount: rows.length, colCount, colWidths, merges, freezeRows, freezeCols, rows });
    sheetRows.push(rows);
  }

  const styles = palette.toArray();
  // 2차 패스 — 시트를 전부 만든 뒤에야 수식이 참조하는 칸이 모두 채워져 있다.
  await applyComputedFormulaValues(sheets, sheetRows, styles, pendingFormulas);
  return { sheets, styles };
}

/** 간단한 RFC4180 CSV 파서 (따옴표·이스케이프 지원) */
export function parseCsvToGrid(text: string): GridDocument {
  const rows: string[][] = [];
  let field = '';
  let record: string[] = [];
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      record.push(field);
      field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i += 1;
      record.push(field);
      field = '';
      rows.push(record);
      record = [];
    } else {
      field += ch;
    }
  }
  if (field !== '' || record.length > 0) {
    record.push(field);
    rows.push(record);
  }

  const bounded = rows.slice(0, MAX_ROWS);
  const colCount = Math.min(Math.max(...bounded.map((r) => r.length), 0), MAX_COLS);
  const gridRows = bounded.map((r) => {
    const cells: Array<GridCell | 0> = new Array(colCount).fill(0);
    for (let c = 0; c < Math.min(r.length, colCount); c += 1) {
      if (r[c] !== '') cells[c] = { v: r[c] };
    }
    return cells;
  });

  return {
    sheets: [
      {
        name: 'Sheet1',
        rowCount: gridRows.length,
        colCount,
        colWidths: new Array(colCount).fill(120),
        merges: [],
        freezeRows: 0,
        freezeCols: 0,
        rows: gridRows,
      },
    ],
    styles: [],
  };
}
