import ExcelJS from 'exceljs';
import type { DocumentStore } from './storage.js';

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

/** JSONL 헤더(1행): 셀 데이터를 제외한 시트 메타 정보 */
export interface GridHeader {
  readonly styles: readonly GridStyle[];
  readonly sheets: ReadonlyArray<Omit<GridSheet, 'rows'>>;
}

/**
 * 그리드를 JSONL로 직렬화한다: 1행 = 헤더, 이후 시트 순서대로 행당 1줄.
 * 수십만 행 문서를 요청 범위만 파싱해 서빙할 수 있도록 하기 위함이다.
 */
export function serializeGrid(grid: GridDocument): Buffer {
  const header: GridHeader = {
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
      // 빈 값 대신 수식 원문을 그대로 돌려준다. 동시편집기는 이 값을 다시 계산해 보여주고,
      // 읽기전용 뷰어는 최소한 "여기 수식이 있다"는 걸 알 수 있다(빈 칸보다 훨씬 낫다).
      if (result === null || result === undefined) return `=${formula.formula}`;
      if (result instanceof Date) return formatDate(result);
      if (typeof result === 'number') return formatNumber(result, cell.numFmt);
      return String(result);
    }
    return '';
  }
  return String(value);
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
 * 파싱하는 건 범위 밖이고, 어차피 우리 툴바도 이 세 가지 버튼만 제공한다. */
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
    // actualColumnCount는 맨 오른쪽 열이 수식 셀만 있을 때 그 열을 세지 않는 경우가 있다(ExcelJS
    // 자체 버그로 보임 — 값이 아니라 수식만 있는 열은 "dimensions" 추적에서 빠지는 듯하다). 이 값만
    // 믿으면 예: 맨 오른쪽 합계(=SUM) 열이 통째로 누락되고, 그 결과 동시편집 저장 시
    // exportXlsxFromGrid의 "구조 밖 열 잘라내기" 로직이 그 열을 실제로 지워버린다. 그래서 전체
    // 행을 훑어 실제로 쓰인 최대 열 번호도 함께 구해 더 큰 쪽을 쓴다. (이 스캔 범위도 위와 같은
    // 이유로 actualRowCount가 아니라 이미 고쳐진 rowCount를 그대로 재사용해야, 성긴 시트에서
    // 뒤쪽 열 스캔 자체가 일찍 끊기는 일이 없다.)
    let maxColUsed = 0;
    const scanRowCount = rowCount;
    for (let r = 1; r <= scanRowCount; r += 1) {
      sheet.getRow(r).eachCell({ includeEmpty: false }, (_cell, colNumber) => {
        if (colNumber > maxColUsed) maxColUsed = colNumber;
      });
      if (r % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
    }
    const colCount = Math.min(Math.max(sheet.actualColumnCount || 0, maxColUsed), MAX_COLS);
    if (rowCount === 0 || colCount === 0) continue;
    // exceljs 타입 선언에는 없지만 런타임에는 시트 보호 시 채워진다(themes와 동일한 이유로 캐스팅)
    const sheetProtected = Boolean((sheet as unknown as { sheetProtection?: unknown }).sheetProtection);

    const colWidths: number[] = [];
    for (let c = 1; c <= colCount; c += 1) {
      const width = sheet.getColumn(c).width;
      colWidths.push(width ? Math.round(width * CHAR_TO_PX) : DEFAULT_COL_WIDTH_PX);
    }

    const rows: Array<Array<GridCell | 0>> = [];
    for (let r = 1; r <= rowCount; r += 1) {
      const row = sheet.getRow(r);
      const cells: Array<GridCell | 0> = new Array(colCount).fill(0);
      // includeEmpty:true여야 한다 — 값은 없지만 스타일(잠금·숫자 전용 유효성 검사·테두리 등)만
      // 있는 칸도 빠짐없이 읽는다. false로 두면 "값이 없는 칸"은 아예 순회 대상에서 빠져, 예를 들어
      // 아직 아무도 입력하지 않은 숫자 전용 칸이 잠금·유효성 검사 정보 없이 렌더링되어 아무 텍스트나
      // 입력·붙여넣기해도 막히지 않는 문제가 있었다(실제로 재현된 버그).
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber > colCount) return;
        const text = displayValue(cell);
        const styleRef = extractStyle(cell, palette, themeColors, sheetProtected);
        if (text === '' && styleRef === undefined) return;
        cells[colNumber - 1] = styleRef === undefined ? { v: text } : { v: text, s: styleRef };
      });
      rows.push(cells);
      if (r % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
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
  }

  return { sheets, styles: palette.toArray() };
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
