import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { promisify } from 'node:util';
import { PDFDocument } from 'pdf-lib';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { trimPdfMargins } from './pdfTrim.js';
import type { Converter } from './types.js';

const execFileAsync = promisify(execFile);

const CONVERT_TIMEOUT_MS = 180_000;

const CANDIDATE_PATHS = [
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  '/usr/bin/soffice',
  '/usr/local/bin/soffice',
];

function detectSoffice(): string | null {
  const fromEnv = process.env.DOCVIEW_SOFFICE;
  if (fromEnv) {
    return existsSync(fromEnv) ? fromEnv : null;
  }
  return CANDIDATE_PATHS.find((candidate) => existsSync(candidate)) ?? null;
}

/** null이면 LibreOffice 미설치 → Office/DXF 변환기가 등록되지 않는다. */
export const sofficePath = detectSoffice();

/** soffice.exe 경로에서 LibreOffice 설치 루트를 역산한다(…/program/soffice.exe -> 그 두 단계 위). */
function libreOfficeInstallRoot(): string | null {
  if (!sofficePath) return null;
  return path.dirname(path.dirname(sofficePath));
}

/** H2Orestart(HWP 5.0/HWPX 임포트 필터) 확장이 설치돼 있는지 — 딱 정해진 API가 없어 설치된
 * 확장 캐시 폴더(share/uno_packages/cache/uno_packages/)를 얕게(2단계) 뒤져 이름에 "h2orestart"가
 * 들어간 항목이 있는지로 판단한다(실제로 이 저장소 개발 중 설치해 확인한 경로 구조와 일치).
 * 관리자 페이지의 "참고용" 표시라 확정적 보장은 아니다 — LibreOffice가 없거나 폴더 구조가 다른
 * 배포판/버전이면 그냥 null(확인 불가)을 돌려주고, 상태 텍스트에서 그렇게 표시한다. */
export async function detectH2Orestart(): Promise<boolean | null> {
  const root = libreOfficeInstallRoot();
  if (!root) return null;
  const cacheDir = path.join(root, 'share', 'uno_packages', 'cache', 'uno_packages');
  try {
    const entries = await fs.readdir(cacheDir);
    for (const entry of entries) {
      if (entry.toLowerCase().includes('h2orestart')) return true;
      const sub = await fs.readdir(path.join(cacheDir, entry)).catch(() => [] as string[]);
      if (sub.some((name) => name.toLowerCase().includes('h2orestart'))) return true;
    }
    return false;
  } catch {
    return null; // 캐시 폴더 자체가 없는 등 — "확인 불가"로 취급(설치 안 됐다고 단정하지 않는다)
  }
}

/** 머리글 등에 노출되는 표시용 파일명 — 경로 구분자와 예약 문자를 제거한다. */
function toSafeDisplayName(sourceName: string | undefined, ext: string): string {
  const base = (sourceName ?? '').replace(/\.[^.]+$/, '');
  const forbidden = new Set(['\\', '/', ':', '*', '?', '"', '<', '>', '|']);
  const safe = [...base]
    .filter((ch) => !forbidden.has(ch) && ch.charCodeAt(0) >= 0x20)
    .join('')
    .trim()
    .slice(0, 80);
  return `${safe.length > 0 ? safe : 'document'}.${ext}`;
}

function transformWorksheets(
  xlsxBytes: Uint8Array,
  transform: (xml: string) => string,
): Uint8Array {
  const entries = unzipSync(xlsxBytes);
  const updated: Record<string, Uint8Array> = {};
  for (const [name, bytes] of Object.entries(entries)) {
    updated[name] = /^xl\/worksheets\/[^/]+\.xml$/.test(name)
      ? strToU8(transform(strFromU8(bytes)))
      : bytes;
  }
  return zipSync(updated);
}

/**
 * 고정 행 높이(ht/customHeight)를 제거한다.
 * Excel과 LibreOffice의 글꼴 폭 계산 차이로, Excel에서 꼭 맞게 저장된 행 높이가
 * LibreOffice에서는 부족해 줄바꿈 텍스트가 잘린다. 높이를 제거하면 Calc가
 * 내용에 맞게 행 높이를 자동 계산하므로 텍스트가 잘리지 않는다.
 */
function stripRowHeightsXml(xml: string): string {
  return xml.replace(/<row\b[^>]*?>/g, (tag) =>
    tag.replace(/ ht="[^"]*"/g, '').replace(/ customHeight="(?:1|true)"/g, ''),
  );
}

/**
 * 시트를 "페이지 폭에 맞춰 축소 + 세로로 일반 페이지 분할"로 인쇄하도록 강제한다.
 * 세로로 매우 긴 리스트 시트에서 SinglePageSheets가 만드는 초장신 페이지
 * (렌더링 픽셀 상한 때문에 뿌옇게 보이는 문제)의 대안 경로.
 */
function forceFitToWidthXml(xml: string): string {
  let out = xml;
  const fitToPage = '<pageSetUpPr fitToPage="1"/>';
  if (/<sheetPr\b[^>]*\/>/.test(out)) {
    out = out.replace(/<sheetPr\b([^>]*)\/>/, `<sheetPr$1>${fitToPage}</sheetPr>`);
  } else if (/<sheetPr\b/.test(out)) {
    out = /<pageSetUpPr\b/.test(out)
      ? out.replace(/<pageSetUpPr\b[^>]*\/>/, fitToPage)
      : out.replace(/<sheetPr\b([^>]*)>/, `<sheetPr$1>${fitToPage}`);
  } else {
    out = out.replace(/(<worksheet\b[^>]*>)/, `$1<sheetPr>${fitToPage}</sheetPr>`);
  }

  const setupAttrs = 'fitToWidth="1" fitToHeight="0" orientation="landscape"';
  if (/<pageSetup\b/.test(out)) {
    out = out.replace(/<pageSetup\b([^>]*?)\/>/, (_match, attrs: string) => {
      const cleaned = attrs
        .replace(/\sfitToWidth="[^"]*"/g, '')
        .replace(/\sfitToHeight="[^"]*"/g, '')
        .replace(/\sorientation="[^"]*"/g, '');
      return `<pageSetup${cleaned} ${setupAttrs}/>`;
    });
  } else {
    // 스키마상 pageSetup 뒤에 올 수 있는 요소들 앞(없으면 문서 끝)에 삽입한다
    out = out.replace(
      /<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<tableParts\b|<\/worksheet>/,
      (anchor) => `<pageSetup ${setupAttrs}/>${anchor}`,
    );
  }
  return out;
}

export function stripCustomRowHeights(xlsxBytes: Uint8Array): Uint8Array {
  return transformWorksheets(xlsxBytes, stripRowHeightsXml);
}

/** 폭 맞춤 페이지 분할만 적용 (수천 행 리스트용 — 행높이 재계산 비용을 피한다) */
export function applyFitToWidth(xlsxBytes: Uint8Array): Uint8Array {
  return transformWorksheets(xlsxBytes, forceFitToWidthXml);
}

/** 폭 맞춤 + 행높이 자동 계산 (일반 크기 시트의 폴백용) */
export function applyFitToWidthAndStrip(xlsxBytes: Uint8Array): Uint8Array {
  return transformWorksheets(xlsxBytes, (xml) => forceFitToWidthXml(stripRowHeightsXml(xml)));
}

/** 시트별 최대 행 번호를 dimension 참조로 가볍게 추정한다 (전체 파싱 없이). */
export function maxSheetRows(xlsxBytes: Uint8Array): number {
  const entries = unzipSync(xlsxBytes);
  let max = 0;
  for (const [name, bytes] of Object.entries(entries)) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(name)) continue;
    const xml = strFromU8(bytes);
    const dimension = xml.match(/<dimension ref="[A-Z]+\d+:[A-Z]+(\d+)"/);
    if (dimension) {
      max = Math.max(max, Number(dimension[1]));
      continue;
    }
    const rowRefs = xml.match(/<row r="(\d+)"[^>]*>(?![\s\S]*<row )/);
    if (rowRefs) max = Math.max(max, Number(rowRefs[1]));
  }
  return max;
}

/** a:fontScheme/a:minorFont(또는 a:majorFont) 안에서 라틴 기본 서체와 한글(Hang) 서체를 뽑는다. */
function themeSlotTypefaces(
  themeXml: string,
  slot: 'minorFont' | 'majorFont',
): { latin: string | null; hang: string | null } {
  const slotMatch = new RegExp(`<a:${slot}>([\\s\\S]*?)</a:${slot}>`).exec(themeXml);
  if (!slotMatch) return { latin: null, hang: null };
  const body = slotMatch[1];
  const latin = /<a:latin typeface="([^"]*)"/.exec(body)?.[1];
  const hang = /<a:font script="Hang" typeface="([^"]*)"/.exec(body)?.[1];
  return { latin: latin || null, hang: hang || null };
}

/**
 * w:rFonts 속성 문자열 하나에서 테마 참조 하나(예: w:eastAsiaTheme="minorEastAsia")를
 * word/theme/theme1.xml이 실제로 가리키는 리터럴 서체 이름으로 바꿔치기한다. 이미 같은
 * 자리의 리터럴 서체(w:eastAsia 등)가 있으면 그 값을 존중하고 테마 참조만 지운다 — 리터럴과
 * 테마 참조가 동시에 있을 때 LibreOffice가 어느 쪽을 우선하는지 불확실하므로 모호함을 없앤다.
 */
function replaceThemeFontAttr(
  attrsStr: string,
  themeAttr: string,
  literalAttr: string,
  script: 'latin' | 'hang',
  themeXml: string,
): string {
  const themeMatch = new RegExp(`\\s${themeAttr}="([^"]*)"`).exec(attrsStr);
  if (!themeMatch) return attrsStr;
  const withoutTheme = attrsStr.replace(themeMatch[0], '');
  if (new RegExp(`\\s${literalAttr}="`).test(attrsStr)) return withoutTheme;
  const slot = /major/i.test(themeMatch[1]) ? 'majorFont' : 'minorFont';
  const typeface = themeSlotTypefaces(themeXml, slot)[script];
  return typeface ? `${withoutTheme} ${literalAttr}="${typeface}"` : withoutTheme;
}

function literalizeThemeFontsXml(xml: string, themeXml: string): string {
  return xml.replace(/<w:rFonts\b([^>]*)\/>/g, (_whole, attrsStr: string) => {
    let out = attrsStr;
    out = replaceThemeFontAttr(out, 'w:asciiTheme', 'w:ascii', 'latin', themeXml);
    out = replaceThemeFontAttr(out, 'w:eastAsiaTheme', 'w:eastAsia', 'hang', themeXml);
    out = replaceThemeFontAttr(out, 'w:hAnsiTheme', 'w:hAnsi', 'latin', themeXml);
    return `<w:rFonts${out}/>`;
  });
}

/**
 * docx 변환 전처리 — LibreOffice가 OOXML 테마 폰트 참조(w:eastAsiaTheme 등)를 안정적으로
 * 해석하지 못해, 실제 설치된 서체(예: 맑은 고딕)를 두고도 자체 기본 서체(굴림 등)로 대체해
 * 버리는 문제가 확인됐다 — 그 결과 줄바꿈·행 높이가 원본 Word 렌더링과 미묘하게 달라져,
 * 표가 딱 맞게 들어가야 할 마지막 페이지에 불필요한 빈 페이지가 추가로 생기기도 한다.
 * 변환 전에 테마 참조를 theme1.xml이 가리키는 실제 서체 이름으로 미리 박아 넣으면 이 문제를
 * 피할 수 있다. 테마 파일이 없으면 원본을 그대로 반환한다.
 */
export function literalizeDocxThemeFonts(docxBytes: Uint8Array): Uint8Array {
  const entries = unzipSync(docxBytes);
  const themeXmlBytes = entries['word/theme/theme1.xml'];
  if (!themeXmlBytes) return docxBytes;
  const themeXml = strFromU8(themeXmlBytes);
  const updated: Record<string, Uint8Array> = { ...entries };
  for (const name of ['word/styles.xml', 'word/document.xml']) {
    const bytes = entries[name];
    if (bytes) updated[name] = strToU8(literalizeThemeFontsXml(strFromU8(bytes), themeXml));
  }
  return zipSync(updated);
}

type SourcePreprocessor = (bytes: Uint8Array) => Uint8Array;

/**
 * 프로필 풀: 매 변환마다 새 프로필을 만들면 LibreOffice가 프로필 초기화에
 * 수 초를 쓰므로, 웜업된 프로필을 슬롯 단위로 재사용한다. 슬롯별로 직렬화되어
 * 같은 프로필을 두 프로세스가 동시에 쓰는 일은 없다.
 */
const PROFILE_POOL_SIZE = 2;
const profileRoot = path.join(os.tmpdir(), 'docview-lo-profiles');
const slotQueues: Array<Promise<unknown>> = Array.from({ length: PROFILE_POOL_SIZE }, () =>
  Promise.resolve(),
);
let nextSlot = 0;

async function withProfileSlot<T>(task: (profileDir: string) => Promise<T>): Promise<T> {
  const slot = nextSlot;
  nextSlot = (nextSlot + 1) % PROFILE_POOL_SIZE;
  const previous = slotQueues[slot];
  let release!: () => void;
  const current = new Promise<void>((resolve) => {
    release = resolve;
  });
  slotQueues[slot] = previous.then(() => current);
  await previous.catch(() => {});
  const profileDir = path.join(profileRoot, `slot-${slot}`);
  try {
    await fs.mkdir(profileDir, { recursive: true });
    return await task(profileDir);
  } catch (error) {
    // 변환 실패가 프로필 손상 때문일 수 있으므로 다음 실행에서 새로 만들도록 비운다.
    // (Windows EBUSY 대비 짧게 재시도 — 실패해도 무시하는 건 기존과 동일)
    await fs.rm(profileDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 }).catch(() => {});
    throw error;
  } finally {
    release();
  }
}

async function convertWithSoffice(
  sourcePath: string,
  targetPdfPath: string,
  extraArgs: readonly string[],
  sourceName?: string,
  preprocess?: SourcePreprocessor,
): Promise<void> {
  if (!sofficePath) {
    throw new Error('LibreOffice(soffice)를 찾을 수 없습니다.');
  }
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-lo-'));
  try {
    // LibreOffice가 문서 머리글에 UUID 대신 원래 파일명을 쓰도록 표시용 이름으로 복사해 변환한다.
    const ext = path.extname(sourcePath).slice(1).toLowerCase();
    const displayName = toSafeDisplayName(sourceName, ext);
    const stagedSource = path.join(workDir, displayName);
    if (preprocess) {
      const original = await fs.readFile(sourcePath);
      let staged: Uint8Array;
      try {
        staged = preprocess(original);
      } catch {
        // 전처리에 실패해도 변환 자체는 원본으로 계속한다.
        staged = original;
      }
      await fs.writeFile(stagedSource, staged);
    } else {
      await fs.copyFile(sourcePath, stagedSource);
    }

    await withProfileSlot(async (profileDir) => {
      await execFileAsync(
        sofficePath,
        [
          '--headless',
          '--norestore',
          `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
          ...extraArgs,
          '--outdir',
          workDir,
          stagedSource,
        ],
        { timeout: CONVERT_TIMEOUT_MS },
      );
    });
    // soffice는 변환 실패 시에도 종료 코드 0을 반환할 수 있어 산출물 존재로 판정한다.
    const produced = path.join(workDir, `${path.parse(displayName).name}.pdf`);
    try {
      await fs.access(produced);
    } catch {
      throw new Error(`LibreOffice가 PDF를 생성하지 못했습니다: ${displayName}`);
    }
    await fs.copyFile(produced, targetPdfPath);
  } finally {
    // Windows에서 백신 등이 방금 만든 임시 파일을 잠깐 스캔하는 동안 삭제가 EBUSY로 실패하는
    // 경우가 있어, 짧게 재시도한다.
    await fs.rm(workDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

/** DXF → SVG (도면 열람 전용 부가 산출물). convertWithSoffice와 달리 산출 확장자가 .pdf가
 * 아니라 이 함수만 따로 둔다 — 실패해도 PDF 변환은 이미 끝난 뒤라 문서 자체는 정상 열람된다
 * (호출자인 dwgConverter/documents.ts가 실패를 잡아 viewKind를 그냥 'pages'로 둔다). */
export async function convertDxfToSvg(dxfPath: string, targetSvgPath: string): Promise<void> {
  if (!sofficePath) {
    throw new Error('LibreOffice(soffice)를 찾을 수 없습니다.');
  }
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-dxfsvg-'));
  try {
    const stagedSource = path.join(workDir, 'drawing.dxf');
    await fs.copyFile(dxfPath, stagedSource);
    await withProfileSlot(async (profileDir) => {
      await execFileAsync(
        sofficePath,
        [
          '--headless',
          '--norestore',
          `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
          '--infilter=DXF',
          '--convert-to',
          'svg:draw_svg_Export',
          '--outdir',
          workDir,
          stagedSource,
        ],
        { timeout: CONVERT_TIMEOUT_MS },
      );
    });
    const produced = path.join(workDir, 'drawing.svg');
    const stat = await fs.stat(produced).catch(() => null);
    if (!stat || stat.size === 0) {
      throw new Error('LibreOffice가 SVG를 생성하지 못했습니다.');
    }
    await fs.copyFile(produced, targetSvgPath);
  } finally {
    await fs.rm(workDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

/**
 * 레거시 스프레드시트(xls/ods) → xlsx. 이 서비스의 그리드 뷰·동시편집은 전부 xlsx(zip+XML) 구조를
 * 전제로 짜여 있어(스프레드시트.ts/room.ts 어디에도 BIFF 바이너리(.xls) 파서가 없다) xls를 실제로
 * "편집"하려면 업로드 시점에 xlsx로 정규화하는 것 외에 다른 방법이 없다. documents.ts의 업로드
 * 핸들러가 findConverter를 부르기 전, ext/데이터를 아예 xlsx로 바꿔치기해 호출한다 — 그 뒤로는
 * 시스템 전체가 원래부터 xlsx가 올라온 것처럼 동작한다(그리드 뷰·동시편집 전부 자동으로 켜짐).
 * 변환 후로는 다운로드도 xlsx로 내려간다(원본 .xls 바이트는 보존하지 않는다) — 편집 자체가 xlsx
 * 모델을 통해서만 가능하므로 "편집됐을 수도 있는 문서"를 원본 바이너리 형식으로 되돌려 줄 방법이
 * 애초에 없다.
 */
export async function convertLegacySpreadsheetToXlsx(
  sourcePath: string,
  sourceExt: 'xls' | 'ods',
  targetXlsxPath: string,
): Promise<void> {
  if (!sofficePath) {
    throw new Error('LibreOffice(soffice)를 찾을 수 없습니다.');
  }
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-legacyxlsx-'));
  try {
    const stagedSource = path.join(workDir, `source.${sourceExt}`);
    await fs.copyFile(sourcePath, stagedSource);
    await withProfileSlot(async (profileDir) => {
      await execFileAsync(
        sofficePath,
        [
          '--headless',
          '--norestore',
          `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
          '--convert-to',
          'xlsx',
          '--outdir',
          workDir,
          stagedSource,
        ],
        { timeout: CONVERT_TIMEOUT_MS },
      );
    });
    const produced = path.join(workDir, 'source.xlsx');
    const stat = await fs.stat(produced).catch(() => null);
    if (!stat || stat.size === 0) {
      throw new Error('LibreOffice가 XLSX를 생성하지 못했습니다.');
    }
    await fs.copyFile(produced, targetXlsxPath);
  } finally {
    await fs.rm(workDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

/**
 * 서버 시작 시 프로필 풀을 미리 웜업해 첫 변환의 콜드 스타트(수 초)를 없앤다.
 * 각 슬롯에서 초소형 문서를 한 번씩 변환해 LibreOffice 프로필을 생성해 둔다.
 */
export async function warmUpProfiles(): Promise<void> {
  if (!sofficePath) return;
  await Promise.all(
    Array.from({ length: PROFILE_POOL_SIZE }, async (_, index) => {
      const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-warmup-'));
      try {
        const seed = path.join(workDir, `warmup-${index}.csv`);
        await fs.writeFile(seed, 'warmup\n', 'utf8');
        await convertWithSoffice(seed, path.join(workDir, 'warmup.pdf'), ['--convert-to', 'pdf']);
      } catch {
        /* 웜업 실패는 무시 — 실제 변환 시 콜드 스타트로 대체된다 */
      } finally {
        await fs.rm(workDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
      }
    }),
  );
}

export const officeConverter: Converter = {
  name: 'libreoffice',
  extensions: ['doc', 'docx', 'ppt', 'pptx', 'odt', 'odp', 'rtf', 'txt', 'hwp', 'hwpx'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    const isDocx = path.extname(sourcePath).toLowerCase() === '.docx';
    await convertWithSoffice(
      sourcePath,
      targetPdfPath,
      ['--convert-to', 'pdf'],
      sourceName,
      isDocx ? literalizeDocxThemeFonts : undefined,
    );
  },
};

/**
 * 스프레드시트는 기본 변환 시 A4 인쇄 폭으로 시트가 조각나 화면이 깨져 보인다.
 * SinglePageSheets 옵션으로 "시트 1장 = PDF 1페이지(내용 크기)"로 내보낸다. (LibreOffice 7.2+)
 */
const SPREADSHEET_PDF_FILTER =
  'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}';

/** 세로 스크롤형 리스트 시트가 만드는 초장신 페이지 판정 기준 */
const EXTREME_PAGE_EDGE_PT = 3600;
const EXTREME_PAGE_ASPECT = 4;

async function hasExtremePages(pdfPath: string): Promise<boolean> {
  const doc = await PDFDocument.load(await fs.readFile(pdfPath), {
    ignoreEncryption: true,
    updateMetadata: false,
  });
  return doc.getPages().some((page) => {
    const { width, height } = page.getSize();
    const long = Math.max(width, height);
    const short = Math.min(width, height);
    return long > EXTREME_PAGE_EDGE_PT || long / short > EXTREME_PAGE_ASPECT;
  });
}

/** 이 행 수를 넘는 시트는 시트당 1페이지 방식을 시도하지 않고 곧장 페이지 분할로 간다 */
const LONG_SHEET_ROWS = 500;

/**
 * xlsx/csv는 그리드 뷰(스프레드시트.ts)가 LibreOffice 없이 열람을 담당하므로
 * 업로드 허용 목록에서는 LibreOffice 유무와 무관하게 항상 등록해야 한다.
 * 이 convert()는 "출력용 PDF 지연 생성" 시에만 호출되며, 그 시점에 LibreOffice가
 * 없으면 convertWithSoffice가 명확한 오류를 던진다 (열람 자체는 영향 없음).
 */
export const gridPrintConverter: Converter = {
  name: 'libreoffice-calc-grid-print',
  extensions: ['xlsx', 'csv'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    await spreadsheetConvert(sourcePath, targetPdfPath, sourceName);
  },
};

/** xls/ods는 그리드 뷰가 없어 PDF 변환(LibreOffice)만이 유일한 열람 경로다. */
export const spreadsheetConverter: Converter = {
  name: 'libreoffice-calc',
  extensions: ['xls', 'ods'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    await spreadsheetConvert(sourcePath, targetPdfPath, sourceName);
  },
};

async function spreadsheetConvert(
  sourcePath: string,
  targetPdfPath: string,
  sourceName?: string,
): Promise<void> {
    const isXlsx = path.extname(sourcePath).toLowerCase() === '.xlsx';

    // 수천 행짜리 리스트는 시트당 1페이지가 수 미터짜리 리본이 되어 뿌옇게 보이고,
    // 행높이 재계산 비용도 커서 처음부터 폭 맞춤 페이지 분할로 변환한다.
    if (isXlsx) {
      let longList = false;
      try {
        longList = maxSheetRows(await fs.readFile(sourcePath)) > LONG_SHEET_ROWS;
      } catch {
        /* 추정 실패 시 일반 경로로 진행 */
      }
      if (longList) {
        await convertWithSoffice(
          sourcePath,
          targetPdfPath,
          ['--convert-to', 'pdf'],
          sourceName,
          applyFitToWidth,
        );
        return;
      }
    }

    await convertWithSoffice(
      sourcePath,
      targetPdfPath,
      ['--convert-to', SPREADSHEET_PDF_FILTER],
      sourceName,
      isXlsx ? stripCustomRowHeights : undefined,
    );
    // 예외적으로 여전히 초장신 페이지가 나오면 폭 맞춤 페이지 분할로 재변환한다.
    if (await hasExtremePages(targetPdfPath)) {
      await convertWithSoffice(
        sourcePath,
        targetPdfPath,
        ['--convert-to', 'pdf'],
        sourceName,
        isXlsx ? applyFitToWidthAndStrip : undefined,
      );
    }
}

export const dxfConverter: Converter = {
  name: 'libreoffice-dxf',
  extensions: ['dxf'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    await convertWithSoffice(
      sourcePath,
      targetPdfPath,
      ['--infilter=DXF', '--convert-to', 'pdf:draw_pdf_Export'],
      sourceName,
    );
    // 도면이 거대한 빈 페이지에 작게 놓이는 문제를 내용 영역 크롭으로 보정한다.
    // 크롭은 표시 품질 개선일 뿐이므로 실패해도 변환 자체를 실패시키지 않는다.
    try {
      await trimPdfMargins(targetPdfPath);
    } catch (error) {
      console.warn('도면 여백 크롭 실패 (원본 페이지 그대로 표시):', error);
    }
  },
};
