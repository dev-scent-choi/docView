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
    await fs.rm(profileDir, { recursive: true, force: true }).catch(() => {});
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
    await fs.rm(workDir, { recursive: true, force: true });
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
        await fs.rm(workDir, { recursive: true, force: true });
      }
    }),
  );
}

export const officeConverter: Converter = {
  name: 'libreoffice',
  extensions: ['doc', 'docx', 'ppt', 'pptx', 'odt', 'odp', 'rtf', 'txt'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    await convertWithSoffice(sourcePath, targetPdfPath, ['--convert-to', 'pdf'], sourceName);
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
