import fontkit from '@pdf-lib/fontkit';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { PDFDocument, StandardFonts, degrees, rgb } from 'pdf-lib';
import { pdf } from 'pdf-to-img';
import { SizedLruCache } from './cache.js';
import { config } from './config.js';
import type { MarkupItem } from './markup.js';

/** 프로젝트 루트 assets/fonts — src/rendering.ts 위치 기준 상대경로라 실행 cwd와 무관하다. */
const FONTS_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'assets', 'fonts');

/** 실제로 내장 가능한 글꼴 계열 — monospace는 마땅한 한글 모노스페이스 폰트가 없어 sans-serif로 근사한다. */
type EmbeddableFontFamily = 'sans-serif' | 'serif';

const FONT_FILES: Record<EmbeddableFontFamily, { regular: string; bold: string }> = {
  'sans-serif': { regular: 'Pretendard-Regular.ttf', bold: 'Pretendard-Bold.ttf' },
  serif: { regular: 'NanumMyeongjo-Regular.ttf', bold: 'NanumMyeongjo-Bold.ttf' },
};

const fontBytesCache = new Map<EmbeddableFontFamily, { regular: Buffer; bold: Buffer }>();

/**
 * pdf-lib 기본 14 폰트는 WinAnsi만 지원해 한글을 전혀 그리지 못한다. 사용자가 PDF에 직접
 * 삽입하는 텍스트(kind:'text')는 한글일 가능성이 높으므로, 한글 글리프를 포함한 폰트(sans-serif:
 * Pretendard, serif: Nanum Myeongjo)를 fontkit으로 내장해 그린다 — 어느 쪽으로 근사할지는
 * src/textStyle.ts가 클릭 지점 근처 텍스트에서 감지한 fontFamily를 따른다(둘 다 SIL OFL,
 * assets/fonts/LICENSE-*.txt 참고).
 */
async function loadFontBytes(family: EmbeddableFontFamily): Promise<{ regular: Buffer; bold: Buffer }> {
  const cached = fontBytesCache.get(family);
  if (cached) return cached;
  const files = FONT_FILES[family];
  const [regular, bold] = await Promise.all([
    fs.readFile(path.join(FONTS_DIR, files.regular)),
    fs.readFile(path.join(FONTS_DIR, files.bold)),
  ]);
  const bytes = { regular, bold };
  fontBytesCache.set(family, bytes);
  return bytes;
}

function embeddableFamilyOf(fontFamily: string): EmbeddableFontFamily {
  return fontFamily === 'serif' ? 'serif' : 'sans-serif';
}

type PdfImageDocument = Awaited<ReturnType<typeof pdf>>;
export type BytesProvider = () => Promise<Buffer>;

export interface PageSize {
  readonly width: number;
  readonly height: number;
}

/** getPageSizes가 매번 PDF를 새로 파싱하지 않도록, 영구 저장된 페이지 크기를 읽고/쓰는
 * 콜백 쌍 — loadBytes와 같은 기존 패턴대로 이 파일은 storage.ts를 직접 import하지 않고
 * 항상 호출자(routes/*.ts)가 만들어 주입한다. */
export interface PersistedPageSizes {
  readonly load: () => Promise<readonly PageSize[] | null>;
  readonly save: (sizes: readonly PageSize[]) => Promise<void>;
}

/** 페이지당 렌더링 픽셀 상한 — 초대형 시트/도면에서 메모리 폭주를 막는다. */
const MAX_BITMAP_EDGE_PX = 6000;
/* 실질적인 안전장치는 페이지 픽셀 상한(MAX_BITMAP_EDGE_PX)이며,
   작은 페이지(크롭된 도면 등)는 높은 배율로 렌더링해야 확대 시 선명하다 */
export const MAX_RENDER_SCALE = 16;

interface CachedRenderDoc {
  readonly doc: PdfImageDocument;
  readonly estimatedBytes: number;
}

/** 렌더링 문서 캐시가 예산의 대부분을 차지하므로 캐시 예산의 65%를 할당한다 */
const documentCache = new SizedLruCache<CachedRenderDoc>({
  maxBytes: Math.floor(config.cacheBudgetBytes * 0.65),
  maxEntries: 8,
  sizeOf: (entry) => entry.estimatedBytes,
});

/** 페이지 크기 배열은 작으므로 개수만 제한한다 */
const pageSizeCache = new SizedLruCache<readonly PageSize[]>({
  maxBytes: Number.POSITIVE_INFINITY,
  maxEntries: 32,
  sizeOf: () => 0,
});

export class PageOutOfRangeError extends Error {
  constructor(pageNumber: number, pageCount: number) {
    super(`페이지 범위를 벗어났습니다: ${pageNumber} (전체 ${pageCount}페이지)`);
    this.name = 'PageOutOfRangeError';
  }
}

/** pdfjs가 전달받은 버퍼 소유권을 가져갈 수 있으므로 항상 복사본을 넘긴다. */
function toOwnedBytes(data: Buffer): Uint8Array {
  return new Uint8Array(data);
}

async function openDocument(
  docId: string,
  scale: number,
  loadBytes: BytesProvider,
): Promise<PdfImageDocument> {
  const key = `${docId}#${scale.toFixed(2)}`;
  const cached = await documentCache.getOrLoad(key, async () => {
    const bytes = await loadBytes();
    const doc = await pdf(toOwnedBytes(bytes), { scale });
    // pdf-to-img는 페이지를 배율²에 비례한 캔버스로 렌더링한다. 원본 크기×배율²을
    // 상대적 메모리 순위 추정치로 삼아(상한 16배) 큰 문서·고배율부터 제거한다.
    const estimatedBytes = bytes.length * Math.min(scale * scale, 16);
    return { doc, estimatedBytes };
  });
  return cached.doc;
}

function getPageSizes(
  docId: string,
  loadBytes: BytesProvider,
  persisted: PersistedPageSizes,
): Promise<readonly PageSize[]> {
  return pageSizeCache.getOrLoad(docId, async () => {
    const existing = await persisted.load();
    if (existing) return existing;
    const bytes = await loadBytes();
    const { sizes } = await getPdfPageInfo(bytes);
    await persisted.save(sizes);
    return sizes;
  });
}

/** 페이지 크기 목록(pt) — 뷰어가 로딩 전 프레임 크기를 고정해 정확한 스크롤 점프를 보장한다 */
export async function getPageDimensions(
  docId: string,
  loadBytes: BytesProvider,
  persisted: PersistedPageSizes,
): Promise<ReadonlyArray<{ w: number; h: number }>> {
  const sizes = await getPageSizes(docId, loadBytes, persisted);
  return sizes.map((size) => ({ w: Math.round(size.width), h: Math.round(size.height) }));
}

export interface PdfPageInfo {
  readonly pageCount: number;
  readonly sizes: readonly PageSize[];
}

/**
 * PDF 바이트를 한 번만 파싱해 페이지 수와 페이지별 크기를 함께 얻는다. 변환 직후(업로드,
 * 동시편집 저장 후 재변환) routes/documents.ts의 convertToPdf가 이 함수로 페이지 수를 얻는
 * 김에 페이지 크기도 함께 영구 저장해 두면, 조회 시점(getPageSizes)엔 다시 파싱할 필요가 없다.
 */
export async function getPdfPageInfo(pdfBytes: Buffer): Promise<PdfPageInfo> {
  const doc = await PDFDocument.load(toOwnedBytes(pdfBytes), {
    ignoreEncryption: true,
    updateMetadata: false,
  });
  const sizes = doc.getPages().map((page) => {
    const { width, height } = page.getSize();
    return { width, height };
  });
  return { pageCount: doc.getPageCount(), sizes };
}

/**
 * 문서가 새 원본으로부터 재변환된 뒤(예: docx 동시편집 저장) 호출한다.
 * 렌더링 문서 캐시(배율별)와 페이지 크기 캐시에 남아 있는 이전 PDF 기준 결과를 모두 지워
 * 다음 조회가 새로 변환된 PDF를 다시 읽게 한다.
 */
export function invalidateRenderCache(docId: string): void {
  documentCache.invalidatePrefix(`${docId}#`);
  pageSizeCache.invalidate(docId);
}

/**
 * 요청 배율을 페이지 크기 기준 픽셀 상한에 맞게 낮춘다.
 * 도면처럼 원본이 벡터인 PDF는 배율을 올릴수록 선명해진다.
 */
function clampScale(requestedScale: number, size: PageSize): number {
  const bounded = Math.min(Math.max(requestedScale, 1), MAX_RENDER_SCALE);
  const longestEdge = Math.max(size.width, size.height);
  const maxByPixels = MAX_BITMAP_EDGE_PX / longestEdge;
  const clamped = Math.min(bounded, maxByPixels);
  // 픽셀 상한보다 큰 초대형 페이지는 1 미만 배율로 낮춰서라도 상한을 지킨다.
  // 캐시 키가 무한히 늘지 않도록 0.25 단위로 정량화한다.
  return Math.max(0.25, Math.floor(clamped * 4) / 4);
}

export async function renderPage(
  docId: string,
  pageNumber: number,
  requestedScale: number,
  loadBytes: BytesProvider,
  persisted: PersistedPageSizes,
): Promise<Buffer> {
  const sizes = await getPageSizes(docId, loadBytes, persisted);
  if (pageNumber < 1 || pageNumber > sizes.length) {
    throw new PageOutOfRangeError(pageNumber, sizes.length);
  }
  const scale = clampScale(requestedScale ?? config.renderScale, sizes[pageNumber - 1]);
  const document = await openDocument(docId, scale, loadBytes);
  return document.getPage(pageNumber);
}

/** Helvetica는 WinAnsi 범위만 인코딩 가능하므로 워터마크 문자열을 ASCII로 정리한다. */
function toWinAnsiSafe(text: string): string {
  const safe = text.replace(/[^\x20-\x7E]/g, '').trim();
  return safe.length > 0 ? safe : 'CONFIDENTIAL';
}

export async function buildPrintPdf(pdfBytes: Buffer, watermarkText: string): Promise<Uint8Array> {
  const doc = await PDFDocument.load(toOwnedBytes(pdfBytes), { ignoreEncryption: true });
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const text = toWinAnsiSafe(watermarkText);
  for (const page of doc.getPages()) {
    const { width, height } = page.getSize();
    const fontSize = Math.max(18, Math.min(width, height) / 14);
    const textWidth = font.widthOfTextAtSize(text, fontSize);
    page.drawText(text, {
      x: width / 2 - textWidth / 2.5,
      y: height / 3,
      size: fontSize,
      font,
      color: rgb(0.55, 0.55, 0.55),
      opacity: 0.3,
      rotate: degrees(35),
    });
  }
  return doc.save();
}

/**
 * 마크업(사각형+코멘트)을 PDF에 그려 넣는다. 원본 파일은 건드리지 않고, 출력/다운로드용으로
 * 새로 만드는 PDF 바이트에만 적용한다(호출부: src/routes/viewer.ts).
 * 좌표는 페이지 크기와 무관한 0~1 비율(src/markup.ts)이라 이 함수가 실제 pt 단위로 환산한다.
 */
export async function burnMarkupIntoPdf(
  pdfBytes: Buffer,
  items: readonly MarkupItem[],
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(toOwnedBytes(pdfBytes), { ignoreEncryption: true });
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const pages = doc.getPages();

  // text 항목이 실제로 쓰는 글꼴 계열만(sans-serif/serif) 내장한다 — 코멘트만 있거나 한쪽
  // 계열만 쓰는 문서에서 불필요한 폰트를 내장하는 비용을 피한다.
  // subset:true는 쓰지 않는다 — @pdf-lib/fontkit의 서브셋터가 이 폰트들의 일부 한글 글리프를
  // 통째로 빠뜨리거나 잘못된 위치에 그리는 버그가 있음을 실제 렌더링 결과로 확인했다(예:
  // "이름:"이 "름"만 남고 "이"가 사라짐). 전체 폰트를 그대로 내장하면(용량은 커지지만) 정확하다.
  const textFamiliesUsed = new Set<EmbeddableFontFamily>();
  for (const item of items) {
    if (item.kind === 'text') textFamiliesUsed.add(embeddableFamilyOf(item.fontFamily));
  }
  const textFonts = new Map<
    EmbeddableFontFamily,
    { regular: Awaited<ReturnType<typeof doc.embedFont>>; bold: Awaited<ReturnType<typeof doc.embedFont>> }
  >();
  if (textFamiliesUsed.size > 0) {
    doc.registerFontkit(fontkit);
    for (const family of textFamiliesUsed) {
      const { regular, bold } = await loadFontBytes(family);
      const [regularFont, boldFont] = await Promise.all([
        doc.embedFont(regular, { subset: false }),
        doc.embedFont(bold, { subset: false }),
      ]);
      textFonts.set(family, { regular: regularFont, bold: boldFont });
    }
  }

  const byPage = new Map<number, MarkupItem[]>();
  for (const item of items) {
    const list = byPage.get(item.page) ?? [];
    list.push(item);
    byPage.set(item.page, list);
  }

  const fontSize = 8;
  const lineHeight = fontSize * 1.3;
  const boxPad = 3;
  const maxCharsPerLine = 42;

  for (const [pageNumber, pageItems] of byPage) {
    const page = pages[pageNumber - 1];
    if (!page) continue; // 범위를 벗어난 페이지 번호는 조용히 건너뛴다
    const { width, height } = page.getSize();

    for (const item of pageItems) {
      if (item.kind === 'text') {
        // x,y는 텍스트 베이스라인 좌측 시작점의 비율 좌표(위에서부터) — pdf-lib은 아래가
        // 원점이라 세로만 뒤집는다. 가로는 뒤집을 필요 없이 그대로 곱한다.
        const fonts = textFonts.get(embeddableFamilyOf(item.fontFamily));
        if (!fonts) continue;
        const chosenFont = item.bold ? fonts.bold : fonts.regular;
        page.drawText(item.text, {
          x: item.x * width,
          y: height - item.y * height,
          size: item.fontSizePt,
          font: chosenFont,
          color: rgb(0, 0, 0),
        });
        continue;
      }

      const boxX = item.x * width;
      const boxW = item.w * width;
      const boxH = item.h * height;
      // pdf-lib 좌표계는 좌하단이 원점이라, 위에서부터 재는 비율 y를 뒤집는다.
      const boxY = height - (item.y + item.h) * height;

      page.drawRectangle({
        x: boxX,
        y: boxY,
        width: boxW,
        height: boxH,
        borderColor: rgb(0.91, 0.42, 0.02),
        borderWidth: 1.5,
        color: rgb(0.91, 0.42, 0.02),
        opacity: 0.08,
        borderOpacity: 0.9,
      });

      const text = toWinAnsiSafe(item.text);
      const lines: string[] = [];
      for (let i = 0; i < text.length; i += maxCharsPerLine) {
        lines.push(text.slice(i, i + maxCharsPerLine));
      }
      const textBlockHeight = lines.length * lineHeight + boxPad * 2;
      const textY = Math.max(0, boxY - textBlockHeight);
      const textWidthPx =
        Math.max(boxW, ...lines.map((line) => font.widthOfTextAtSize(line, fontSize))) + boxPad * 2;

      page.drawRectangle({
        x: boxX,
        y: textY,
        width: textWidthPx,
        height: textBlockHeight,
        color: rgb(1, 0.97, 0.88),
        borderColor: rgb(0.91, 0.42, 0.02),
        borderWidth: 1,
        opacity: 0.95,
      });

      lines.forEach((line, i) => {
        page.drawText(line, {
          x: boxX + boxPad,
          y: textY + textBlockHeight - boxPad - (i + 1) * lineHeight + (lineHeight - fontSize) / 2,
          size: fontSize,
          font,
          color: rgb(0.2, 0.15, 0.02),
        });
      });
    }
  }

  return doc.save();
}
