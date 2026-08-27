import { PDFDocument, StandardFonts, degrees, rgb } from 'pdf-lib';
import { pdf } from 'pdf-to-img';
import { SizedLruCache } from './cache.js';
import { config } from './config.js';

type PdfImageDocument = Awaited<ReturnType<typeof pdf>>;
export type BytesProvider = () => Promise<Buffer>;

interface PageSize {
  readonly width: number;
  readonly height: number;
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

function getPageSizes(docId: string, loadBytes: BytesProvider): Promise<readonly PageSize[]> {
  return pageSizeCache.getOrLoad(docId, async () => {
    const bytes = await loadBytes();
    const doc = await PDFDocument.load(toOwnedBytes(bytes), {
      ignoreEncryption: true,
      updateMetadata: false,
    });
    return doc.getPages().map((page) => {
      const { width, height } = page.getSize();
      return { width, height };
    });
  });
}

/** 페이지 크기 목록(pt) — 뷰어가 로딩 전 프레임 크기를 고정해 정확한 스크롤 점프를 보장한다 */
export async function getPageDimensions(
  docId: string,
  loadBytes: BytesProvider,
): Promise<ReadonlyArray<{ w: number; h: number }>> {
  const sizes = await getPageSizes(docId, loadBytes);
  return sizes.map((size) => ({ w: Math.round(size.width), h: Math.round(size.height) }));
}

export async function getPageCountFromBytes(pdfBytes: Buffer): Promise<number> {
  const doc = await PDFDocument.load(toOwnedBytes(pdfBytes), {
    ignoreEncryption: true,
    updateMetadata: false,
  });
  return doc.getPageCount();
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
): Promise<Buffer> {
  const sizes = await getPageSizes(docId, loadBytes);
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
