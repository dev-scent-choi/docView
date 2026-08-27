import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

const require = createRequire(import.meta.url);
/** 표준 14폰트 미내장 PDF에서 경고가 뜨지 않도록 pdfjs 동봉 폰트 경로를 지정한다 */
const STANDARD_FONT_DATA_URL = pathToFileURL(
  path.join(path.dirname(require.resolve('pdfjs-dist/package.json')), 'standard_fonts'),
).href.concat('/');

/** PDF 페이지에서 추출한 텍스트 조각 (좌표는 pt, y는 페이지 상단 기준) */
export interface PageTextItem {
  readonly str: string;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export interface PageText {
  readonly width: number;
  readonly height: number;
  readonly items: readonly PageTextItem[];
}

/** 하이라이트 사각형 — 페이지 크기 대비 비율(0~1)로 반환해 클라이언트가 배율과 무관하게 그린다 */
export interface SearchRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export interface PdfSearchHit {
  readonly page: number;
  readonly rects: readonly SearchRect[];
  readonly snippet: string;
}

const DEFAULT_ITEM_HEIGHT = 10;

/**
 * LibreOffice 변환 PDF에는 텍스트 레이어가 살아 있으므로 OCR 없이
 * pdfjs로 페이지별 텍스트와 좌표를 추출한다.
 */
export async function extractPdfText(pdfBytes: Buffer): Promise<PageText[]> {
  const task = getDocument({
    data: new Uint8Array(pdfBytes),
    disableFontFace: true,
    standardFontDataUrl: STANDARD_FONT_DATA_URL,
  });
  const doc = await task.promise;
  try {
    const pages: PageText[] = [];
    for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber += 1) {
      const page = await doc.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1 });
      // 텍스트 좌표는 크롭 전 원지 좌표계이므로 CropBox 원점(view)을 빼서 보정한다
      // (도면 여백 자동 크롭이 적용된 문서에서 하이라이트가 어긋나던 원인)
      const [cropX, cropY] = page.view;
      const content = await page.getTextContent();
      const items: PageTextItem[] = [];
      for (const item of content.items) {
        if (!('str' in item) || item.str.trim() === '') continue;
        const width = item.width || DEFAULT_ITEM_HEIGHT;
        const height = item.height || DEFAULT_ITEM_HEIGHT;
        // 도면의 치수 문자처럼 회전된 텍스트는 transform 행렬로 실제 진행 방향을
        // 구해 네 꼭짓점의 축 정렬 바운딩 박스를 계산해야 하이라이트가 맞는다.
        const [a, b, c, d, e, f] = item.transform;
        const wNorm = Math.hypot(a, b) || 1;
        const hNorm = Math.hypot(c, d) || 1;
        // 행렬의 x축(a,b)·y축(c,d) 방향 그대로 폭·높이를 전개해야
        // 뒤집힘(d<0)·회전이 섞인 행렬에서도 바운딩 박스가 맞는다
        const wx = (a / wNorm) * width;
        const wy = (b / wNorm) * width;
        const hx = (c / hNorm) * height;
        const hy = (d / hNorm) * height;
        const corners: Array<[number, number]> = [
          [e, f],
          [e + wx, f + wy],
          [e + hx, f + hy],
          [e + wx + hx, f + wy + hy],
        ];
        const xs = corners.map((corner) => corner[0]);
        const ys = corners.map((corner) => corner[1]);
        const minX = Math.min(...xs) - cropX;
        const maxY = Math.max(...ys) - cropY;
        items.push({
          str: item.str,
          x: minX,
          y: viewport.height - maxY,
          w: Math.max(...xs) - Math.min(...xs),
          h: maxY - (Math.min(...ys) - cropY),
        });
      }
      pages.push({ width: viewport.width, height: viewport.height, items });
      page.cleanup();
    }
    return pages;
  } finally {
    await task.destroy();
  }
}

/**
 * 페이지 텍스트에서 부분 문자열을 대소문자 구분 없이 찾는다.
 * 질의가 여러 텍스트 조각에 걸쳐 있으면 해당 조각들의 사각형을 모두 반환한다.
 */
export function searchPdfText(
  pages: readonly PageText[],
  query: string,
  limit: number,
): { hits: PdfSearchHit[]; total: number } {
  const needle = query.toLowerCase();
  const hits: PdfSearchHit[] = [];
  let total = 0;

  for (let pageIdx = 0; pageIdx < pages.length; pageIdx += 1) {
    const page = pages[pageIdx];
    // 조각 사이에 공백 1칸을 두고 이어붙인 페이지 전문에서 검색한다
    let joined = '';
    const spans: Array<{ start: number; end: number; item: PageTextItem }> = [];
    for (const item of page.items) {
      const start = joined.length;
      joined += item.str;
      spans.push({ start, end: joined.length, item });
      joined += ' ';
    }
    const haystack = joined.toLowerCase();

    let from = 0;
    for (;;) {
      const at = haystack.indexOf(needle, from);
      if (at === -1) break;
      from = at + Math.max(needle.length, 1);
      total += 1;
      if (hits.length >= limit) continue;

      const matchEnd = at + needle.length;
      const covered = spans.filter((span) => span.end > at && span.start < matchEnd);
      if (covered.length === 0) continue;
      const rects = covered.map(({ item }) => ({
        x: item.x / page.width,
        y: item.y / page.height,
        w: item.w / page.width,
        h: item.h / page.height,
      }));
      const snippetStart = Math.max(0, at - 30);
      const snippet = joined.slice(snippetStart, matchEnd + 30).trim();
      hits.push({ page: pageIdx + 1, rects, snippet });
    }
  }
  return { hits, total };
}
