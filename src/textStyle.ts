import type { PageText, PageTextItem, PdfFontFamily } from './search.js';

export interface DetectedTextStyle {
  readonly fontSizePt: number;
  readonly bold: boolean;
  readonly fontFamily: PdfFontFamily;
  /** 텍스트가 놓일 베이스라인의 페이지 상단 기준 비율(0~1) — 클릭 지점 근처 줄에 자동 정렬한다. */
  readonly baselineY: number;
  /** 근처에 실제 텍스트가 있어서 감지된 값인지, 없어서 기본값을 쓴 것인지 */
  readonly matched: boolean;
}

const DEFAULT_FONT_SIZE_PT = 12;
const DEFAULT_FONT_FAMILY: PdfFontFamily = 'sans-serif';
/** 가장 가까운 줄의 높이 대비 이 배수보다 멀면 "근처"로 보지 않는다 */
const PROXIMITY_LINE_HEIGHT_MULTIPLIER = 2;
/** 볼드 판정 시 페이지 중앙값 대비 필요한 최소 밀도 배수 */
const BOLD_DENSITY_RATIO = 1.08;
/** 밀도를 비교할 만한 최소 참고 줄 수 — 이보다 적으면 볼드 판정을 하지 않는다 */
const MIN_REFERENCE_LINES = 3;

function widthDensity(item: PageTextItem): number | null {
  const chars = item.str.trim().length;
  if (chars === 0 || item.h <= 0) return null;
  return item.w / chars / item.h;
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

/**
 * 클릭 지점(0~1 비율)에서 가장 가까운 기존 텍스트 줄을 찾아 글자 크기·글꼴 계열·베이스라인·볼드
 * 여부를 추정한다.
 *
 * 글자 크기: pdfjs가 주는 텍스트 박스 높이(item.h)는 PDF 생성 도구·pdf-lib·LibreOffice 변환본
 * 전부에서 실제 pt 크기와 정확히 같다는 것을 실측으로 확인했다(예: 12pt로 그린 텍스트가 항상
 * height:12로 나옴) — 별도 보정 없이 그대로 쓴다.
 *
 * 글꼴 계열: pdfjs 텍스트 레이어의 styles.fontFamily가 "sans-serif"/"serif"/"monospace" 중
 * 하나로 일반화해서 알려주므로(src/search.ts의 normalizeFontFamily), 실제 폰트를 재사용할 수는
 * 없어도 최소한 세 계열 중 어느 쪽에 가까운지는 정확히 반영한다.
 *
 * 볼드: pdfjs는 폰트 굵기 정보를 노출하지 않으므로(styles.fontFamily는 계열만 알려줌), 같은
 * 페이지의 문자당 평균 폭(밀도)과 비교하는 통계적 근사로 추정한다 — 완벽하지 않지만 실제
 * 볼드/일반 텍스트를 상당히 잘 구분한다(검증: Helvetica-Bold가 Helvetica보다 문자당 폭이 약
 * 16% 더 넓게 나왔다).
 */
export function detectNearbyTextStyle(
  pageText: PageText,
  xRatio: number,
  yRatio: number,
): DetectedTextStyle {
  const clickX = xRatio * pageText.width;
  const clickY = yRatio * pageText.height;

  let best: PageTextItem | null = null;
  let bestDist = Number.POSITIVE_INFINITY;
  for (const candidate of pageText.items) {
    const itemMidY = candidate.y + candidate.h / 2;
    const withinX = clickX >= candidate.x - candidate.w * 0.25 && clickX <= candidate.x + candidate.w * 1.25;
    if (!withinX) continue;
    const dist = Math.abs(itemMidY - clickY);
    if (dist < bestDist) {
      bestDist = dist;
      best = candidate;
    }
  }

  if (!best || bestDist > best.h * PROXIMITY_LINE_HEIGHT_MULTIPLIER) {
    return {
      fontSizePt: DEFAULT_FONT_SIZE_PT,
      bold: false,
      fontFamily: DEFAULT_FONT_FAMILY,
      baselineY: yRatio,
      matched: false,
    };
  }

  const fontSizePt = Math.max(1, best.h);
  const baselineY = (best.y + best.h) / pageText.height;

  const densities = pageText.items
    .map(widthDensity)
    .filter((value): value is number => value !== null);
  const bestDensity = widthDensity(best);
  const bold =
    bestDensity !== null && densities.length >= MIN_REFERENCE_LINES
      ? bestDensity >= median(densities) * BOLD_DENSITY_RATIO
      : false;

  return { fontSizePt, bold, fontFamily: best.fontFamily, baselineY, matched: true };
}
