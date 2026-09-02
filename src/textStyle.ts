import type { PageText, PageTextItem } from './search.js';

export interface DetectedTextStyle {
  readonly fontSizePt: number;
  readonly bold: boolean;
  /** 텍스트가 놓일 베이스라인의 페이지 상단 기준 비율(0~1) — 클릭 지점 근처 줄에 자동 정렬한다. */
  readonly baselineY: number;
  /** 근처에 실제 텍스트가 있어서 감지된 값인지, 없어서 기본값을 쓴 것인지 */
  readonly matched: boolean;
}

const DEFAULT_FONT_SIZE_PT = 12;
/** 가장 가까운 줄의 높이 대비 이 배수보다 멀면 "근처"로 보지 않는다 */
const PROXIMITY_LINE_HEIGHT_MULTIPLIER = 2;
/** PDF 텍스트 박스 높이는 보통 실제 글자 크기보다 다소 크므로 보정한다 */
const FONT_SIZE_FROM_HEIGHT_RATIO = 0.8;
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
 * 클릭 지점(0~1 비율)에서 가장 가까운 기존 텍스트 줄을 찾아 글자 크기·베이스라인·볼드 여부를
 * 추정한다. pdfjs의 텍스트 레이어는 폰트 굵기 정보를 노출하지 않으므로(styles.fontFamily는
 * "sans-serif" 같은 일반 이름만 준다), 같은 페이지의 문자당 평균 폭(밀도)과 비교하는 통계적
 * 근사로 볼드를 추정한다 — 완벽하지 않지만 실제 볼드/일반 텍스트를 상당히 잘 구분한다
 * (검증: Helvetica vs Helvetica-Bold로 그린 텍스트가 문자당 폭이 약 16% 더 넓게 나왔다).
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
    return { fontSizePt: DEFAULT_FONT_SIZE_PT, bold: false, baselineY: yRatio, matched: false };
  }

  const fontSizePt = Math.max(1, best.h * FONT_SIZE_FROM_HEIGHT_RATIO);
  const baselineY = (best.y + best.h) / pageText.height;

  const densities = pageText.items
    .map(widthDensity)
    .filter((value): value is number => value !== null);
  const bestDensity = widthDensity(best);
  const bold =
    bestDensity !== null && densities.length >= MIN_REFERENCE_LINES
      ? bestDensity >= median(densities) * BOLD_DENSITY_RATIO
      : false;

  return { fontSizePt, bold, baselineY, matched: true };
}
