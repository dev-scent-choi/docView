import { describe, expect, test } from 'vitest';
import { detectNearbyTextStyle } from '../src/textStyle.js';
import type { PageText, PageTextItem } from '../src/search.js';

function item(overrides: Partial<PageTextItem> = {}): PageTextItem {
  return { str: 'regular text', x: 100, y: 100, w: 84, h: 12, ...overrides };
}

function page(items: PageTextItem[], width = 600, height = 800): PageText {
  return { width, height, items };
}

describe('detectNearbyTextStyle', () => {
  test('returns the default style with matched:false on a page with no text', () => {
    const result = detectNearbyTextStyle(page([]), 0.5, 0.5);
    expect(result.matched).toBe(false);
    expect(result.fontSizePt).toBe(12);
    expect(result.bold).toBe(false);
    expect(result.baselineY).toBe(0.5);
  });

  test('matches the nearest line and derives fontSizePt/baselineY from it', () => {
    const line = item({ str: 'hello there friend', x: 100, y: 200, w: 120, h: 14 });
    const p = page([line]);
    // 클릭 지점: 그 줄의 x/y 범위 안, 페이지 비율로 변환
    const xRatio = (100 + 20) / p.width;
    const yRatio = (200 + 5) / p.height;
    const result = detectNearbyTextStyle(p, xRatio, yRatio);
    expect(result.matched).toBe(true);
    expect(result.fontSizePt).toBeGreaterThan(0);
    expect(result.baselineY).toBeCloseTo((line.y + line.h) / p.height, 5);
  });

  test('falls back to the default style when the click is far from any text', () => {
    const line = item({ x: 100, y: 100, w: 80, h: 12 });
    const p = page([line]);
    const result = detectNearbyTextStyle(p, 0.9, 0.9);
    expect(result.matched).toBe(false);
    expect(result.fontSizePt).toBe(12);
  });

  test('classifies a visually denser (bold) line as bold relative to the page median', () => {
    const regularLines = [
      item({ str: 'aaaaaaaaaa', x: 50, y: 50, w: 60, h: 12 }),
      item({ str: 'bbbbbbbbbb', x: 50, y: 80, w: 60, h: 12 }),
      item({ str: 'cccccccccc', x: 50, y: 110, w: 60, h: 12 }),
    ];
    // 같은 글자수·같은 폰트 크기인데 폭이 눈에 띄게 넓은 줄 = 볼드로 추정
    const boldLine = item({ str: 'dddddddddd', x: 50, y: 300, w: 84, h: 12 });
    const p = page([...regularLines, boldLine]);

    const boldResult = detectNearbyTextStyle(
      p,
      (boldLine.x + 10) / p.width,
      (boldLine.y + 5) / p.height,
    );
    expect(boldResult.matched).toBe(true);
    expect(boldResult.bold).toBe(true);

    const regularResult = detectNearbyTextStyle(
      p,
      (regularLines[0].x + 10) / p.width,
      (regularLines[0].y + 5) / p.height,
    );
    expect(regularResult.bold).toBe(false);
  });

  test('does not classify as bold when there are too few reference lines to compare against', () => {
    const onlyLine = item({ str: 'lonely line', x: 50, y: 50, w: 84, h: 12 });
    const p = page([onlyLine]);
    const result = detectNearbyTextStyle(p, (onlyLine.x + 10) / p.width, (onlyLine.y + 5) / p.height);
    expect(result.bold).toBe(false);
  });
});
