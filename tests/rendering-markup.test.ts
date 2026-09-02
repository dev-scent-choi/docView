import { PDFDocument } from 'pdf-lib';
import { describe, expect, test } from 'vitest';
import { burnMarkupIntoPdf } from '../src/rendering.js';
import type { MarkupItem } from '../src/markup.js';

async function makeTwoPagePdf(): Promise<Buffer> {
  const doc = await PDFDocument.create();
  doc.addPage([200, 300]);
  doc.addPage([200, 300]);
  return Buffer.from(await doc.save());
}

function makeItem(overrides: Partial<MarkupItem> = {}): MarkupItem {
  return {
    id: 'm1',
    page: 1,
    x: 0.1,
    y: 0.1,
    w: 0.3,
    h: 0.1,
    text: '이 부분을 수정해주세요',
    author: '검토자A',
    createdAt: new Date().toISOString(),
    ...overrides,
  };
}

describe('burnMarkupIntoPdf', () => {
  test('returns a valid PDF with the same page count when no items are given', async () => {
    const original = await makeTwoPagePdf();
    const result = await burnMarkupIntoPdf(original, []);
    const reopened = await PDFDocument.load(result);
    expect(reopened.getPageCount()).toBe(2);
  });

  test('burns a markup item into the correct page without changing page count', async () => {
    const original = await makeTwoPagePdf();
    const result = await burnMarkupIntoPdf(original, [makeItem({ page: 2 })]);
    const reopened = await PDFDocument.load(result);
    expect(reopened.getPageCount()).toBe(2);
    // 원본보다 산출물이 확실히 커야 한다(사각형+텍스트가 실제로 그려졌다는 근거)
    expect(result.length).toBeGreaterThan(original.length);
  });

  test('handles multiple items across multiple pages', async () => {
    const original = await makeTwoPagePdf();
    const items = [
      makeItem({ id: 'm1', page: 1, text: '첫 페이지 코멘트' }),
      makeItem({ id: 'm2', page: 2, text: '둘째 페이지 코멘트' }),
    ];
    const result = await burnMarkupIntoPdf(original, items);
    const reopened = await PDFDocument.load(result);
    expect(reopened.getPageCount()).toBe(2);
  });

  test('ignores an item whose page number is out of range instead of throwing', async () => {
    const original = await makeTwoPagePdf();
    await expect(burnMarkupIntoPdf(original, [makeItem({ page: 99 })])).resolves.toBeInstanceOf(
      Uint8Array,
    );
  });
});
