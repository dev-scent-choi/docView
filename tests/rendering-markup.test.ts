import { PDFDocument } from 'pdf-lib';
import { describe, expect, test } from 'vitest';
import { burnMarkupIntoPdf } from '../src/rendering.js';
import type { CommentMarkupItem, TextMarkupItem } from '../src/markup.js';

async function makeTwoPagePdf(): Promise<Buffer> {
  const doc = await PDFDocument.create();
  doc.addPage([200, 300]);
  doc.addPage([200, 300]);
  return Buffer.from(await doc.save());
}

function makeItem(overrides: Partial<CommentMarkupItem> = {}): CommentMarkupItem {
  return {
    id: 'm1',
    kind: 'comment',
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

function makeTextItem(overrides: Partial<TextMarkupItem> = {}): TextMarkupItem {
  return {
    id: 't1',
    kind: 'text',
    page: 1,
    x: 0.1,
    y: 0.1,
    fontSizePt: 14,
    bold: false,
    fontFamily: 'sans-serif',
    text: '삽입 텍스트',
    author: '업로더',
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

  test('burns a kind:text item with Korean text without throwing, growing the file', async () => {
    const original = await makeTwoPagePdf();
    const result = await burnMarkupIntoPdf(original, [makeTextItem({ text: '한글 텍스트 삽입 테스트' })]);
    const reopened = await PDFDocument.load(result);
    expect(reopened.getPageCount()).toBe(2);
    expect(result.length).toBeGreaterThan(original.length);
  });

  test('renders bold and regular text items using different embedded fonts without error', async () => {
    const original = await makeTwoPagePdf();
    const items = [
      makeTextItem({ id: 't1', bold: false, text: '일반체' }),
      makeTextItem({ id: 't2', bold: true, text: '굵은체', y: 0.3 }),
    ];
    await expect(burnMarkupIntoPdf(original, items)).resolves.toBeInstanceOf(Uint8Array);
  });

  test('renders a serif text item using the Nanum Myeongjo font without error', async () => {
    const original = await makeTwoPagePdf();
    const items = [makeTextItem({ fontFamily: 'serif', text: '명조체 테스트' })];
    await expect(burnMarkupIntoPdf(original, items)).resolves.toBeInstanceOf(Uint8Array);
  });

  test('renders sans-serif and serif text items together without error', async () => {
    const original = await makeTwoPagePdf();
    const items = [
      makeTextItem({ id: 't1', fontFamily: 'sans-serif', text: '고딕체', y: 0.2 }),
      makeTextItem({ id: 't2', fontFamily: 'serif', text: '명조체', y: 0.4 }),
    ];
    await expect(burnMarkupIntoPdf(original, items)).resolves.toBeInstanceOf(Uint8Array);
  });

  test('falls back to sans-serif for monospace, since no Korean monospace font is bundled', async () => {
    const original = await makeTwoPagePdf();
    const items = [makeTextItem({ fontFamily: 'monospace', text: '고정폭 근사' })];
    await expect(burnMarkupIntoPdf(original, items)).resolves.toBeInstanceOf(Uint8Array);
  });

  test('handles a mix of comment and text items on the same page', async () => {
    const original = await makeTwoPagePdf();
    const items = [makeItem({ id: 'c1' }), makeTextItem({ id: 't1' })];
    const result = await burnMarkupIntoPdf(original, items);
    const reopened = await PDFDocument.load(result);
    expect(reopened.getPageCount()).toBe(2);
  });

  test('an out-of-range text item is skipped without throwing', async () => {
    const original = await makeTwoPagePdf();
    await expect(
      burnMarkupIntoPdf(original, [makeTextItem({ page: 99 })]),
    ).resolves.toBeInstanceOf(Uint8Array);
  });
});
