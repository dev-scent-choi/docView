import { afterEach, describe, expect, test, vi } from 'vitest';
import type { DocumentMeta } from '../src/types.js';

async function freshMarkupModule() {
  vi.resetModules();
  process.env.DOCVIEW_STORAGE = 'memory';
  return import('../src/markup.js');
}

function makeMeta(overrides: Partial<DocumentMeta> = {}): DocumentMeta {
  return {
    id: '11111111-1111-4111-8111-111111111111',
    name: 'sample.pdf',
    ext: 'pdf',
    status: 'ready',
    pageCount: 3,
    createdAt: new Date().toISOString(),
    ...overrides,
  };
}

describe('markup', () => {
  afterEach(() => {
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
  });

  test('listMarkup on a document with no markup returns an empty array', async () => {
    const { listMarkup } = await freshMarkupModule();
    const items = await listMarkup(makeMeta().id);
    expect(items).toEqual([]);
  });

  test('addMarkup stores a new item with a generated id and timestamp, listMarkup returns it', async () => {
    const { addMarkup, listMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      page: 2,
      x: 0.1,
      y: 0.2,
      w: 0.3,
      h: 0.1,
      text: '이 부분을 수정해주세요',
      author: '검토자A',
    });
    expect(item.id).toBeTruthy();
    expect(item.page).toBe(2);
    expect(item.text).toBe('이 부분을 수정해주세요');
    expect(item.author).toBe('검토자A');
    expect(new Date(item.createdAt).toString()).not.toBe('Invalid Date');

    const items = await listMarkup(meta.id);
    expect(items).toHaveLength(1);
    expect(items[0]).toEqual(item);
  });

  test('addMarkup rejects a page number outside the document range', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta({ pageCount: 3 });
    await expect(
      addMarkup(meta.id, meta, { page: 4, x: 0, y: 0, w: 0.1, h: 0.1, text: 'x', author: 'a' }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('addMarkup rejects a zero-area rectangle', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, { page: 1, x: 0.5, y: 0.5, w: 0, h: 0.1, text: 'x', author: 'a' }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('addMarkup rejects a rectangle that overflows the page bounds', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, { page: 1, x: 0.9, y: 0.1, w: 0.5, h: 0.1, text: 'x', author: 'a' }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('addMarkup rejects empty comment text', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, { page: 1, x: 0, y: 0, w: 0.1, h: 0.1, text: '   ', author: 'a' }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('deleteMarkup removes the item and returns true; returns false for an unknown id', async () => {
    const { addMarkup, deleteMarkup, listMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      page: 1,
      x: 0,
      y: 0,
      w: 0.1,
      h: 0.1,
      text: 'x',
      author: 'a',
    });
    expect(await deleteMarkup(meta.id, item.id)).toBe(true);
    expect(await listMarkup(meta.id)).toEqual([]);
    expect(await deleteMarkup(meta.id, item.id)).toBe(false);
  });

  test('two markups added concurrently on the same document are both persisted (no lost update)', async () => {
    const { addMarkup, listMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    await Promise.all([
      addMarkup(meta.id, meta, { page: 1, x: 0, y: 0, w: 0.1, h: 0.1, text: 'one', author: 'a' }),
      addMarkup(meta.id, meta, { page: 1, x: 0.2, y: 0.2, w: 0.1, h: 0.1, text: 'two', author: 'b' }),
    ]);
    const items = await listMarkup(meta.id);
    expect(items.map((i) => i.text).sort()).toEqual(['one', 'two']);
  });
});
