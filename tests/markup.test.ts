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

  test('addMarkup with kind omitted defaults to a comment item', async () => {
    const { addMarkup } = await freshMarkupModule();
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
    expect(item.kind).toBe('comment');
  });

  test('addMarkup with kind:text stores fontSizePt/bold and no w/h, listMarkup returns it', async () => {
    const { addMarkup, listMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      kind: 'text',
      page: 1,
      x: 0.3,
      y: 0.4,
      fontSizePt: 14,
      bold: true,
      text: '삽입된 텍스트',
      author: '업로더',
    });
    if (item.kind !== 'text') throw new Error('expected a text markup item');
    expect(item.fontSizePt).toBe(14);
    expect(item.bold).toBe(true);
    expect(item.text).toBe('삽입된 텍스트');
    expect(item.fontFamily).toBe('sans-serif');
    expect('w' in item).toBe(false);

    const items = await listMarkup(meta.id);
    expect(items).toHaveLength(1);
    expect(items[0]).toEqual(item);
  });

  test('addMarkup with kind:text stores an explicit fontFamily', async () => {
    const { addMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      kind: 'text',
      page: 1,
      x: 0,
      y: 0,
      fontSizePt: 12,
      bold: false,
      fontFamily: 'serif',
      text: 'x',
      author: 'a',
    });
    if (item.kind !== 'text') throw new Error('expected a text markup item');
    expect(item.fontFamily).toBe('serif');
  });

  test('addMarkup with kind:text rejects an invalid fontFamily', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, {
        kind: 'text',
        page: 1,
        x: 0,
        y: 0,
        fontSizePt: 12,
        bold: false,
        fontFamily: 'comic-sans' as never,
        text: 'x',
        author: 'a',
      }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('listMarkup backfills fontFamily:sans-serif for text items saved before that field existed', async () => {
    const { listMarkup } = await freshMarkupModule();
    const { store } = await import('../src/storage.js');
    const meta = makeMeta();
    const legacyTextItem = {
      id: 'legacy-text-1',
      kind: 'text',
      page: 1,
      x: 0.2,
      y: 0.2,
      fontSizePt: 12,
      bold: false,
      text: '레거시 텍스트',
      author: 'a',
      createdAt: new Date().toISOString(),
    };
    await store.saveDerived(meta.id, 'markup', Buffer.from(JSON.stringify([legacyTextItem]), 'utf8'));
    const items = await listMarkup(meta.id);
    expect(items).toHaveLength(1);
    expect(items[0].kind).toBe('text');
    if (items[0].kind !== 'text') throw new Error('expected a text markup item');
    expect(items[0].fontFamily).toBe('sans-serif');
  });

  test('addMarkup with kind:text rejects a non-positive font size', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, {
        kind: 'text',
        page: 1,
        x: 0,
        y: 0,
        fontSizePt: 0,
        bold: false,
        text: 'x',
        author: 'a',
      }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('addMarkup with kind:text rejects an absurdly large font size', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, {
        kind: 'text',
        page: 1,
        x: 0,
        y: 0,
        fontSizePt: 5000,
        bold: false,
        text: 'x',
        author: 'a',
      }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('addMarkup with kind:text rejects empty text', async () => {
    const { addMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    await expect(
      addMarkup(meta.id, meta, {
        kind: 'text',
        page: 1,
        x: 0,
        y: 0,
        fontSizePt: 12,
        bold: false,
        text: '   ',
        author: 'a',
      }),
    ).rejects.toThrow(MarkupValidationError);
  });

  test('listMarkup backfills kind:comment for legacy items saved before kind existed', async () => {
    const { listMarkup } = await freshMarkupModule();
    const { store } = await import('../src/storage.js');
    const meta = makeMeta();
    const legacyItem = {
      id: 'legacy-1',
      page: 1,
      x: 0.1,
      y: 0.1,
      w: 0.2,
      h: 0.1,
      text: '레거시 코멘트',
      author: 'a',
      createdAt: new Date().toISOString(),
    };
    await store.saveDerived(meta.id, 'markup', Buffer.from(JSON.stringify([legacyItem]), 'utf8'));
    const items = await listMarkup(meta.id);
    expect(items).toHaveLength(1);
    expect(items[0].kind).toBe('comment');
    expect(items[0].id).toBe('legacy-1');
  });

  test('updateMarkup moves an item to a new position and persists it', async () => {
    const { addMarkup, updateMarkup, listMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      kind: 'text',
      page: 1,
      x: 0.1,
      y: 0.1,
      fontSizePt: 12,
      bold: false,
      text: 'x',
      author: 'a',
    });
    const updated = await updateMarkup(meta.id, item.id, { x: 0.5, y: 0.6 });
    expect(updated?.x).toBe(0.5);
    expect(updated?.y).toBe(0.6);
    expect(updated?.text).toBe('x');

    const items = await listMarkup(meta.id);
    expect(items[0].x).toBe(0.5);
    expect(items[0].y).toBe(0.6);
  });

  test('updateMarkup returns null for an unknown id', async () => {
    const { updateMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    expect(await updateMarkup(meta.id, 'nope', { x: 0.1, y: 0.1 })).toBeNull();
  });

  test('updateMarkup rejects a position that would push a comment rectangle off the page', async () => {
    const { addMarkup, updateMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      page: 1,
      x: 0.1,
      y: 0.1,
      w: 0.3,
      h: 0.1,
      text: 'x',
      author: 'a',
    });
    await expect(updateMarkup(meta.id, item.id, { x: 0.9, y: 0.1 })).rejects.toThrow(MarkupValidationError);
  });

  test('updateMarkup changes fontSizePt on a text item and keeps its other fields', async () => {
    const { addMarkup, updateMarkup } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      kind: 'text',
      page: 1,
      x: 0.1,
      y: 0.1,
      fontSizePt: 12,
      bold: true,
      fontFamily: 'serif',
      text: 'x',
      author: 'a',
    });
    const updated = await updateMarkup(meta.id, item.id, { x: 0.1, y: 0.1, fontSizePt: 18 });
    if (updated?.kind !== 'text') throw new Error('expected a text markup item');
    expect(updated.fontSizePt).toBe(18);
    expect(updated.bold).toBe(true);
    expect(updated.fontFamily).toBe('serif');
  });

  test('updateMarkup rejects an out-of-range fontSizePt', async () => {
    const { addMarkup, updateMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      kind: 'text',
      page: 1,
      x: 0.1,
      y: 0.1,
      fontSizePt: 12,
      bold: false,
      text: 'x',
      author: 'a',
    });
    await expect(updateMarkup(meta.id, item.id, { x: 0.1, y: 0.1, fontSizePt: 0 })).rejects.toThrow(
      MarkupValidationError,
    );
  });

  test('updateMarkup rejects fontSizePt on a comment item', async () => {
    const { addMarkup, updateMarkup, MarkupValidationError } = await freshMarkupModule();
    const meta = makeMeta();
    const item = await addMarkup(meta.id, meta, {
      page: 1,
      x: 0.1,
      y: 0.1,
      w: 0.2,
      h: 0.1,
      text: 'x',
      author: 'a',
    });
    await expect(updateMarkup(meta.id, item.id, { x: 0.1, y: 0.1, fontSizePt: 18 })).rejects.toThrow(
      MarkupValidationError,
    );
  });
});
