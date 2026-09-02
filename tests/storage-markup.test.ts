import { mkdtempSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, test, vi } from 'vitest';

async function loadStoreWithTempDir() {
  const dir = mkdtempSync(path.join(os.tmpdir(), 'docview-storage-test-'));
  vi.resetModules();
  process.env.DOCVIEW_DATA_DIR = dir;
  process.env.DOCVIEW_STORAGE = 'disk';
  const storageModule = await import('../src/storage.js');
  await storageModule.ensureDataDirs();
  return { dir, storageModule };
}

describe('storage: markup derived kind', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('markup and grid derived kinds do not collide on disk', async () => {
    const { dir, storageModule } = await loadStoreWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.writeMeta({
      id,
      name: 'a.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
    });
    await store.saveDerived(id, 'grid', Buffer.from('{"grid":true}', 'utf8'));
    await store.saveDerived(id, 'markup', Buffer.from('[]', 'utf8'));
    const grid = await store.readDerived(id, 'grid');
    const markup = await store.readDerived(id, 'markup');
    expect(grid.toString('utf8')).toBe('{"grid":true}');
    expect(markup.toString('utf8')).toBe('[]');
  });

  test('hasDerived reports false before save and true after', async () => {
    const { dir, storageModule } = await loadStoreWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.writeMeta({
      id,
      name: 'a.pdf',
      ext: 'pdf',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
    });
    expect(await store.hasDerived(id, 'markup')).toBe(false);
    await store.saveDerived(id, 'markup', Buffer.from('[]', 'utf8'));
    expect(await store.hasDerived(id, 'markup')).toBe(true);
  });

  test('deleteDocument removes the markup derived file along with original/meta', async () => {
    const { dir, storageModule } = await loadStoreWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.writeMeta({
      id,
      name: 'a.pdf',
      ext: 'pdf',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
    });
    await store.saveOriginal(id, 'pdf', Buffer.from('%PDF-1.4', 'utf8'));
    await store.saveDerived(id, 'markup', Buffer.from('[]', 'utf8'));
    await store.deleteDocument(id);
    expect(await store.readMeta(id)).toBeNull();
    expect(await store.hasDerived(id, 'markup')).toBe(false);
  });
});
