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

describe('storage: pagesizes/gridindex derived kinds', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('pagesizes and gridindex derived kinds do not collide with each other or with grid/pdf/markup', async () => {
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
    await store.saveDerived(id, 'pagesizes', Buffer.from('[{"width":595,"height":842}]', 'utf8'));
    await store.saveDerived(id, 'gridindex', Buffer.from('{"lineOffsets":[0,10]}', 'utf8'));
    expect((await store.readDerived(id, 'grid')).toString('utf8')).toBe('{"grid":true}');
    expect((await store.readDerived(id, 'pagesizes')).toString('utf8')).toBe(
      '[{"width":595,"height":842}]',
    );
    expect((await store.readDerived(id, 'gridindex')).toString('utf8')).toBe(
      '{"lineOffsets":[0,10]}',
    );
  });

  test('hasDerived reports false before save and true after, for both new kinds', async () => {
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
    expect(await store.hasDerived(id, 'pagesizes')).toBe(false);
    expect(await store.hasDerived(id, 'gridindex')).toBe(false);
    await store.saveDerived(id, 'pagesizes', Buffer.from('[]', 'utf8'));
    await store.saveDerived(id, 'gridindex', Buffer.from('{"lineOffsets":[]}', 'utf8'));
    expect(await store.hasDerived(id, 'pagesizes')).toBe(true);
    expect(await store.hasDerived(id, 'gridindex')).toBe(true);
  });

  test('deleteDocument removes pagesizes and gridindex derived files along with original/meta', async () => {
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
    await store.saveDerived(id, 'pagesizes', Buffer.from('[]', 'utf8'));
    await store.saveDerived(id, 'gridindex', Buffer.from('{"lineOffsets":[]}', 'utf8'));
    await store.deleteDocument(id);
    expect(await store.readMeta(id)).toBeNull();
    expect(await store.hasDerived(id, 'pagesizes')).toBe(false);
    expect(await store.hasDerived(id, 'gridindex')).toBe(false);
  });
});
