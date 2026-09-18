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

describe('storage: MemoryStore memory-budget enforcement (incremental byte counter)', () => {
  afterEach(() => {
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
  });

  test('evicts the oldest document once total bytes exceed the budget, protecting the document just being written', async () => {
    vi.resetModules();
    process.env.DOCVIEW_STORAGE = 'memory';
    const { MemoryStore, newDocumentId } = await import('../src/storage.js');
    const store = new MemoryStore(25); // 25바이트 예산 — 10바이트짜리 문서 2개가 넘는다

    const idA = newDocumentId();
    await store.writeMeta({ id: idA, name: 'a.pdf', ext: 'pdf', status: 'ready', pageCount: 1, createdAt: '2020-01-01T00:00:00.000Z' });
    await store.saveOriginal(idA, 'pdf', Buffer.alloc(10, 'a'));

    const idB = newDocumentId();
    await store.writeMeta({ id: idB, name: 'b.pdf', ext: 'pdf', status: 'ready', pageCount: 1, createdAt: '2020-01-02T00:00:00.000Z' });
    await store.saveOriginal(idB, 'pdf', Buffer.alloc(10, 'b'));

    const idC = newDocumentId();
    await store.writeMeta({ id: idC, name: 'c.pdf', ext: 'pdf', status: 'ready', pageCount: 1, createdAt: '2020-01-03T00:00:00.000Z' });
    await store.saveOriginal(idC, 'pdf', Buffer.alloc(10, 'c'));

    // 세 문서(30바이트)가 예산(25바이트)을 넘으므로 가장 오래된 A가 제거된다.
    expect(await store.readMeta(idA)).toBeNull();
    expect((await store.readOriginal(idB, 'pdf')).length).toBe(10);
    expect((await store.readOriginal(idC, 'pdf')).length).toBe(10);
  });

  test('re-saving the same id with a smaller buffer does not leak the old size into the running byte counter (would otherwise evict too eagerly)', async () => {
    vi.resetModules();
    process.env.DOCVIEW_STORAGE = 'memory';
    const { MemoryStore, newDocumentId } = await import('../src/storage.js');
    const store = new MemoryStore(15);

    const id = newDocumentId();
    await store.writeMeta({ id, name: 'a.pdf', ext: 'pdf', status: 'ready', pageCount: 1, createdAt: '2020-01-01T00:00:00.000Z' });
    await store.saveOriginal(id, 'pdf', Buffer.alloc(10, 'x'));
    // 같은 id를 더 작은 버퍼로 재저장 — 이전 크기를 빼지 않으면 카운터가 부풀어서, 실제로는
    // 예산(15바이트) 안에 드는 상황인데도 불필요하게 이 문서 자신을 내쫓으려 시도하게 된다.
    await store.saveOriginal(id, 'pdf', Buffer.alloc(3, 'y'));

    expect(await store.readMeta(id)).not.toBeNull();
    expect((await store.readOriginal(id, 'pdf')).length).toBe(3);
  });

  test('deleteDocument correctly frees counted bytes so a later save is not evicted by a phantom leftover count', async () => {
    vi.resetModules();
    process.env.DOCVIEW_STORAGE = 'memory';
    const { MemoryStore, newDocumentId } = await import('../src/storage.js');
    const store = new MemoryStore(12);

    const idA = newDocumentId();
    await store.writeMeta({ id: idA, name: 'a.pdf', ext: 'pdf', status: 'ready', pageCount: 1, createdAt: '2020-01-01T00:00:00.000Z' });
    await store.saveOriginal(idA, 'pdf', Buffer.alloc(10, 'a'));
    await store.deleteDocument(idA);

    const idB = newDocumentId();
    await store.writeMeta({ id: idB, name: 'b.pdf', ext: 'pdf', status: 'ready', pageCount: 1, createdAt: '2020-01-02T00:00:00.000Z' });
    await store.saveOriginal(idB, 'pdf', Buffer.alloc(10, 'b'));

    // A를 지운 뒤 바이트가 실제로 빠지지 않았다면(카운터 누수) B를 저장할 때 예산을 초과한 것으로
    // 잘못 판단해 보호 대상인 B 자신을 내쫓으려 하다가(protectedId 예외로 실패) 무한 루프에
    // 빠지거나, 애초에 남아있어야 할 B가 사라진다.
    expect(await store.readMeta(idB)).not.toBeNull();
    expect((await store.readOriginal(idB, 'pdf')).length).toBe(10);
  });
});

describe('storage: MemoryStore derived-kind cleanup is generic (no per-kind hardcoding)', () => {
  afterEach(() => {
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
  });

  test('deleteDocument removes every derived kind for that id via a prefix sweep, including newly-added kinds', async () => {
    vi.resetModules();
    process.env.DOCVIEW_STORAGE = 'memory';
    const { store, newDocumentId } = await import('../src/storage.js');
    const id = newDocumentId();
    await store.writeMeta({
      id,
      name: 'a.pdf',
      ext: 'pdf',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
    });
    const kinds = ['pdf', 'grid', 'markup', 'pagesizes', 'gridindex'] as const;
    for (const kind of kinds) {
      await store.saveDerived(id, kind, Buffer.from(`data-for-${kind}`, 'utf8'));
    }
    await store.deleteDocument(id);
    for (const kind of kinds) {
      expect(await store.hasDerived(id, kind)).toBe(false);
    }
  });
});
