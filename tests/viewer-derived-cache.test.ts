import { mkdtempSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, test, vi } from 'vitest';

async function loadViewerModuleWithTempDir() {
  const dir = mkdtempSync(path.join(os.tmpdir(), 'docview-viewer-test-'));
  vi.resetModules();
  process.env.DOCVIEW_DATA_DIR = dir;
  process.env.DOCVIEW_STORAGE = 'disk';
  const storageModule = await import('../src/storage.js');
  await storageModule.ensureDataDirs();
  const viewerModule = await import('../src/routes/viewer.js');
  return { dir, storageModule, viewerModule };
}

function sampleGridHeader() {
  return {
    styles: [],
    sheets: [
      { name: 'S1', rowCount: 2, colCount: 1, colWidths: [], merges: [], freezeRows: 0, freezeCols: 0 },
    ],
  };
}

describe('viewer.ts: loadGrid self-heals and reuses a persisted gridindex', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('computes and persists gridindex when missing, matching a manual scan', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    const jsonl = [
      JSON.stringify(sampleGridHeader()),
      JSON.stringify([{ v: 'a' }]),
      JSON.stringify([{ v: 'b' }]),
    ].join('\n');
    await store.saveDerived(id, 'grid', Buffer.from(jsonl, 'utf8'));

    expect(await store.hasDerived(id, 'gridindex')).toBe(false);
    const grid = await viewerModule.loadGrid(id);
    expect(grid.header.sheets[0].name).toBe('S1');
    expect(await store.hasDerived(id, 'gridindex')).toBe(true);
    const persisted = JSON.parse((await store.readDerived(id, 'gridindex')).toString('utf8'));
    expect(persisted.lineOffsets).toEqual(grid.lineOffsets);
  });

  test('uses a persisted gridindex as-is instead of rescanning the buffer', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    const jsonl = [
      JSON.stringify(sampleGridHeader()),
      JSON.stringify([{ v: 'a' }]),
      JSON.stringify([{ v: 'b' }]),
    ].join('\n');
    await store.saveDerived(id, 'grid', Buffer.from(jsonl, 'utf8'));
    // 실제 스캔이라면 절대 나올 수 없는 표시용 오프셋(마지막 값 999) — loadGrid가 이 값을
    // 그대로 쓴다면(다시 스캔하지 않는다면) 아래 assertion에서 바로 드러난다. 첫 두 오프셋
    // (0, 122)은 이 fixture의 실제 헤더 줄 경계와 일치시켜, buildLoadedGridFromOffsets의
    // 헤더 파싱(JSON.parse)이 이 표시용 오프셋과 무관하게 정상 동작하도록 한다.
    await store.saveDerived(
      id,
      'gridindex',
      Buffer.from(JSON.stringify({ lineOffsets: [0, 122, 999] }), 'utf8'),
    );

    const grid = await viewerModule.loadGrid(id);
    expect(grid.lineOffsets).toEqual([0, 122, 999]);
  });
});

describe('viewer.ts: pageSizePersistence', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('load() returns null when no pagesizes derived data exists', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { newDocumentId } = storageModule;
    const id = newDocumentId();
    expect(await viewerModule.pageSizePersistence(id).load()).toBeNull();
  });

  test('save() then load() round-trips the page sizes', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { newDocumentId } = storageModule;
    const id = newDocumentId();
    const persistence = viewerModule.pageSizePersistence(id);
    await persistence.save([{ width: 595, height: 842 }]);
    expect(await persistence.load()).toEqual([{ width: 595, height: 842 }]);
  });
});
