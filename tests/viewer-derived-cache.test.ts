import { mkdtempSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { afterEach, describe, expect, test, vi } from 'vitest';
import { GRID_CACHE_FORMAT_VERSION } from '../src/spreadsheet.js';

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

// formatVersion 필드를 일부러 넣지 않는다 — 이 describe 블록의 테스트들은 gridindex 자가치유만
// 검증한다. store.writeMeta를 호출하지 않으므로(각 테스트가 store.readMeta(id)로 얻을 meta가
// 없다) regenerateStaleGridCache는 항상 즉시 false로 끝나 아래 loadGrid 호출은 기존 gridindex
// 경로 그대로 동작한다 — 두 관심사(캐시 버전 자가치유는 아래 별도 describe)가 서로 간섭하지 않는다.
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
    // 실제 스캔이라면 절대 나올 수 없는(하지만 여전히 버퍼 길이 이내라 offsetsLookValid는
    // 통과하는) 표시용 마지막 오프셋 — loadGrid가 이 값을 그대로 쓴다면(다시 스캔하지
    // 않는다면) 아래 assertion에서 바로 드러난다. 첫 두 오프셋(0, headerEnd)은 이 fixture의
    // 실제 헤더 줄 경계와 일치시켜(하드코딩 대신 실제로 계산해), buildLoadedGridFromOffsets의
    // 헤더 파싱(JSON.parse)이 정상 동작하도록 한다.
    const headerEnd = jsonl.indexOf('\n') + 1;
    await store.saveDerived(
      id,
      'gridindex',
      Buffer.from(JSON.stringify({ lineOffsets: [0, headerEnd, 140] }), 'utf8'),
    );

    const grid = await viewerModule.loadGrid(id);
    expect(grid.lineOffsets).toEqual([0, headerEnd, 140]);
  });

  test('self-heals when the persisted gridindex is corrupt (not valid JSON)', async () => {
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
    await store.saveDerived(id, 'gridindex', Buffer.from('{not valid json', 'utf8'));

    const grid = await viewerModule.loadGrid(id);
    expect(grid.header.sheets[0].name).toBe('S1');
    // 손상된 파일이 실제로 다시 스캔한 올바른 값으로 덮어써졌는지(자가치유) 확인한다.
    const persisted = JSON.parse((await store.readDerived(id, 'gridindex')).toString('utf8'));
    expect(persisted.lineOffsets).toEqual(grid.lineOffsets);
  });

  test('self-heals when the persisted gridindex has an out-of-bounds offset (the exact clamp-bug class this guards against)', async () => {
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
    // 마지막 오프셋이 버퍼 길이(145바이트)보다 큼 — 그대로 신뢰했다면 Buffer.toString이 end를
    // 조용히 clamp해 여러 줄이 뒤섞인 문자열을 반환하고 JSON.parse가 깨졌을 상황이다.
    await store.saveDerived(
      id,
      'gridindex',
      Buffer.from(JSON.stringify({ lineOffsets: [0, 122, 9999] }), 'utf8'),
    );

    const grid = await viewerModule.loadGrid(id);
    expect(grid.lineOffsets).toEqual([0, 122, 134]); // 실제로 다시 스캔한 값
  });
});

/** LoadedGrid(grid.buffer + lineOffsets + sheetStartLine)에서 셀 하나의 표시 텍스트를 직접
 * 읽는다 — viewer.ts의 sliceRows와 같은 인덱싱 규칙이지만, 테스트에서만 필요해 여기 따로 둔다. */
function cellTextAt(
  grid: { buffer: Buffer; lineOffsets: readonly number[]; sheetStartLine: readonly number[] },
  sheetIdx: number,
  row: number,
  col: number,
): string {
  const line = grid.sheetStartLine[sheetIdx] + row;
  const start = grid.lineOffsets[line];
  const end = line + 1 < grid.lineOffsets.length ? grid.lineOffsets[line + 1] - 1 : grid.buffer.length;
  const rowCells = JSON.parse(grid.buffer.toString('utf8', start, end)) as Array<{ v: string } | 0>;
  const cell = rowCells[col];
  return cell === 0 || cell === undefined ? '' : cell.v;
}

async function makeCrossSheetXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet1 = workbook.addWorksheet('Sheet1');
  sheet1.getCell('A1').value = 2287;
  const sheet2 = workbook.addWorksheet('Sheet2');
  sheet2.getCell('A1').value = { formula: 'SUM(Sheet1!A1:A1)' };
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

/** xlsx/formulas.ts의 시트 간 참조 버그가 고쳐지기 전에 만들어졌을 법한 grid.json을 그대로
 * 흉내낸다 — formatVersion 필드 자체가 없고(옛 serializeGrid는 이 필드를 쓰지 않았다),
 * Sheet2!A1(SUM(Sheet1!A1:A1))이 실제로는 '0'(브리프가 재현한 그 버그의 산출물)으로 잘못
 * 캐시돼 있다. */
function staleCrossSheetGridJsonl(): Buffer {
  const header = {
    styles: [],
    sheets: [
      { name: 'Sheet1', rowCount: 1, colCount: 1, colWidths: [], merges: [], freezeRows: 0, freezeCols: 0 },
      { name: 'Sheet2', rowCount: 1, colCount: 1, colWidths: [], merges: [], freezeRows: 0, freezeCols: 0 },
    ],
  };
  const parts = [
    JSON.stringify(header),
    JSON.stringify([{ v: '2287' }]),
    JSON.stringify([{ v: '0' }]), // 버그 당시의 (틀린) 캐시된 결과
  ];
  return Buffer.from(parts.join('\n'), 'utf8');
}

describe('viewer.ts: loadGrid self-heals a stale-format grid.json cache (GRID_CACHE_FORMAT_VERSION)', () => {
  let tempDirs: string[] = [];

  afterEach(() => {
    delete process.env.DOCVIEW_DATA_DIR;
    delete process.env.DOCVIEW_STORAGE;
    vi.resetModules();
    for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true });
    tempDirs = [];
  });

  test('a grid.json with no formatVersion field (pre-fix cache) is regenerated from the original ' +
    'file on next view, replacing a wrong cross-sheet-formula value with the correct one', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();

    await store.saveOriginal(id, 'xlsx', await makeCrossSheetXlsx());
    await store.writeMeta({
      id,
      name: 'x.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 2,
      createdAt: new Date().toISOString(),
      progress: 100,
      viewKind: 'grid',
    });
    // 업로드 시점에 만들어진 뒤로 한 번도 재생성되지 않은, 실제 버그의 산출물을 흉내낸 낡은 캐시.
    await store.saveDerived(id, 'grid', staleCrossSheetGridJsonl());

    const grid = await viewerModule.loadGrid(id);

    expect(grid.header.formatVersion).toBe(GRID_CACHE_FORMAT_VERSION);
    expect(cellTextAt(grid, 1, 0, 0)).toBe('2287'); // 더 이상 '0'이 아니다.

    // 재생성 결과가 디스크에도 실제로 저장돼, 다음 조회부터는 다시 계산할 필요가 없어야 한다.
    const persisted = (await store.readDerived(id, 'grid')).toString('utf8');
    expect(persisted).toContain('"2287"');
    expect(JSON.parse(persisted.split('\n')[0]).formatVersion).toBe(GRID_CACHE_FORMAT_VERSION);
  });

  test('a stale-format grid.json falls back to serving the existing cache as-is (no crash) when ' +
    'the original file is unavailable for regeneration', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();

    // 의도적으로 원본을 저장하지 않는다(예: 오래돼 정리된 원본, 혹은 메모리 모드에서 만료됨을
    // 흉내). 자가치유가 실패해도 조회 자체는 절대 깨지면 안 된다.
    await store.writeMeta({
      id,
      name: 'x.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 2,
      createdAt: new Date().toISOString(),
      progress: 100,
      viewKind: 'grid',
    });
    await store.saveDerived(id, 'grid', staleCrossSheetGridJsonl());

    const grid = await viewerModule.loadGrid(id);

    // 재생성이 실패했으니 낡은(옳지 않을 수 있는) 캐시를 그대로 서빙한다 — 크래시하지 않는 것이 핵심.
    expect(grid.header.formatVersion).toBeUndefined();
    expect(cellTextAt(grid, 1, 0, 0)).toBe('0');
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

  test('self-heals (returns null) when the persisted pagesizes file is corrupt (not valid JSON)', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.saveDerived(id, 'pagesizes', Buffer.from('{not valid json', 'utf8'));
    expect(await viewerModule.pageSizePersistence(id).load()).toBeNull();
  });

  test('self-heals (returns null) when the persisted pagesizes file is valid JSON but not an array', async () => {
    const { dir, storageModule, viewerModule } = await loadViewerModuleWithTempDir();
    tempDirs.push(dir);
    const { store, newDocumentId } = storageModule;
    const id = newDocumentId();
    await store.saveDerived(id, 'pagesizes', Buffer.from('{"oops":true}', 'utf8'));
    expect(await viewerModule.pageSizePersistence(id).load()).toBeNull();
  });
});
