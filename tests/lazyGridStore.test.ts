import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { gridCellKey } from '../src/collab/gridShared.js';
import type { GridCell } from '../src/spreadsheet.js';
import {
  applyColShift,
  applyRowShift,
  closeLazyGridStore,
  getCell,
  getRow,
  growSheet,
  ingestXlsxLazily,
  setDirtyCell,
  shiftDirtyOverlayKeys,
  sweepOrphanedLazyGridFiles,
} from '../src/xlsx/lazyGridStore.js';

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('B1').value = 'Score';
  sheet.getCell('A2').value = 'Alice';
  sheet.getCell('B2').value = 10;
  sheet.getRow(1).getCell(1).font = { bold: true };
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('xlsx/lazyGridStore.ts: ingestXlsxLazily', () => {
  test('produces correct sheet metadata and an on-disk row store for a small file', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);

    expect(store.sheetMeta).toHaveLength(1);
    expect(store.sheetMeta[0].name).toBe('Sheet1');
    expect(store.sheetMeta[0].rowCount).toBe(2);
    expect(store.sheetMeta[0].colCount).toBe(2);
    expect(store.styles.length).toBeGreaterThanOrEqual(1); // A1의 bold 서식

    // 항등 매핑으로 초기화돼야 한다
    expect(store.logicalToOriginalRow[0]).toEqual([0, 1]);
    expect(store.logicalToOriginalCol[0]).toEqual([0, 1]);

    // 온디스크 저장소에 실제로 두 행이 기록됐어야 한다
    expect(store.rowIndex[0][0]).toBeDefined();
    expect(store.rowIndex[0][1]).toBeDefined();

    await import('node:fs/promises').then((fs) => fs.rm(store.rowStorePath, { force: true }));
  });

  test('trims trailing formatting-only rows the same way parseXlsxToGrid does', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 'hi';
    sheet.getRow(10).height = 30; // 값 없이 서식(행 높이)만 — 꼬리 행
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const store = await ingestXlsxLazily(randomUUID(), bytes);
    // parseXlsxToGrid와 동일하게: 실제 값이 있는 마지막 행(1)까지만 rowCount로 잡는다.
    expect(store.sheetMeta[0].rowCount).toBe(1);

    await import('node:fs/promises').then((fs) => fs.rm(store.rowStorePath, { force: true }));
  });

  test('resolves theme colors the same way as the non-streaming parser', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 'x';
    sheet.getCell('A1').font = { color: { theme: 1 } }; // dk1 테마 색
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const { parseXlsxToGrid } = await import('../src/spreadsheet.js');
    const viaFullParse = await parseXlsxToGrid(bytes);
    const cellViaFullParse = viaFullParse.sheets[0].rows[0][0];
    const styleViaFullParse =
      cellViaFullParse !== 0 && cellViaFullParse.s !== undefined
        ? viaFullParse.styles[cellViaFullParse.s]
        : undefined;

    const store = await ingestXlsxLazily(randomUUID(), bytes);
    const fs = await import('node:fs/promises');
    const raw = await fs.readFile(store.rowStorePath);
    const entry = store.rowIndex[0][0]!;
    const row = JSON.parse(raw.subarray(entry.byteOffset, entry.byteOffset + entry.byteLength).toString('utf8'));
    const cellViaStreaming = row[0];
    const styleViaStreaming =
      cellViaStreaming !== 0 && cellViaStreaming.s !== undefined ? store.styles[cellViaStreaming.s] : undefined;

    expect(styleViaStreaming).toEqual(styleViaFullParse);
    await fs.rm(store.rowStorePath, { force: true });
  });
});

describe('xlsx/lazyGridStore.ts: getCell/getRow', () => {
  test('reads an unedited cell from disk on first access, then from cache on repeat access', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);

    expect(await getCell(store, 0, 0, 0)).toEqual({ v: 'Name', s: expect.any(Number) });
    expect(await getCell(store, 0, 1, 0)).toEqual({ v: 'Alice' });
    // 캐시가 채워졌는지 — 같은 원본 행에 대해 readCache에 항목이 생겼어야 한다.
    expect(store.readCache.has('0:0')).toBe(true);

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('a dirty overlay entry always wins over the on-disk original, even if cached', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    await getCell(store, 0, 1, 0); // 먼저 캐시를 데운다
    store.dirtyOverlay.set(gridCellKey(0, 1, 0), { v: 'Alice Kim' });

    expect(await getCell(store, 0, 1, 0)).toEqual({ v: 'Alice Kim' });

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('a coordinate beyond logicalToOriginalRow/Col bounds (newly inserted, unedited) returns 0 without touching disk', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    store.logicalToOriginalRow[0].splice(1, 0, -1); // 1번 위치에 "새 행" 삽입

    expect(await getCell(store, 0, 1, 0)).toBe(0); // 새로 삽입된 빈 행
    expect(await getCell(store, 0, 2, 0)).toEqual({ v: 'Alice' }); // 밀려난 원래 2번째 행

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('readCache evicts the oldest row once over the configured limit, and re-reads correctly after eviction', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    store.readCacheLimit = 1;

    await getRow(store, 0, 0);
    expect(store.readCache.has('0:0')).toBe(true);
    await getRow(store, 0, 1);
    expect(store.readCache.has('0:0')).toBe(false); // 밀려남
    expect(store.readCache.has('0:1')).toBe(true);

    // 밀려난 뒤에도 다시 읽으면 같은 값
    expect(await getCell(store, 0, 0, 0)).toEqual({ v: 'Name', s: expect.any(Number) });

    await fs.rm(store.rowStorePath, { force: true });
  });
});

describe('xlsx/lazyGridStore.ts: setDirtyCell/shiftDirtyOverlayKeys', () => {
  test('setDirtyCell makes getCell return the new value immediately', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    setDirtyCell(store, 0, 1, 0, { v: 'Alice Kim' });
    expect(await getCell(store, 0, 1, 0)).toEqual({ v: 'Alice Kim' });
    await fs.rm(store.rowStorePath, { force: true });
  });

  test('shiftDirtyOverlayKeys moves an edited coordinate down when a row is inserted above it', () => {
    const overlay = new Map<string, GridCell | 0>([[gridCellKey(0, 5, 2), { v: 'edited' }]]);
    shiftDirtyOverlayKeys(overlay, 0, 'row', 2, 1, true); // 2번 위치에 행 1개 삽입
    expect(overlay.has(gridCellKey(0, 5, 2))).toBe(false);
    expect(overlay.get(gridCellKey(0, 6, 2))).toEqual({ v: 'edited' });
  });

  test('shiftDirtyOverlayKeys drops an edited coordinate that falls inside a deleted range', () => {
    const overlay = new Map<string, GridCell | 0>([[gridCellKey(0, 5, 2), { v: 'edited' }]]);
    shiftDirtyOverlayKeys(overlay, 0, 'row', 4, 3, false); // 4,5,6번 행 삭제(5가 그 안에 포함)
    expect(overlay.has(gridCellKey(0, 5, 2))).toBe(false);
    expect(overlay.size).toBe(0);
  });

  test('shiftDirtyOverlayKeys leaves coordinates in other sheets untouched', () => {
    const overlay = new Map<string, GridCell | 0>([[gridCellKey(1, 5, 2), { v: 'other sheet' }]]);
    shiftDirtyOverlayKeys(overlay, 0, 'row', 0, 1, true);
    expect(overlay.get(gridCellKey(1, 5, 2))).toEqual({ v: 'other sheet' });
  });
});

describe('xlsx/lazyGridStore.ts: applyRowShift/applyColShift/growSheet', () => {
  test('inserting a row shifts logicalToOriginalRow and dirtyOverlay together, keeping getCell correct', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    setDirtyCell(store, 0, 1, 0, { v: 'Alice Kim' }); // 논리 1행(원래 Alice)을 편집

    applyRowShift(store, 0, 1, 1, true); // 1번 위치에 행 1개 삽입 — Alice가 2번째로 밀림

    expect(store.logicalToOriginalRow[0]).toEqual([0, -1, 1]);
    expect(await getCell(store, 0, 1, 0)).toBe(0); // 새로 삽입된 빈 행
    expect(await getCell(store, 0, 2, 0)).toEqual({ v: 'Alice Kim' }); // 편집 내용이 함께 밀림

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('deleting a row removes the mapping and any dirty overlay entry inside the deleted range', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    applyRowShift(store, 0, 0, 1, false); // 0번째 행(Name/Score 헤더) 삭제

    expect(store.logicalToOriginalRow[0]).toEqual([1]);
    expect(await getCell(store, 0, 0, 0)).toEqual({ v: 'Alice' });

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('applyColShift mirrors applyRowShift for the column axis', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    applyColShift(store, 0, 0, 1, true); // 맨 앞에 열 1개 삽입 — Name/Alice가 오른쪽으로 밀림

    expect(store.logicalToOriginalCol[0]).toEqual([-1, 0, 1]);
    expect(await getCell(store, 0, 0, 0)).toBe(0);
    expect(await getCell(store, 0, 0, 1)).toEqual({ v: 'Name', s: expect.any(Number) });

    await fs.rm(store.rowStorePath, { force: true });
  });

  test('growSheet extends logicalToOriginalRow/Col with -1 entries without touching existing mappings', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    growSheet(store, 0, 5, 4); // 2행×2열 시트를 5행×4열로 직접 확장(삽입 연산 없이)

    expect(store.logicalToOriginalRow[0]).toEqual([0, 1, -1, -1, -1]);
    expect(store.logicalToOriginalCol[0]).toEqual([0, 1, -1, -1]);
    expect(await getCell(store, 0, 4, 3)).toBe(0); // 새로 확장된 영역

    await fs.rm(store.rowStorePath, { force: true });
  });
});

describe('xlsx/lazyGridStore.ts: cleanup', () => {
  test('closeLazyGridStore deletes the on-disk row store file', async () => {
    const bytes = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    await closeLazyGridStore(store);
    await expect(fs.access(store.rowStorePath)).rejects.toThrow();
  });

  test('sweepOrphanedLazyGridFiles removes leftover docview-lazygrid-* files in the OS temp dir', async () => {
    const os = await import('node:os');
    const path = await import('node:path');
    const orphanPath = path.join(os.tmpdir(), `docview-lazygrid-orphan-${randomUUID()}.jsonl`);
    await fs.writeFile(orphanPath, 'leftover');

    await sweepOrphanedLazyGridFiles();

    await expect(fs.access(orphanPath)).rejects.toThrow();
  });

  test('sweepOrphanedLazyGridFiles does not touch unrelated temp files', async () => {
    const os = await import('node:os');
    const path = await import('node:path');
    const unrelatedPath = path.join(os.tmpdir(), `not-ours-${randomUUID()}.txt`);
    await fs.writeFile(unrelatedPath, 'leave me alone');

    await sweepOrphanedLazyGridFiles();

    await expect(fs.access(unrelatedPath)).resolves.toBeUndefined();
    await fs.rm(unrelatedPath, { force: true });
  });
});
