import { randomUUID } from 'node:crypto';
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { ingestXlsxLazily } from '../src/xlsx/lazyGridStore.js';

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
