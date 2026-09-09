import { randomUUID } from 'node:crypto';
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { gridCellKey } from '../src/collab/gridShared.js';
import { findOrAddGridStyleIndex } from '../src/xlsx/collab.js';
import { exportXlsxFromGridStreaming } from '../src/xlsx/exportStreaming.js';
import { growSheet, ingestXlsxLazily, setDirtyCell } from '../src/xlsx/lazyGridStore.js';

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('B1').value = 'Score';
  sheet.getCell('A2').value = 'Alice';
  sheet.getCell('B2').value = 10;
  sheet.getCell('A3').value = 'Bob';
  sheet.getCell('B3').value = 20;
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

describe('xlsx/exportStreaming.ts: exportXlsxFromGridStreaming (large-file fallback)', () => {
  test('rejects when structuralOps is non-empty (unsupported scope, fail loud instead of corrupting)', async () => {
    const original = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), original);
    await expect(
      exportXlsxFromGridStreaming(original, store, [{ sheet: 0, kind: 'insertRows', at: 0, count: 1 }]),
    ).rejects.toThrow(/구조 변경 없이|structuralOps|삽입.*삭제/);
  });

  test('an unedited file round-trips with all original values and styles intact', async () => {
    const original = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), original);

    const exported = await exportXlsxFromGridStreaming(original, store, []);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);
    const sheet = reopened.getWorksheet('Sheet1')!;

    expect(sheet.getCell('A1').value).toBe('Name');
    expect(sheet.getCell('B1').value).toBe('Score');
    expect(sheet.getCell('A2').value).toBe('Alice');
    expect(sheet.getCell('B2').value).toBe(10);
    expect(sheet.getCell('A3').value).toBe('Bob');
    expect(sheet.getCell('B3').value).toBe(20);
  });

  test('only cells marked dirty are overwritten; untouched cells keep their original value', async () => {
    const original = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), original);
    setDirtyCell(store, 0, 1, 0, { v: 'Alice Kim' }); // A2만 편집됨

    const exported = await exportXlsxFromGridStreaming(original, store, []);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);
    const sheet = reopened.getWorksheet('Sheet1')!;

    expect(sheet.getCell('A2').value).toBe('Alice Kim');
    expect(sheet.getCell('A1').value).toBe('Name'); // 손대지 않은 셀은 그대로
    expect(sheet.getCell('B2').value).toBe(10);
  });

  test('applies style for a dirty cell using the canonical style palette', async () => {
    const original = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), original);
    const styleIndex = findOrAddGridStyleIndex(store, { b: 1, col: '#ff0000' });
    setDirtyCell(store, 0, 0, 0, { v: 'Name', s: styleIndex }); // A1

    const exported = await exportXlsxFromGridStreaming(original, store, []);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);
    const a1 = reopened.getWorksheet('Sheet1')!.getCell('A1');

    expect(a1.font?.bold).toBe(true);
    expect(a1.font?.color?.argb).toBe('FFFF0000');
    expect(reopened.getWorksheet('Sheet1')!.getCell('B1').font?.bold).toBeFalsy();
  });

  test('a sheet added during collab (not present in the original file) is written in full', async () => {
    const original = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), original);

    // room.ts의 ensureLazySheetSlot이 새 시트를 만들 때 하는 것과 동일한 자리 맞추기.
    store.sheetMeta.push({
      name: 'Sheet2',
      rowCount: 1,
      colCount: 1,
      colWidths: [],
      merges: [],
      freezeRows: 0,
      freezeCols: 0,
    });
    store.rowIndex.push([]);
    store.logicalToOriginalRow.push([]);
    store.logicalToOriginalCol.push([]);
    growSheet(store, 1, 1, 1);
    setDirtyCell(store, 1, 0, 0, { v: 'brand-new' });

    const exported = await exportXlsxFromGridStreaming(original, store, []);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);

    expect(reopened.getWorksheet('Sheet2')!.getCell('A1').value).toBe('brand-new');
    expect(reopened.getWorksheet('Sheet1')!.getCell('A1').value).toBe('Name');
  });

  test('a direct rowCount shrink (no structural op) truncates the sheet in the output', async () => {
    const original = await makeSampleXlsx();
    const store = await ingestXlsxLazily(randomUUID(), original);
    store.sheetMeta[0].rowCount = 2; // 3행짜리를 2행으로 직접 축소(그리드-에디터 툴바 버튼과 같은 모양)

    const exported = await exportXlsxFromGridStreaming(original, store, []);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);
    const reopenedSheet = reopened.getWorksheet('Sheet1')!;

    expect(reopenedSheet.getCell('A2').value).toBe('Alice');
    expect(reopenedSheet.getCell('A3').value).toBeFalsy(); // 잘려나감
  });

  test('sheetDisplayOrder reorders sheets in the streamed output', async () => {
    const workbookIn = new ExcelJS.Workbook();
    workbookIn.addWorksheet('First').getCell('A1').value = 'first';
    workbookIn.addWorksheet('Second').getCell('A1').value = 'second';
    const original = Buffer.from((await workbookIn.xlsx.writeBuffer()) as ArrayBuffer);
    const store = await ingestXlsxLazily(randomUUID(), original);

    const exported = await exportXlsxFromGridStreaming(original, store, [], {
      sheetDisplayOrder: new Map([
        [0, 1],
        [1, 0],
      ]),
    });
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);

    expect(reopened.worksheets.map((s) => s.name)).toEqual(['Second', 'First']);
  });

  test('a removed sheet is excluded from the streamed output', async () => {
    const workbookIn = new ExcelJS.Workbook();
    workbookIn.addWorksheet('First').getCell('A1').value = 'first';
    workbookIn.addWorksheet('Second').getCell('A1').value = 'second';
    const original = Buffer.from((await workbookIn.xlsx.writeBuffer()) as ArrayBuffer);
    const store = await ingestXlsxLazily(randomUUID(), original);

    const exported = await exportXlsxFromGridStreaming(original, store, [], {
      removedSheetIndices: new Set([0]),
    });
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);

    expect(reopened.worksheets.map((s) => s.name)).toEqual(['Second']);
  });
});
