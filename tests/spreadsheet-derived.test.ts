import ExcelJS from 'exceljs';
import { describe, expect, test, vi } from 'vitest';
import {
  computeLineOffsets,
  displayValue,
  extractStyle,
  parseMergeRange,
  parseXlsxToGrid,
  saveGridDerived,
  StylePalette,
  YIELD_EVERY_N_ROWS,
  type GridDocument,
} from '../src/spreadsheet.js';

describe('spreadsheet: computeLineOffsets', () => {
  test('returns byte offsets for each line, starting with 0', () => {
    const buffer = Buffer.from('line one\nline two\nline three', 'utf8');
    const offsets = computeLineOffsets(buffer);
    expect(offsets).toEqual([0, 9, 18]);
    expect(buffer.toString('utf8', offsets[1], offsets[2] - 1)).toBe('line two');
  });

  test('a buffer with no newline yields just the initial 0 offset', () => {
    const buffer = Buffer.from('single line, no newline', 'utf8');
    expect(computeLineOffsets(buffer)).toEqual([0]);
  });
});

function makeTestGrid(): GridDocument {
  return {
    styles: [],
    sheets: [
      {
        name: 'Sheet1',
        rowCount: 2,
        colCount: 1,
        colWidths: [],
        merges: [],
        freezeRows: 0,
        freezeCols: 0,
        rows: [[{ v: 'a' }], [{ v: 'b' }]],
      },
    ],
  };
}

describe('spreadsheet: saveGridDerived', () => {
  test('saves both grid and gridindex derived kinds, with offsets matching the serialized grid bytes', async () => {
    const saveDerived = vi.fn(async () => undefined);
    await saveGridDerived({ saveDerived }, 'doc-1', makeTestGrid());

    expect(saveDerived).toHaveBeenCalledTimes(2);
    const [gridCall, indexCall] = saveDerived.mock.calls as unknown[][];
    expect(gridCall[0]).toBe('doc-1');
    expect(gridCall[1]).toBe('grid');
    const serializedBytes = gridCall[2] as Buffer;

    expect(indexCall[0]).toBe('doc-1');
    expect(indexCall[1]).toBe('gridindex');
    const { lineOffsets } = JSON.parse((indexCall[2] as Buffer).toString('utf8'));
    expect(lineOffsets).toEqual(computeLineOffsets(serializedBytes));
  });
});

describe('spreadsheet.ts: YIELD_EVERY_N_ROWS is exported', () => {
  test('is a positive integer', () => {
    expect(Number.isInteger(YIELD_EVERY_N_ROWS)).toBe(true);
    expect(YIELD_EVERY_N_ROWS).toBeGreaterThan(0);
  });
});

describe('parseXlsxToGrid: sparse trailing rows', () => {
  test('a sheet with real data far below a naive actualRowCount is not truncated', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    // 헤더 행 하나, 그리고 그 사이 348개 행은 완전히 비워둔 채(어떤 Row 객체도 건드리지 않는다)
    // 350행에만 실데이터를 심는다 — sheet.actualRowCount는 "값이 있는 행의 개수"만 세므로 여기서는
    // 2(1행 + 350행)를 돌려주지만, 진짜 최대 행 번호는 350이다. 이 시나리오가 바로 이 버그를 처음
    // 발견하게 만든 실제 재현 조건이다(로우 350에 실데이터, 그 앞은 거의 비어 있는 성긴 시트).
    sheet.getCell('A1').value = 'header';
    sheet.getCell('B350').value = 'far-down-data';

    const buffer = await workbook.xlsx.writeBuffer();

    // 재현 전제 확인 — actualRowCount가 실제로 성긴 시트를 과소평가하는지부터 검증한다. 이게 깨지면
    // (예: exceljs가 언젠가 이 동작을 바꾸면) 아래 본 검증이 버그를 재현하지 못하는 것이므로 먼저
    // 알아챌 수 있어야 한다.
    const reloaded = new ExcelJS.Workbook();
    await reloaded.xlsx.load(buffer as unknown as ArrayBuffer);
    const reloadedSheet = reloaded.worksheets[0]!;
    expect(reloadedSheet.actualRowCount).toBeLessThan(350);
    expect(reloadedSheet.rowCount).toBe(350);

    const grid = await parseXlsxToGrid(Buffer.from(buffer));
    const gridSheet = grid.sheets[0]!;
    expect(gridSheet.rowCount).toBe(350);
    expect(gridSheet.rows.length).toBe(350);
    const farDownRow = gridSheet.rows[349]!;
    expect(farDownRow[1]).not.toBe(0);
    expect((farDownRow[1] as { v: string }).v).toBe('far-down-data');
  });
});

describe('spreadsheet.ts: newly-exported helpers behave identically to before', () => {
  test('displayValue/extractStyle/parseMergeRange work when called directly (not just via parseXlsxToGrid)', () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 'hello';
    sheet.getCell('A1').font = { bold: true };
    sheet.mergeCells('A1:B1');

    expect(displayValue(sheet.getCell('A1'))).toBe('hello');

    const palette = new StylePalette();
    const styleIndex = extractStyle(sheet.getCell('A1'), palette, [], false);
    expect(styleIndex).toBe(0);
    expect(palette.toArray()[0]).toEqual({ b: 1 });

    expect(parseMergeRange('A1:B1')).toEqual({ r: 1, c: 1, rs: 1, cs: 2 });
  });
});
