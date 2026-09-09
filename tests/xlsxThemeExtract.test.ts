import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { extractThemeColorsFromZip } from '../src/xlsx/themeExtract.js';

describe('xlsx/themeExtract.ts: extractThemeColorsFromZip', () => {
  test('extracts the same theme colors that ExcelJS.Workbook.xlsx.load would expose via workbook.model.themes', async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet('Sheet1').getCell('A1').value = 'x';
    const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const viaZip = await extractThemeColorsFromZip(bytes);

    const reloaded = new ExcelJS.Workbook();
    await reloaded.xlsx.load(bytes as unknown as ArrayBuffer);
    const themes = (reloaded.model as unknown as { themes?: Record<string, string> }).themes;
    const { parseThemeColors } = await import('../src/spreadsheet.js');
    const viaFullLoad = parseThemeColors(themes?.theme1);

    expect(viaZip).toEqual(viaFullLoad);
    expect(viaZip.length).toBeGreaterThan(0);
  });

  test('returns an empty array for a zip with no theme1.xml entry (never throws)', async () => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();
    zip.file('empty.txt', 'nothing here');
    const bytes = Buffer.from(await zip.generateAsync({ type: 'nodebuffer' }));

    await expect(extractThemeColorsFromZip(bytes)).resolves.toEqual([]);
  });
});
