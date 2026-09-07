import ExcelJS from 'exceljs';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { GRID_CELLS_MAP, gridCellKey } from '../src/collab/gridShared.js';
import { parseXlsxToGrid } from '../src/spreadsheet.js';
import { exportXlsxFromGrid, seedGridFromXlsx } from '../src/xlsx/collab.js';
import type { GridCell, GridDocument, GridStyle } from '../src/spreadsheet.js';

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('B1').value = 'Score';
  sheet.getCell('A2').value = 'Alice';
  sheet.getCell('B2').value = 10;
  sheet.getCell('A3').value = 'Bob';
  sheet.getCell('B3').value = 20;
  sheet.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 30 };
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

describe('collab: xlsx <-> Yjs grid bridge', () => {
  test('a formula cell with no cached result seeds as the raw formula text, not blank', async () => {
    // exportXlsxFromGrid이 저장하는 수식은 { formula } 만 있고 result가 없다(계산은 우리가
    // 안 하니까) — 그런 파일을 다시 seed할 때 빈 문자열이 아니라 "=수식" 텍스트가 나와야
    // 동시편집기의 실시간 계산 엔진이 다시 계산해 보여줄 수 있다.
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 10;
    sheet.getCell('A2').value = 20;
    sheet.getCell('B1').value = { formula: 'SUM(A1:A2)' }; // result 없음 — 우리 export가 만드는 모양 그대로
    const buffer = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);

    const grid = await parseXlsxToGrid(buffer);
    const b1 = grid.sheets[0].rows[0][1];
    expect(b1).not.toBe(0);
    expect((b1 as { v: string }).v).toBe('=SUM(A1:A2)');

    const ydoc = new Y.Doc();
    await seedGridFromXlsx(ydoc, grid);
    expect(ydoc.getMap<string>(GRID_CELLS_MAP).get(gridCellKey(0, 0, 1))).toBe('=SUM(A1:A2)');
  });

  test('seeds a canonical grid from xlsx, edits a cell, and flushes back to a valid xlsx while preserving formulas', async () => {
    const original = await makeSampleXlsx();
    const baselineGrid = await parseXlsxToGrid(original);
    const canonicalGrid = structuredClone(baselineGrid) as GridDocument;

    // "다른 브라우저 탭"이 편집한 것처럼 canonical을 직접 바꾼다(room.ts의
    // applyCellsChangeToCanonical이 하는 것과 동일한 형태의 직접 대입).
    (canonicalGrid.sheets[0].rows[1] as Array<GridCell | 0>)[0] = { v: 'Alice Kim' }; // A2

    const exported = await exportXlsxFromGrid(original, canonicalGrid, baselineGrid);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);
    const sheet = reopened.getWorksheet('Sheet1')!;

    expect(sheet.getCell('A2').value).toBe('Alice Kim');
    // 손대지 않은 숫자 셀은 그대로
    expect(sheet.getCell('B2').value).toBe(10);
    // 손대지 않은 수식 셀은 정적 값으로 깨지지 않고 수식 그대로 보존된다
    const formulaCell = sheet.getCell('B4').value as { formula?: string } | number | null;
    expect(formulaCell).not.toBeNull();
    expect(typeof formulaCell === 'object' && formulaCell !== null ? formulaCell.formula : undefined).toBe(
      'SUM(B2:B3)',
    );
  });

  test('two independent cell edits in different cells both survive (CRDT merge, no last-writer-wins loss)', async () => {
    const original = await makeSampleXlsx();
    const baselineGrid = await parseXlsxToGrid(original);

    // CRDT 병합 자체는 여전히 Yjs 레벨(Task 6 이전까지는 전체 시딩)에서 일어난다 — 두 "브라우저
    // 탭"이 각각 독립된 Y.Doc에서 시작해 서로 다른 셀을 편집하고, 그 업데이트들을 하나의 Y.Doc에
    // 합친다. 이 부분은 예전 테스트와 동일하게 그대로 둔다.
    const ydoc = new Y.Doc();
    await seedGridFromXlsx(ydoc, baselineGrid);

    const clientA = new Y.Doc();
    Y.applyUpdate(clientA, Y.encodeStateAsUpdate(ydoc));
    const clientB = new Y.Doc();
    Y.applyUpdate(clientB, Y.encodeStateAsUpdate(ydoc));

    clientA.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 1, 0), 'Edited by A'); // A2
    clientB.getMap<string>(GRID_CELLS_MAP).set(gridCellKey(0, 2, 0), 'Edited by B'); // A3

    Y.applyUpdate(ydoc, Y.encodeStateAsUpdate(clientA, Y.encodeStateVector(ydoc)));
    Y.applyUpdate(ydoc, Y.encodeStateAsUpdate(clientB, Y.encodeStateVector(ydoc)));

    const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
    expect(cells.get(gridCellKey(0, 1, 0))).toBe('Edited by A');
    expect(cells.get(gridCellKey(0, 2, 0))).toBe('Edited by B');

    // canonical은 그 병합 결과가 반영된 "단일 진실"이다 — room.ts의 cells.observe(...)가 실제로
    // 하는 일(병합이 끝난 뒤 Yjs 상태를 canonical 그리드에 반영)을 여기서 직접 흉내낸다.
    const canonicalGrid = structuredClone(baselineGrid) as GridDocument;
    (canonicalGrid.sheets[0].rows[1] as Array<GridCell | 0>)[0] = { v: cells.get(gridCellKey(0, 1, 0))! }; // A2
    (canonicalGrid.sheets[0].rows[2] as Array<GridCell | 0>)[0] = { v: cells.get(gridCellKey(0, 2, 0))! }; // A3

    const exported = await exportXlsxFromGrid(original, canonicalGrid, baselineGrid);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);
    const sheet = reopened.getWorksheet('Sheet1')!;
    expect(sheet.getCell('A2').value).toBe('Edited by A');
    expect(sheet.getCell('A3').value).toBe('Edited by B');
  });

  test('typing a cell style applies font/fill/alignment on export without touching other cells', async () => {
    const original = await makeSampleXlsx();
    const baselineGrid = await parseXlsxToGrid(original);
    const canonicalGrid = structuredClone(baselineGrid) as GridDocument;

    // 새 스타일을 공용 styles 배열에 추가하고, A1 셀이 그 인덱스를 가리키도록 한다 — canonical
    // 그리드의 셀은 통짜 서식 객체가 아니라 styles 배열의 인덱스(.s)만 들고 있다.
    const styles = canonicalGrid.styles as GridStyle[];
    const styleIndex = styles.length;
    styles.push({ b: 1, col: '#ff0000', bg: '#ffff00', a: 'c' });
    (canonicalGrid.sheets[0].rows[0] as Array<GridCell | 0>)[0] = { v: 'Name', s: styleIndex }; // A1

    const exported = await exportXlsxFromGrid(original, canonicalGrid, baselineGrid);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);
    const sheet = reopened.getWorksheet('Sheet1')!;

    const a1 = sheet.getCell('A1');
    expect(a1.font?.bold).toBe(true);
    expect(a1.font?.color?.argb).toBe('FFFF0000');
    expect((a1.fill as ExcelJS.FillPattern).fgColor?.argb).toBe('FFFFFF00');
    expect(a1.alignment?.horizontal).toBe('center');
    // 손대지 않은 셀은 서식이 그대로(볼드 아님)
    expect(sheet.getCell('B1').font?.bold).toBeFalsy();
  });

  test('a cell value starting with "=" is written back as a real xlsx formula', async () => {
    const original = await makeSampleXlsx();
    const baselineGrid = await parseXlsxToGrid(original);
    const canonicalGrid = structuredClone(baselineGrid) as GridDocument;
    (canonicalGrid.sheets[0].rows[0] as Array<GridCell | 0>)[0] = { v: '=B2+B3' }; // A1

    const exported = await exportXlsxFromGrid(original, canonicalGrid, baselineGrid);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);
    const sheet = reopened.getWorksheet('Sheet1')!;
    const value = sheet.getCell('A1').value as { formula?: string } | null;
    expect(value).not.toBeNull();
    expect(typeof value === 'object' && value !== null ? value.formula : undefined).toBe('B2+B3');
  });

  test('row/column shrink at the end truncates the sheet, clearing edits will drop cleared cells', async () => {
    const original = await makeSampleXlsx();
    const baselineGrid = await parseXlsxToGrid(original);
    const canonicalGrid = structuredClone(baselineGrid) as GridDocument;

    const sheet = canonicalGrid.sheets[0] as unknown as {
      rowCount: number;
      rows: Array<ReadonlyArray<GridCell | 0>>;
    };
    const beforeRowCount = sheet.rowCount;
    sheet.rowCount = beforeRowCount - 1;
    sheet.rows = sheet.rows.slice(0, beforeRowCount - 1); // 마지막(수식) 행 삭제

    const exported = await exportXlsxFromGrid(original, canonicalGrid, baselineGrid);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(exported as unknown as ArrayBuffer);
    const reopenedSheet = reopened.getWorksheet('Sheet1')!;
    expect(reopenedSheet.rowCount).toBe(beforeRowCount - 1);
  });
});
