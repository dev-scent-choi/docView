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

/** 워크시트를 xlsx 바이트로 굽고 다시 parseXlsxToGrid로 읽어, 첫 시트의 표시 텍스트만 꺼낸다.
 * (ExcelJS가 쓰는 수식에는 계산 결과가 캐시되지 않으므로 — displayValue 문서 주석 참고 —
 * 이 경로로 만든 수식 셀은 전부 "캐시된 결과 없음", 즉 서버가 직접 계산해야 하는 칸이다.) */
async function parseSheet(build: (sheet: ExcelJS.Worksheet) => void) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  build(sheet);
  const buffer = await workbook.xlsx.writeBuffer();
  const grid = await parseXlsxToGrid(Buffer.from(buffer));
  const gridSheet = grid.sheets[0]!;
  const textAt = (row: number, col: number): string => {
    const cell = gridSheet.rows[row - 1]?.[col - 1];
    return cell === 0 || cell === undefined ? '' : cell.v;
  };
  return { grid, gridSheet, textAt };
}

describe('parseXlsxToGrid: formulas are computed server-side for the read-only viewer', () => {
  test('a formula with no cached result shows the COMPUTED value, not the raw formula text', async () => {
    const { textAt } = await parseSheet((sheet) => {
      sheet.getCell('A1').value = 2;
      sheet.getCell('A2').value = 40;
      sheet.getCell('A3').value = { formula: 'SUM(A1:A2)', result: undefined };
    });
    expect(textAt(3, 1)).toBe('42');
  });

  test('a chained formula (a formula referencing another formula cell) computes correctly — the ' +
    'non-trivial case the shared per-pass memo has to get right', async () => {
    // A1=1, A2=A1+1, ... A31=A30+1 → 사슬 끝은 31이어야 한다. memo를 패스마다 하나씩만 만드는
    // 구현이라면 각 칸이 정확히 한 번씩만 계산된다(O(N²) 폭증 없음, computeGridCellDisplay 주석 참고).
    const N = 31;
    const { textAt } = await parseSheet((sheet) => {
      sheet.getCell('A1').value = 1;
      for (let r = 2; r <= N; r += 1) {
        sheet.getCell(`A${r}`).value = { formula: `A${r - 1}+1`, result: undefined };
      }
    });
    expect(textAt(N, 1)).toBe(String(N));
    expect(textAt(2, 1)).toBe('2');
  });

  test('a formula whose range covers other FORMULA cells sums their computed values ' +
    '(the real file’s `N49=SUM(B49:M49)` shape, where B49..M49 are themselves `=SUM(...)`)', async () => {
    const { textAt } = await parseSheet((sheet) => {
      sheet.getCell('A1').value = 1;
      sheet.getCell('A2').value = 2;
      sheet.getCell('B1').value = 10;
      sheet.getCell('B2').value = 20;
      sheet.getCell('A3').value = { formula: 'SUM(A1:A2)', result: undefined };
      sheet.getCell('B3').value = { formula: 'SUM(B1:B2)', result: undefined };
      sheet.getCell('C3').value = { formula: 'SUM(A3:B3)', result: undefined };
    });
    expect(textAt(3, 1)).toBe('3');
    expect(textAt(3, 2)).toBe('30');
    expect(textAt(3, 3)).toBe('33');
  });

  test('a circular reference yields the #REF! error string instead of hanging or crashing', async () => {
    // tests/xlsxFormulas.test.ts의 기존 순환 참조 픽스처와 같은 모양(A1=B1, B1=A1).
    const { textAt } = await parseSheet((sheet) => {
      sheet.getCell('A1').value = { formula: 'B1', result: undefined };
      sheet.getCell('B1').value = { formula: 'A1', result: undefined };
    });
    expect(textAt(1, 1)).toBe('#REF!');
    expect(textAt(1, 2)).toBe('#REF!');
  });

  test('cells formatted with a thousands separator (#,##0) are still summed as NUMBERS — the ' +
    'display text "1,761" must not be read back as text and silently summed to 0', async () => {
    // data/originals/1a982ed5-…xlsx(실제 사용자 파일)의 재현 조건: B..M은 전부 nf="#,##0",
    // N은 =SUM(B:M). 표시 텍스트를 그대로 수식 엔진에 먹이면 Number("1,761")=NaN이라 합계가 0이 된다.
    const values = [1761, 1331, 1677, 1839, 1540, 1668, 1917, 1628];
    const { textAt } = await parseSheet((sheet) => {
      values.forEach((value, i) => {
        const cell = sheet.getCell(1, i + 1);
        cell.value = value;
        cell.numFmt = '#,##0';
      });
      const total = sheet.getCell(1, values.length + 1);
      total.value = { formula: `SUM(A1:${String.fromCharCode(64 + values.length)}1)`, result: undefined };
      total.numFmt = '#,##0';
    });
    expect(textAt(1, 1)).toBe('1,761'); // 표시 형식 자체는 그대로 유지된다
    // 합계는 실제 합(13,361)이어야 하고, 일반 숫자 칸과 **같은** 쉼표 서식으로 보여야 한다.
    expect(textAt(1, values.length + 1)).toBe('13,361');
  });

  test('a formula that Excel DID cache keeps showing the cached result (no behaviour change)', async () => {
    const { textAt } = await parseSheet((sheet) => {
      sheet.getCell('A1').value = 2;
      sheet.getCell('A2').value = { formula: 'A1*3', result: 6 };
    });
    expect(textAt(2, 1)).toBe('6');
  });

  test('a plain TEXT cell that merely starts with "=" is left alone (not evaluated)', async () => {
    const { textAt } = await parseSheet((sheet) => {
      sheet.getCell('A1').value = '=not a formula';
    });
    expect(textAt(1, 1)).toBe('=not a formula');
  });

  test('a formula the engine cannot evaluate degrades to an error string or the raw formula text, ' +
    'never to a blank cell', async () => {
    const { textAt } = await parseSheet((sheet) => {
      sheet.getCell('A1').value = { formula: 'TOTALLYUNKNOWNFN(1,2)', result: undefined };
    });
    const shown = textAt(1, 1);
    expect(shown).not.toBe('');
    expect(shown === '=TOTALLYUNKNOWNFN(1,2)' || shown.startsWith('#')).toBe(true);
  });

  // 리뷰가 확인한 버그의 실측 재현: computeGridCellDisplay가 parser.parse()에 하드코딩된
  // sheet:'Sheet1'을 넘기고 onCell/onRange가 ref.sheet를 무시한 채 호출자의 sheetIdx로만
  // 재귀해서, 시트를 넘나드는 참조가 조용히 "참조 중인 셀 자신"을 가리켜 순환 가드(#REF!)에
  // 걸리고 그 #REF!가 산술을 거치며 0이나 #VALUE!로 둔갑했다. 이 앱 자신의 xlsx export 형태
  // (캐시된 result 없이 formula만 있는)를 그대로 재현한다.
  describe('cross-sheet formula references (the bug this task fixes)', () => {
    async function parseTwoSheetWorkbook(
      build: (sheet1: ExcelJS.Worksheet, sheet2: ExcelJS.Worksheet) => void,
    ) {
      const workbook = new ExcelJS.Workbook();
      const sheet1 = workbook.addWorksheet('Sheet1');
      const sheet2 = workbook.addWorksheet('Sheet2');
      build(sheet1, sheet2);
      const buffer = await workbook.xlsx.writeBuffer();
      const grid = await parseXlsxToGrid(Buffer.from(buffer));
      // sheetNumber는 1-based(엑셀 관례, row/col과 같은 규칙) — Sheet1=1, Sheet2=2.
      const textAt = (sheetNumber: number, row: number, col: number): string => {
        const cell = grid.sheets[sheetNumber - 1]?.rows[row - 1]?.[col - 1];
        return cell === 0 || cell === undefined ? '' : cell.v;
      };
      return { grid, textAt };
    }

    test('Sheet2!A1 = SUM(Sheet1!A1:A1), no cached result — computes 2287, not 0 (exact brief ' +
      'reproduction)', async () => {
      const { textAt } = await parseTwoSheetWorkbook((sheet1, sheet2) => {
        sheet1.getCell('A1').value = 2287;
        sheet2.getCell('A1').value = { formula: 'SUM(Sheet1!A1:A1)', result: undefined };
      });
      expect(textAt(2, 1, 1)).toBe('2287');
    });

    test('a direct cross-sheet cell reference (Sheet2!A1 = "Sheet1!A1*1") computes 2287, not ' +
      '#VALUE! (the brief\'s second reproduction)', async () => {
      const { textAt } = await parseTwoSheetWorkbook((sheet1, sheet2) => {
        sheet1.getCell('A1').value = 2287;
        sheet2.getCell('A1').value = { formula: 'Sheet1!A1*1', result: undefined };
      });
      expect(textAt(2, 1, 1)).toBe('2287');
    });

    test('a formula referencing a sheet name that does not exist in the workbook degrades to ' +
      '#REF!, never a crash and never a silently wrong number', async () => {
      const { textAt } = await parseTwoSheetWorkbook((sheet1) => {
        sheet1.getCell('A1').value = { formula: 'DeletedSheet!A1', result: undefined };
      });
      expect(textAt(1, 1, 1)).toBe('#REF!');
    });

    test('same-sheet formulas in a MULTI-sheet workbook are completely unaffected — an unqualified ' +
      'reference on Sheet2 still resolves against Sheet2\'s own cells, not Sheet1\'s', async () => {
      const { textAt } = await parseTwoSheetWorkbook((sheet1, sheet2) => {
        sheet1.getCell('A1').value = 999; // 다른 값 — 잘못 해석되면 이게 섞여든다.
        sheet2.getCell('A1').value = 10;
        sheet2.getCell('A2').value = { formula: 'A1+1', result: undefined };
      });
      expect(textAt(2, 2, 1)).toBe('11');
    });
  });

  test('a few thousand formula rows parse in a reasonable time (no O(N²) blow-up)', async () => {
    // 3000행 × (독립 수식 1개 + 앞 행을 참조하는 체인형 수식 1개) — 체인형이 memo 없이 계산되면
    // 총 계산 횟수가 O(N²)로 폭증해 이 한계를 한참 넘긴다.
    const ROWS = 3000;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    for (let r = 1; r <= ROWS; r += 1) {
      sheet.getCell(r, 1).value = r;
      sheet.getCell(r, 2).value = { formula: `A${r}*2`, result: undefined };
      sheet.getCell(r, 3).value = r === 1 ? { formula: 'A1', result: undefined } : { formula: `C${r - 1}+A${r}`, result: undefined };
    }
    const buffer = await workbook.xlsx.writeBuffer();

    const start = Date.now();
    const grid = await parseXlsxToGrid(Buffer.from(buffer));
    const elapsedMs = Date.now() - start;

    const gridSheet = grid.sheets[0]!;
    const textAt = (row: number, col: number) => {
      const cell = gridSheet.rows[row - 1]?.[col - 1];
      return cell === 0 || cell === undefined ? '' : cell.v;
    };
    expect(textAt(ROWS, 2)).toBe(String(ROWS * 2));
    // 체인 끝 = 1..ROWS의 합
    expect(textAt(ROWS, 3)).toBe(String((ROWS * (ROWS + 1)) / 2));
    expect(elapsedMs).toBeLessThan(30_000);
  }, 60_000);
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
