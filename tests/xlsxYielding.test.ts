import ExcelJS from 'exceljs';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { parseXlsxToGrid } from '../src/spreadsheet.js';
import { seedGridFromXlsx } from '../src/xlsx/collab.js';

/**
 * 큰 xlsx를 만든다 — YIELD_EVERY_N_ROWS(2000)보다 훨씬 많은 행을 둬서, 양보 지점이 없다면
 * 파싱/시딩이 한 번도 안 쉬는 긴 동기 구간이 될 만큼 크게 잡는다.
 */
async function makeLargeSampleXlsx(rows: number): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Big');
  for (let r = 1; r <= rows; r += 1) {
    sheet.getCell(r, 1).value = `row-${r}`;
    sheet.getCell(r, 2).value = r;
  }
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

/**
 * 이벤트 루프에 대기 중인 setImmediate 콜백을 처리할 기회가 실제로 주어지는 횟수를 센다.
 * 측정 대상 작업이 한 번도 안 쉬는 동기 블록이라면, 그 작업이 끝날 때까지 이 재귀 티커는
 * 단 한 번도(또는 아주 적게) 못 돈다 — 반대로 작업이 주기적으로 await로 제어권을 돌려주면
 * 그때마다 이 티커도 끼어들어 돈다.
 */
function startTicker(): { stop: () => number } {
  let ticks = 0;
  let running = true;
  function schedule() {
    setImmediate(() => {
      if (!running) return;
      ticks += 1;
      schedule();
    });
  }
  schedule();
  return {
    stop: () => {
      running = false;
      return ticks;
    },
  };
}

describe('spreadsheet.ts: parseXlsxToGrid still parses correctly after the yield-point refactor', () => {
  test('correctly parses a sheet spanning many yield-chunk boundaries', async () => {
    // 정확성 검증: eachSheet(콜백형)를 workbook.worksheets 순회로, sheet.eachRow를 수동 for
    // 루프로 바꾼 구조 변경이 기존 동작과 동일한 결과를 내는지 — 특히 YIELD_EVERY_N_ROWS(2000)
    // 경계를 여러 번 넘는 행 수로 첫/중간/마지막 행이 모두 올바른지 확인한다.
    //
    // (setImmediate 틱 수를 세어 "실제로 양보했는지"를 이 함수에서 직접 검증하는 건 포기했다 —
    // 실측해보니 ExcelJS/Node 자체의 배경 setImmediate 사용이 행 수와 무관하게도, 어느 정도
    // 행 수에 비례해서도 이미 상당히 커서(0행짜리 빈 시트조차 틱 40회) 우리 쪽 양보 지점만
    // 분리해 셀 수 있는 신뢰할 만한 신호가 없었다 — 두 가지 접근(절대 임계값, 파일 크기별
    // 증가폭) 모두 픽스를 되돌려도 통과해버려 직접 실증해 기각했다. 같은 yieldToEventLoop
    // 메커니즘은 xlsx/collab.ts의 seedGridFromXlsx 테스트에서 깨끗하게 검증된다 — Y.Map 연산은
    // 순수 동기 JS라 배경 잡음이 없어 "픽스 없으면 반드시 실패"가 실제로 성립한다.)
    const rows = 5000;
    const xlsxBytes = await makeLargeSampleXlsx(rows);
    const grid = await parseXlsxToGrid(xlsxBytes);

    expect(grid.sheets[0].rowCount).toBe(rows);
    expect(grid.sheets[0].rows[0][0]).toEqual({ v: 'row-1' });
    expect(grid.sheets[0].rows[2000][0]).toEqual({ v: 'row-2001' }); // 첫 양보 경계 바로 다음 행
    expect(grid.sheets[0].rows[rows - 1][0]).toEqual({ v: `row-${rows}` });
  });
});

describe('xlsx/collab.ts: seedGridFromXlsx does not block the event loop for the whole seed', () => {
  test('yields to the event loop multiple times while seeding a large sheet into Yjs', async () => {
    const rows = 5000;
    const xlsxBytes = await makeLargeSampleXlsx(rows);
    const grid = await parseXlsxToGrid(xlsxBytes);
    const ydoc = new Y.Doc();

    const ticker = startTicker();
    await seedGridFromXlsx(ydoc, grid);
    const ticks = ticker.stop();

    const cells = ydoc.getMap<string>('cells');
    expect(cells.get('0:0:0')).toBe('row-1');
    expect(cells.get(`0:${rows - 1}:0`)).toBe(`row-${rows}`);
    // 2000행마다 트랜잭션 하나 + 양보 1회 — 5000행이면 최소 2번은 나와야 한다.
    expect(ticks).toBeGreaterThanOrEqual(2);
  });
});
