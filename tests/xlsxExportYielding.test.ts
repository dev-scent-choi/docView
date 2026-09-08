import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { exportXlsxFromGrid } from '../src/xlsx/collab.js';
import type { GridCell, GridDocument, GridSheet } from '../src/spreadsheet.js';

/**
 * 최종 전체 브랜치 리뷰의 I1을 붙잡는 테스트 — `exportXlsxFromGrid`의 두 전체 그리드 스캔(값 diff,
 * 서식 diff)에 양보 지점이 하나도 없어, 큰 시트에서 자동저장·다운로드마다 이벤트 루프가 통째로
 * 멈추던 문제.
 *
 * **측정 방법**: ExcelJS의 zip/stream 자체가 배경에서 setImmediate를 많이 쓰기 때문에(그래서
 * xlsxYielding.test.ts는 parseXlsxToGrid에 대한 절대 틱 임계값 검증을 실증적으로 기각했다) 단순히
 * "틱 수가 많다"로는 우리 쪽 양보를 분리해 셀 수 없다. 대신 **스캔 대상 자체를 계측한다**: 시트의
 * rows 배열을 Proxy로 감싸 export가 행 하나를 읽을 때마다 "그 시점까지 이벤트 루프가 몇 번 돌았는지"를
 * 기록한다. 양보가 전혀 없으면 모든 행 접근이 **같은 틱 값**을 갖는다(그 사이 루프가 한 번도 못
 * 돌았다는 직접 증거). 양보가 있으면 청크 경계마다 틱 값이 올라간다.
 */
function startTicker(): { count: () => number; stop: () => number } {
  let ticks = 0;
  let running = true;
  function schedule(): void {
    setImmediate(() => {
      if (!running) return;
      ticks += 1;
      schedule();
    });
  }
  schedule();
  return {
    count: () => ticks,
    stop: () => {
      running = false;
      return ticks;
    },
  };
}

function makeRows(rowCount: number, colCount: number): Array<Array<GridCell | 0>> {
  const rows: Array<Array<GridCell | 0>> = [];
  for (let r = 0; r < rowCount; r += 1) {
    const row: Array<GridCell | 0> = new Array(colCount).fill(0);
    row[0] = { v: `row-${r}` };
    rows.push(row);
  }
  return rows;
}

function makeSheet(name: string, rows: Array<Array<GridCell | 0>>, colCount: number): GridSheet {
  return {
    name,
    rowCount: rows.length,
    colCount,
    colWidths: [],
    merges: [],
    freezeRows: 0,
    freezeCols: 0,
    rows,
  };
}

async function makeTinyOriginal(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Big');
  sheet.getCell('A1').value = 'row-0';
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('xlsx/collab.ts: exportXlsxFromGrid yields to the event loop while scanning a large grid', () => {
  test(
    'both full-grid passes (value diff, style diff) hit real yield points',
    async () => {
      const ROWS = 10_000; // YIELD_EVERY_N_ROWS(2000) 경계를 각 패스마다 4번 넘긴다
      const COLS = 3;
      const original = await makeTinyOriginal();

      const canonicalRows = makeRows(ROWS, COLS);
      const baselineRows = makeRows(ROWS, COLS); // 내용이 동일 → ExcelJS 쓰기는 한 번도 안 일어나고 순수 스캔만 남는다

      const ticker = startTicker();
      const ticksAtRowRead: number[] = [];
      const instrumentedRows = new Proxy(canonicalRows, {
        get(target, prop, receiver) {
          if (typeof prop === 'string' && /^\d+$/.test(prop)) ticksAtRowRead.push(ticker.count());
          return Reflect.get(target, prop, receiver);
        },
      });

      const canonicalGrid: GridDocument = {
        styles: [],
        sheets: [makeSheet('Big', instrumentedRows, COLS)],
      };
      const baselineGrid: GridDocument = {
        styles: [],
        sheets: [makeSheet('Big', baselineRows, COLS)],
      };

      await exportXlsxFromGrid(original, canonicalGrid, baselineGrid);
      ticker.stop();

      // 두 패스 모두 실제로 전체 행을 훑었는지(계측이 붙긴 했는지) 먼저 확인한다.
      expect(ticksAtRowRead.length).toBeGreaterThanOrEqual(ROWS * 2);

      // 양보 지점이 없던 예전 코드에서는 이 값이 1이다 — 전체 스캔이 한 호흡의 동기 블록이라
      // 그동안 이벤트 루프가 단 한 번도 돌지 못하기 때문이다.
      const distinctTicks = new Set(ticksAtRowRead).size;
      expect(distinctTicks).toBeGreaterThanOrEqual(6);

      // "한 번도 안 쉬고 연달아 훑은 최대 행 수"도 직접 확인한다 — 청크 크기(2000행)보다
      // 크게 벗어나면 어딘가 양보가 빠진 것이다(약간의 여유는 둔다).
      let longestUnyieldedRun = 1;
      let currentRun = 1;
      for (let i = 1; i < ticksAtRowRead.length; i += 1) {
        currentRun = ticksAtRowRead[i] === ticksAtRowRead[i - 1] ? currentRun + 1 : 1;
        longestUnyieldedRun = Math.max(longestUnyieldedRun, currentRun);
      }
      expect(longestUnyieldedRun).toBeLessThanOrEqual(2100);
    },
    120_000,
  );
});
