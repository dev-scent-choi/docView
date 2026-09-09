import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { gridCellKey } from '../src/collab/gridShared.js';
import { exportXlsxFromGrid } from '../src/xlsx/collab.js';
import type { GridCell, GridDocument, GridSheet } from '../src/spreadsheet.js';

/**
 * 최종 전체 브랜치 리뷰의 I1을 붙잡는 테스트 — `exportXlsxFromGrid`가 큰 시트에서 자동저장·다운로드
 * 마다 이벤트 루프를 통째로 멈추던 문제.
 *
 * baselineGrid 전체 복제본을 없앤 재설계(2026-09-09) 이후로는 "값이 같은지 매 칸 비교"하는 전체
 * 스캔 자체가 사라지고, 편집된 칸(dirtyCells)만 훑는 구조가 됐다 — 그래서 이 테스트도 "많은 칸이
 * 편집됐을 때 그 쓰기 루프가 양보하는지"를 검증하도록 바꿨다(예전에는 "값 diff/서식 diff 전체
 * 스캔"이 대상이었다).
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

describe('xlsx/collab.ts: exportXlsxFromGrid yields to the event loop while writing many edited cells', () => {
  test(
    'the dirty-cell write pass hits real yield points when many cells were edited',
    async () => {
      const ROWS = 10_000; // YIELD_EVERY_N_ROWS(2000) 경계를 4번 넘긴다
      const COLS = 3;
      const original = await makeTinyOriginal();

      const canonicalRows = makeRows(ROWS, COLS);

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
      // 원본은 1행짜리라 shrink-check는 트리거되지 않는다(target.rowCount가 baseline보다 훨씬 크다).
      const baselineSheetExtents = [{ rowCount: 1, colCount: 1 }];
      // 모든 행의 0번 열을 "편집됨"으로 표시 — 예전 전체-스캔 테스트와 같은 규모(ROWS개 항목)로
      // 새 쓰기 루프(dirtyCells 기반)가 청크마다 실제로 양보하는지 검증한다.
      const dirtyCells = new Set<string>();
      for (let r = 0; r < ROWS; r += 1) dirtyCells.add(gridCellKey(0, r, 0));

      await exportXlsxFromGrid(original, canonicalGrid, baselineSheetExtents, dirtyCells, []);
      ticker.stop();

      // 편집된 칸 전체를 실제로 훑었는지(계측이 붙긴 했는지) 먼저 확인한다.
      expect(ticksAtRowRead.length).toBeGreaterThanOrEqual(ROWS);

      // 양보 지점이 없던 예전 코드에서는 이 값이 1이다 — 전체 쓰기가 한 호흡의 동기 블록이라
      // 그동안 이벤트 루프가 단 한 번도 돌지 못하기 때문이다.
      const distinctTicks = new Set(ticksAtRowRead).size;
      expect(distinctTicks).toBeGreaterThanOrEqual(3);

      // "한 번도 안 쉬고 연달아 훑은 최대 행 수"도 직접 확인한다 — 청크 크기(2000개)보다
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
