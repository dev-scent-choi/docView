import { randomUUID } from 'node:crypto';
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { exportXlsxFromGrid } from '../src/xlsx/collab.js';
import { growSheet, ingestXlsxLazily, setDirtyCell } from '../src/xlsx/lazyGridStore.js';

/**
 * 최종 전체 브랜치 리뷰의 I1을 붙잡는 테스트 — `exportXlsxFromGrid`가 큰 시트에서 자동저장·다운로드
 * 마다 이벤트 루프를 통째로 멈추던 문제.
 *
 * LazyGridStore 재설계(2026-09-09) 이후로는 값 diff 루프가 `store.dirtyOverlay`(편집된 칸만)를
 * 직접 순회한다 — 그래서 이 테스트도 "많은 칸이 편집됐을 때 그 쓰기 루프가 양보하는지"를
 * 검증한다.
 *
 * **측정 방법**: ExcelJS의 zip/stream 자체가 배경에서 setImmediate를 많이 쓰기 때문에(그래서
 * xlsxYielding.test.ts는 parseXlsxToGrid에 대한 절대 틱 임계값 검증을 실증적으로 기각했다) 단순히
 * "틱 수가 많다"로는 우리 쪽 양보를 분리해 셀 수 없다. 대신 **쓰기 지점 자체를 계측한다**:
 * `ExcelJS.Worksheet.prototype.getCell`(값 쓰기 루프가 편집된 칸마다 부르는 지점)를 감싸 호출될
 * 때마다 "그 시점까지 이벤트 루프가 몇 번 돌았는지"를 기록한다. 양보가 전혀 없으면 모든 호출이
 * **같은 틱 값**을 갖는다(그 사이 루프가 한 번도 못 돌았다는 직접 증거). 양보가 있으면 청크
 * 경계마다 틱 값이 올라간다.
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

async function makeTinyOriginal(): Promise<{ bytes: Buffer; worksheetPrototype: ExcelJS.Worksheet }> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Big');
  sheet.getCell('A1').value = 'row-0';
  const bytes = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
  // ExcelJS는 Worksheet 클래스를 export하지 않으므로, 실제 인스턴스에서 프로토타입을 구한다 —
  // exportXlsxFromGrid 내부에서 새로 만드는 워크북의 시트도 같은 프로토타입을 공유한다.
  const worksheetPrototype = Object.getPrototypeOf(sheet) as ExcelJS.Worksheet;
  return { bytes, worksheetPrototype };
}

describe('xlsx/collab.ts: exportXlsxFromGrid yields to the event loop while writing many edited cells', () => {
  test(
    'the dirty-cell write pass hits real yield points when many cells were edited',
    async () => {
      const ROWS = 10_000; // YIELD_EVERY_N_ROWS(2000) 경계를 4번 넘긴다
      const COLS = 3;
      const { bytes: original, worksheetPrototype } = await makeTinyOriginal();

      const store = await ingestXlsxLazily(randomUUID(), original);
      // 원본은 1행짜리 — 그리드 에디터의 "행 추가"와 같은 모양으로 ROWS행까지 직접 확장한다
      // (growSheet는 매핑 배열만 -1로 채운다: 새로 생긴 행은 전부 원본이 없는 칸이 된다).
      growSheet(store, 0, ROWS, COLS);
      // 모든 행의 0번 열을 "편집됨"으로 표시 — 예전 전체-스캔 테스트와 같은 규모(ROWS개 항목)로
      // 새 쓰기 루프(dirtyOverlay 기반)가 청크마다 실제로 양보하는지 검증한다.
      for (let r = 0; r < ROWS; r += 1) setDirtyCell(store, 0, r, 0, { v: `row-${r}` });

      const ticker = startTicker();
      const ticksAtCellWrite: number[] = [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const originalGetCell = (worksheetPrototype as any).getCell;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (worksheetPrototype as any).getCell = function (...args: unknown[]) {
        if (typeof args[0] === 'number') ticksAtCellWrite.push(ticker.count());
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return originalGetCell.apply(this, args);
      };

      try {
        // 원본은 1행짜리라 shrink-check는 트리거되지 않는다(target.rowCount가 baseline보다 훨씬 크다).
        const baselineSheetExtents = [{ rowCount: 1, colCount: 1 }];
        await exportXlsxFromGrid(original, store, baselineSheetExtents, []);
      } finally {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (worksheetPrototype as any).getCell = originalGetCell;
        ticker.stop();
      }

      // 편집된 칸 전체를 실제로 훑었는지(계측이 붙긴 했는지) 먼저 확인한다. 칸마다 값·서식으로
      // getCell이 두 번씩 불리므로 최소 ROWS번은 기록돼야 한다.
      expect(ticksAtCellWrite.length).toBeGreaterThanOrEqual(ROWS);

      // 양보 지점이 없던 예전 코드에서는 이 값이 1이다 — 전체 쓰기가 한 호흡의 동기 블록이라
      // 그동안 이벤트 루프가 단 한 번도 돌지 못하기 때문이다.
      const distinctTicks = new Set(ticksAtCellWrite).size;
      expect(distinctTicks).toBeGreaterThanOrEqual(3);

      // "한 번도 안 쉬고 연달아 훑은 최대 호출 수"도 직접 확인한다 — 청크 크기(2000개, 칸당 두
      // 번 호출되므로 여유를 둔다)보다 크게 벗어나면 어딘가 양보가 빠진 것이다.
      let longestUnyieldedRun = 1;
      let currentRun = 1;
      for (let i = 1; i < ticksAtCellWrite.length; i += 1) {
        currentRun = ticksAtCellWrite[i] === ticksAtCellWrite[i - 1] ? currentRun + 1 : 1;
        longestUnyieldedRun = Math.max(longestUnyieldedRun, currentRun);
      }
      expect(longestUnyieldedRun).toBeLessThanOrEqual(4200);
    },
    120_000,
  );
});
