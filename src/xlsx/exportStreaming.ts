import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { parseGridCellKey, type GridStructureOpRecord } from '../collab/gridShared.js';
import type { GridCell } from '../spreadsheet.js';
import { yieldToEventLoop, YIELD_EVERY_N_ROWS } from '../spreadsheet.js';
import { applyStyleToCell, cellStyle, cellText, coerceCellValue, toCellStyleAttrs, type XlsxExportOptions } from './collab.js';
import type { LazyGridStore } from './lazyGridStore.js';

/**
 * `exportXlsxFromGrid`(collab.ts)의 **오직 대용량 파일 전용 폴백**이다 — 정상 경로가 아니다.
 *
 * **왜 존재하는가**: 292,244행 × 36열, 서식이 거의 모든 셀에 걸린 실제 파일에서
 * `workbook.xlsx.writeBuffer()`가 `RangeError: Invalid array length`로 실패하는 것을 확인했다
 * (ExcelJS 4.4.0 자체의 대용량 재직렬화 한계 — 우리 코드 버그가 아니다, docs/TODO.md 참고).
 * `ExcelJS.stream.xlsx.WorkbookWriter`(행 단위로 스트리밍 직렬화)는 같은 파일에서 성공하고
 * 메모리도 훨씬 안정적이다(스파이크로 확인 완료).
 *
 * **왜 collab.ts의 exportXlsxFromGrid를 이걸로 통째로 바꾸지 않았는가**: 스트리밍 writer는
 * "기존 워크북을 불러와 그 자리에서 수정"하는 API가 없다 — 새 워크북을 처음부터 다시 쓰면서
 * 원본 시트를 행 단위로 그대로 옮겨야 한다. 이는 오늘 광범위하게 검증된 mid-sheet 행/열
 * 삽입·삭제 재생 로직(structuralOps)을 이 다른 API 위에서 처음부터 다시 만들고 다시 검증해야
 * 한다는 뜻이라 — 정상 크기 파일에 쓰는 이미 안전한 경로를 건드리는 대신, "그 경로가 실제로
 * 실패했을 때만" 쓰는 별도의, 기능이 좁은 폴백으로 분리했다(room.ts의 호출부 주석 참고).
 *
 * **이 폴백의 알려진 스코프 제한(의도적)**:
 * - **행/열 삽입·삭제(structuralOps)를 지원하지 않는다** — 이 정도로 거대한 파일에서 그런 구조
 *   변경까지 겪을 일은 실사용상 드물다고 보고, 정확성보다 스코프를 좁혀 위험을 줄였다.
 *   structuralOps가 하나라도 있으면 조용히 잘못된 파일을 만드는 대신 즉시 던진다(fail loud).
 * - 시트 보호(sheetProtection) 재적용을 하지 않는다.
 * - "한 번 편집된 칸은 계속 편집됨으로 남는다"는 dirtyCells의 트레이드오프는 collab.ts와 동일하다.
 *
 * 값은 버퍼가 아니라 **파일 스트림**으로 쓴다(그래서 애초에 무제한 배열 문제를 피한다) — 이 함수는
 * 로컬 임시 파일에 스트리밍으로 쓴 뒤 그 결과를 한 번에 읽어 돌려준다(임시 파일은 항상 정리한다).
 * 최종 반환 타입은 exportXlsxFromGrid와 동일한 `Uint8Array`라 room.ts 호출부는 폴백 분기 하나만
 * 추가하면 된다.
 *
 * Task 11부터 데이터 소스는 `canonicalGrid: GridDocument` 대신 `room.lazyGrid`(`LazyGridStore`)다 —
 * 편집된 칸은 `store.dirtyOverlay`를 직접 읽는다(옛 `dirtyCells: Set<string>` 인자는 필요 없다,
 * 오버레이의 키·값 자체가 이미 그 정보다). 이 폴백은 구조연산을 지원하지 않으므로(아래 스코프
 * 제한 참고) `store.logicalToOriginalRow/Col`을 참조할 필요가 없다 — "새 행"(원본 범위 밖)의
 * 유일한 가능한 소스는 dirtyOverlay뿐이다.
 */
export async function exportXlsxFromGridStreaming(
  originalXlsxBytes: Uint8Array,
  store: LazyGridStore,
  structuralOps: readonly GridStructureOpRecord[],
  options: XlsxExportOptions = {},
): Promise<Uint8Array> {
  if (structuralOps.length > 0) {
    throw new Error(
      '대용량 파일 스트리밍 저장은 행/열 삽입·삭제를 아직 지원하지 않습니다 (structuralOps.length > 0).',
    );
  }
  const removedSheets = options.removedSheetIndices ?? new Set<number>();

  const src = new ExcelJS.Workbook();
  await src.xlsx.load(originalXlsxBytes as unknown as ArrayBuffer);
  const srcSheets = src.worksheets.filter((sheet) => !sheet.state || sheet.state === 'visible');
  const originalSheetCount = srcSheets.length;

  // 편집된 칸을 시트+행별로 미리 묶는다 — collab.ts의 exportXlsxFromGrid와 같은 이유. 값 자체를
  // 함께 들고 있으므로(옛 "row:col" 좌표 집합과 달리) 아래 쓰기 지점에서 store를 다시 조회할
  // 필요가 없다.
  const dirtyByRow = new Map<string, Array<{ col: number; cell: GridCell | 0 }>>(); // "sheetIdx:row" -> [{col, cell}]
  for (const [key, cell] of store.dirtyOverlay) {
    const parsed = parseGridCellKey(key);
    if (!parsed || removedSheets.has(parsed.sheetIdx)) continue;
    const rowKey = `${parsed.sheetIdx}:${parsed.row}`;
    let list = dirtyByRow.get(rowKey);
    if (!list) {
      list = [];
      dirtyByRow.set(rowKey, list);
    }
    list.push({ col: parsed.col, cell });
  }

  const tempPath = path.join(os.tmpdir(), `docview-xlsx-stream-${randomUUID()}.xlsx`);
  const out = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: tempPath,
    useStyles: true,
    useSharedStrings: true,
  });

  try {
    // 시트 탭 순서(sheetDisplayOrder) 반영 — 스트리밍 writer는 사후에 순서를 바꿀 방법이 없으므로
    // (ExcelJS.Workbook의 orderNo 재배치 같은 API가 없다) 애초에 원하는 순서대로 addWorksheet를
    // 부른다. WorkbookWriter는 addWorksheet가 불린 순서대로 시트를 써 내려간다.
    const sheetIndices = [...store.sheetMeta.keys()].filter((i) => !removedSheets.has(i));
    sheetIndices.sort((a, b) => {
      const orderA = options.sheetDisplayOrder?.get(a) ?? a;
      const orderB = options.sheetDisplayOrder?.get(b) ?? b;
      return orderA - orderB;
    });

    for (const sheetIdx of sheetIndices) {
      const target = store.sheetMeta[sheetIdx];
      const srcSheet = srcSheets[sheetIdx] as ExcelJS.Worksheet | undefined;
      const dstSheet = out.addWorksheet(target.name, srcSheet ? { views: srcSheet.views } : undefined);

      if (srcSheet?.columns) {
        srcSheet.columns.forEach((col, i) => {
          if (col?.width) dstSheet.getColumn(i + 1).width = col.width;
        });
      }

      const originalRowCount = srcSheet?.rowCount ?? 0;

      for (let r = 1; r <= target.rowCount; r += 1) {
        const dirtyInRow = dirtyByRow.get(`${sheetIdx}:${r - 1}`);
        const dstRow = dstSheet.getRow(r);

        if (srcSheet && r <= originalRowCount) {
          // 원본에 있던 행 — 값·서식을 그대로 옮기고, 실제로 편집된 칸만 store의 현재 값으로 덮어쓴다.
          const srcRow = srcSheet.getRow(r);
          srcRow.eachCell({ includeEmpty: true }, (srcCell, colNumber) => {
            if (colNumber > target.colCount) return;
            const dstCell = dstRow.getCell(colNumber);
            dstCell.value = srcCell.value;
            if (srcCell.style) dstCell.style = srcCell.style;
          });
          if (srcRow.height) dstRow.height = srcRow.height;
          if (dirtyInRow) {
            for (const { col, cell } of dirtyInRow) {
              if (col >= target.colCount) continue;
              writeExportCell(dstRow, col, cell, store);
            }
          }
        } else {
          // 원본에 없던 행(직접 rowCount를 늘려 시트 끝에 추가된 행, 또는 통째로 새 시트) — 옮겨올
          // 원본이 없으므로, 이 행의 유일한 가능한 값 출처인 dirtyOverlay만 쓴다(원본 매핑이 없는
          // 칸은 편집되지 않았다면 애초에 값이 존재할 수 없다 — LazyGridStore.getCell과 동일한 이유).
          if (dirtyInRow) {
            for (const { col, cell } of dirtyInRow) {
              if (col >= target.colCount || cell === 0 || cell === undefined) continue;
              writeExportCell(dstRow, col, cell, store);
            }
          }
        }

        dstRow.commit();
        if (r % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
      }

      dstSheet.commit();
    }

    await out.commit();
    return new Uint8Array(await fs.readFile(tempPath));
  } finally {
    await fs.rm(tempPath, { force: true, maxRetries: 5, retryDelay: 100 }).catch(() => {
      /* 임시 파일 정리 실패는 치명적이지 않다 — OS 임시 폴더 청소에 맡긴다 */
    });
  }
}

function writeExportCell(
  dstRow: ExcelJS.Row,
  col: number,
  cell: GridCell | 0,
  store: LazyGridStore,
): void {
  const text = cellText(cell);
  const dstCell = dstRow.getCell(col + 1);
  dstCell.value = text === '' ? null : coerceCellValue(text);
  const style = cellStyle(cell, store);
  applyStyleToCell(dstCell, style ? toCellStyleAttrs(style) : undefined);
}
