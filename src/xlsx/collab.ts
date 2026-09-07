import ExcelJS from 'exceljs';
import * as Y from 'yjs';
import {
  GRID_BASELINE_MAP,
  GRID_CELLS_MAP,
  GRID_MERGES_MAP,
  GRID_STRUCTURE_MAP,
  GRID_STYLES_BASELINE_MAP,
  GRID_STYLES_MAP,
  gridCellKey,
  type CellStyleAttrs,
  type GridMergeShape,
  type SheetStructure,
} from '../collab/gridShared.js';
import type { GridCell, GridDocument, GridStyle } from '../spreadsheet.js';
import { parseXlsxToGrid, yieldToEventLoop, YIELD_EVERY_N_ROWS } from '../spreadsheet.js';


function toCellStyleAttrs(style: GridStyle): CellStyleAttrs {
  const attrs: { -readonly [K in keyof CellStyleAttrs]?: CellStyleAttrs[K] } = {};
  if (style.b) attrs.b = 1;
  if (style.i) attrs.i = 1;
  if (style.u) attrs.u = 1;
  if (style.st) attrs.st = 1;
  if (style.col) attrs.col = style.col;
  if (style.bg) attrs.bg = style.bg;
  if (style.a) attrs.a = style.a;
  if (style.va) attrs.va = style.va;
  if (style.w) attrs.wrap = 1;
  if (style.ff) attrs.ff = style.ff;
  if (style.fs) attrs.fs = style.fs;
  if (style.nf) attrs.nf = style.nf;
  if (style.nd) attrs.nd = style.nd;
  if (style.lk) attrs.lk = 1;
  if (style.nu) attrs.nu = 1;
  if (style.bt) attrs.bt = style.bt;
  if (style.br) attrs.br = style.br;
  if (style.bb) attrs.bb = style.bb;
  if (style.bl) attrs.bl = style.bl;
  return attrs;
}

/** 병합 범위(1-based, {r,c,rs,cs})에서 "앵커가 아닌, 병합에 덮인" 칸의 (0-based r:c) 키 집합을 만든다.
 * 이런 칸은 ExcelJS가 앵커와 동일한 값을 그대로 돌려주는 것뿐(진짜 값이 아님)이라 시딩에서 제외해야
 * 병합이 "14칸에 같은 제목이 중복 입력된 것"이 아니라 진짜 병합으로 남는다(그리드 에디터가 렌더링
 * 시 앵커에만 colspan/rowspan을 주고 이 칸들은 아예 그리지 않는다 — viewer.js의 처리와 동일한 원리). */
function mergeCoveredCells(merges: readonly GridMergeShape[]): Set<string> {
  const covered = new Set<string>();
  for (const m of merges) {
    const r0 = m.r - 1;
    const c0 = m.c - 1;
    for (let r = r0; r < r0 + m.rs; r += 1) {
      for (let c = c0; c < c0 + m.cs; c += 1) {
        if (r === r0 && c === c0) continue;
        covered.add(`${r}:${c}`);
      }
    }
  }
  return covered;
}

/** parseXlsxToGrid가 만든 표시용 텍스트(예: "2,287")를 협업 문서의 실제 셀 값으로 그대로 심으면
 * 안 된다 — Number("2,287")는 NaN이라, 이 값을 참조하는 다른 수식(SUM 등)이 이 셀을 숫자가 아닌
 * 것으로 취급해 조용히 0으로 계산해버린다(실제 재현된 버그: 합계 열이 항상 0으로 보임). 쉼표 구분
 * 표시 형식(number/currency)이 걸린 칸만, 화면표시용 텍스트에서 쉼표를 제거해 원래 숫자 문자열로
 * 되돌린 뒤 심는다 — "화면에 어떻게 보일지"는 style.nf/nd로 이미 별도로 심어지므로 여기서는 순수
 * 값만 복원하면 된다. */
function seedableCellValue(text: string, style: GridStyle | undefined): string {
  if (!style?.nf || (style.nf !== 'number' && style.nf !== 'currency')) return text;
  const stripped = text.replace(/[₩,\s]/g, '');
  return Number.isFinite(Number(stripped)) && stripped !== '' ? stripped : text;
}

/** 읽기전용 그리드 뷰(parseXlsxToGrid의 결과)를 협업용 Y.Doc에 시딩한다.
 * 큰 시트를 셀 하나하나 동기로 심으면 그동안 서버 전체가 멈추므로(spreadsheet.ts의
 * yieldToEventLoop 주석 참고), 행을 YIELD_EVERY_N_ROWS개씩 묶어 트랜잭션 하나로 심고 그 사이
 * 사이에 이벤트 루프로 제어권을 돌려준다. */
export async function seedGridFromXlsx(ydoc: Y.Doc, grid: GridDocument): Promise<void> {
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const baseline = ydoc.getMap<string>(GRID_BASELINE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const stylesBaseline = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_BASELINE_MAP);
  const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  const merges = ydoc.getMap<readonly GridMergeShape[]>(GRID_MERGES_MAP);

  for (let sheetIdx = 0; sheetIdx < grid.sheets.length; sheetIdx += 1) {
    const sheet = grid.sheets[sheetIdx];
    ydoc.transact(() => {
      structure.set(String(sheetIdx), {
        name: sheet.name,
        rowCount: sheet.rowCount,
        colCount: sheet.colCount,
        freezeRows: sheet.freezeRows,
        freezeCols: sheet.freezeCols,
        colWidths: sheet.colWidths,
      });
      if (sheet.merges.length > 0) merges.set(String(sheetIdx), sheet.merges);
    });
    const covered = mergeCoveredCells(sheet.merges);

    for (let start = 0; start < sheet.rows.length; start += YIELD_EVERY_N_ROWS) {
      const end = Math.min(start + YIELD_EVERY_N_ROWS, sheet.rows.length);
      ydoc.transact(() => {
        for (let r = start; r < end; r += 1) {
          const row = sheet.rows[r];
          for (let c = 0; c < row.length; c += 1) {
            const cell = row[c];
            if (cell === 0) continue;
            if (covered.has(`${r}:${c}`)) continue; // 병합에 덮인 칸 — 앵커의 중복 값이라 시딩하지 않는다
            const key = gridCellKey(sheetIdx, r, c);
            const style = cell.s !== undefined ? grid.styles[cell.s] : undefined;
            if (cell.v !== '') {
              const value = seedableCellValue(cell.v, style);
              cells.set(key, value);
              baseline.set(key, value);
            }
            if (style) {
              const attrs = toCellStyleAttrs(style);
              if (Object.keys(attrs).length > 0) {
                styles.set(key, attrs);
                stylesBaseline.set(key, attrs);
              }
            }
          }
        }
      });
      await yieldToEventLoop();
    }
  }
}

/** '='로 시작하면 수식으로 저장한다(계산은 하지 않음 — 실제 결과는 저장 후 엑셀/뷰어에서 재계산된다). */
function coerceCellValue(raw: string): string | number | { formula: string } {
  const trimmed = raw.trim();
  if (trimmed.startsWith('=') && trimmed.length > 1) {
    return { formula: trimmed.slice(1) };
  }
  if (trimmed !== '' && /^-?\d+(\.\d+)?$/.test(trimmed)) {
    const num = Number(trimmed);
    if (Number.isFinite(num)) return num;
  }
  return raw;
}

const ALIGN_TO_EXCEL: Record<'l' | 'c' | 'r', ExcelJS.Alignment['horizontal']> = {
  l: 'left',
  c: 'center',
  r: 'right',
};
const VALIGN_TO_EXCEL: Record<'t' | 'm' | 'b', ExcelJS.Alignment['vertical']> = {
  t: 'top',
  m: 'middle',
  b: 'bottom',
};
const BORDER_WEIGHT_TO_EXCEL: Record<1 | 2, ExcelJS.BorderStyle> = {
  1: 'thin',
  2: 'medium',
};
/** spreadsheet.ts의 PT_TO_PX(96/72)와 정반대 변환 — 저장은 항상 px 기준(CellStyleAttrs.fs)이라
 * 실제 엑셀 파일에 쓸 때만 pt로 되돌린다. */
const PX_TO_PT = 72 / 96;

function numFmtStringFor(style: CellStyleAttrs): string | undefined {
  if (!style.nf) return undefined;
  const decimals = style.nd ? `.${'0'.repeat(style.nd)}` : '';
  switch (style.nf) {
    case 'number': return `#,##0${decimals}`;
    case 'currency': return `"₩"#,##0${decimals}`;
    case 'percent': return `0${decimals}%`;
    default: return undefined;
  }
}

function applyStyleToCell(cell: ExcelJS.Cell, style: CellStyleAttrs | undefined): void {
  if (!style) {
    cell.font = {};
    cell.fill = { type: 'pattern', pattern: 'none' };
    cell.alignment = {};
    cell.border = {};
    cell.numFmt = 'General';
    return;
  }
  cell.font = {
    bold: style.b === 1,
    italic: style.i === 1,
    underline: style.u === 1,
    strike: style.st === 1,
    color: style.col ? { argb: `FF${style.col.replace('#', '').toUpperCase()}` } : undefined,
    name: style.ff,
    size: style.fs ? Math.round(style.fs * PX_TO_PT) : undefined,
  };
  cell.fill = style.bg
    ? { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${style.bg.replace('#', '').toUpperCase()}` } }
    : { type: 'pattern', pattern: 'none' };
  cell.alignment = {
    horizontal: style.a ? ALIGN_TO_EXCEL[style.a] : undefined,
    vertical: style.va ? VALIGN_TO_EXCEL[style.va] : undefined,
    wrapText: style.wrap === 1,
  };
  cell.border = {
    top: style.bt ? { style: BORDER_WEIGHT_TO_EXCEL[style.bt] } : undefined,
    right: style.br ? { style: BORDER_WEIGHT_TO_EXCEL[style.br] } : undefined,
    bottom: style.bb ? { style: BORDER_WEIGHT_TO_EXCEL[style.bb] } : undefined,
    left: style.bl ? { style: BORDER_WEIGHT_TO_EXCEL[style.bl] } : undefined,
  };
  cell.numFmt = numFmtStringFor(style) ?? 'General';
}

/** canonicalGrid/baselineGrid의 셀(GridCell | 0)에서 표시 텍스트만 뽑는다 — 빈 칸(0 sentinel)과
 * baseline 쪽 시트 자체가 없는 경우(새로 추가된 시트) 모두 빈 문자열로 취급한다. */
function cellText(cell: GridCell | 0 | undefined): string {
  return cell === 0 || cell === undefined ? '' : cell.v;
}

/** canonicalGrid/baselineGrid의 셀이 참조하는 스타일 인덱스를 실제 GridStyle로 해석한다 —
 * grid.styles는 공용 스타일 배열이라 셀은 인덱스(.s)만 들고 있다(Yjs styles 맵과 달리 셀마다
 * 통짜 서식 객체를 직접 들고 있지 않다). */
function cellStyle(cell: GridCell | 0 | undefined, grid: GridDocument): GridStyle | undefined {
  if (cell === 0 || cell === undefined || cell.s === undefined) return undefined;
  return grid.styles[cell.s];
}

/**
 * 원본 xlsx(수식·서식 등 우리가 편집하지 않는 모든 것)를 그대로 로드한 뒤,
 * 실제로 사용자가 바꾼 셀(값·서식 각각)만 덮어써 재직렬화한다 — docx export와 같은
 * "바뀐 부분만 교체" 원칙. baselineGrid와 값이 같은 셀(한 번도 편집되지 않은 셀 — 수식 셀 포함)은
 * 건드리지 않는다: 수식 셀은 시딩 시 계산된 결괏값 문자열로만 채워지므로, 건드리지 않았는데도
 * canonicalGrid에 있다고 전부 써버리면 수식 자체가 정적 값으로 깨진다.
 * 행/열 삭제(구조 축소)는 시트 끝에서부터만 지원한다(v1 스코프 단순화).
 *
 * 데이터 소스는 Yjs Y.Doc이 아니라 canonicalGrid/baselineGrid다(Task 3) — canonicalGrid는 룸이
 * 항상 최신 상태로 유지하는 "단일 진실"이고, baselineGrid는 그 룸이 처음 시딩될 때의 깊은 복사본
 * (한 번도 편집되지 않은 셀을 가려내는 기준선 — 예전 Yjs baseline/stylesBaseline 맵과 같은 역할).
 */
export async function exportXlsxFromGrid(
  originalXlsxBytes: Uint8Array,
  canonicalGrid: GridDocument,
  baselineGrid: GridDocument,
): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalXlsxBytes as unknown as ArrayBuffer);
  const visibleSheets = workbook.worksheets.filter((sheet) => !sheet.state || sheet.state === 'visible');
  const originalSheetCount = visibleSheets.length;

  // ExcelJS는 시트 보호(sheetProtection)를 읽어오긴 하지만, 아무 것도 안 건드리고 그대로
  // writeBuffer()만 하면 그 정보를 재직렬화하지 않고 조용히 잃어버린다(ExcelJS 자체의 한계로 보임 —
  // protect()를 이번 세션에서 다시 명시적으로 불러야만 저장 시 살아남는다). 그래서 로드 직후
  // 상태를 기억해뒀다가 쓰기 직전에 다시 걸어준다 — 안 그러면 편집기 저장을 한 번이라도 거치고
  // 나면(사용자가 아무것도 안 건드려도 유휴 자동저장으로 매번 일어남) 잠금이 전부 풀려버린다.
  const wasProtected = visibleSheets.map((sheet) =>
    Boolean((sheet as unknown as { sheetProtection?: unknown }).sheetProtection),
  );

  for (let sheetIdx = 0; sheetIdx < canonicalGrid.sheets.length; sheetIdx += 1) {
    const target = canonicalGrid.sheets[sheetIdx];
    // 협업 중 새로 추가된 시트(webviewer 쪽 "+"/삽입 메뉴)는 원본 xlsx에 애초에 없어서
    // visibleSheets에 자리가 없다 — 지금 이름으로 새 워크시트를 만들어 그 자리에 채워 넣는다.
    let sheet = visibleSheets[sheetIdx];
    if (!sheet) {
      sheet = workbook.addWorksheet(target.name);
      visibleSheets[sheetIdx] = sheet;
    } else if (sheet.name !== target.name) {
      // 시트 이름 변경(우클릭 "이름 바꾸기")도 지금까지 원본 워크시트 객체를 그대로 재사용하면서
      // name만 안 옮겨 써서 저장 후 되돌아가 있었다 — 여기서 실제로 반영한다.
      sheet.name = target.name;
    }

    if (sheet.rowCount > target.rowCount) {
      sheet.spliceRows(target.rowCount + 1, sheet.rowCount - target.rowCount);
    }
    if (sheet.columnCount > target.colCount) {
      sheet.spliceColumns(target.colCount + 1, sheet.columnCount - target.colCount);
    }
  }

  // 반대로, 협업 중 삭제된 시트(원래 있었는데 지금 canonicalGrid.sheets엔 그 인덱스가 없는 것 —
  // 배열이라 삭제는 항상 끝에서부터만 반영된다는 뜻, v1 스코프와 동일한 단순화)는 원본에서 그대로
  // 남아 재저장되고 있었다 — 실제로 워크북에서 제거한다.
  for (let sheetIdx = canonicalGrid.sheets.length; sheetIdx < originalSheetCount; sheetIdx += 1) {
    const sheet = visibleSheets[sheetIdx];
    if (sheet) workbook.removeWorksheet(sheet.id);
  }

  for (let sheetIdx = 0; sheetIdx < canonicalGrid.sheets.length; sheetIdx += 1) {
    const currentSheet = canonicalGrid.sheets[sheetIdx];
    const baselineSheet = baselineGrid.sheets[sheetIdx];
    const sheet = visibleSheets[sheetIdx];
    if (!sheet) continue;

    for (let r = 0; r < currentSheet.rows.length; r += 1) {
      const currentRow = currentSheet.rows[r];
      const baselineRow = baselineSheet?.rows[r];
      for (let c = 0; c < currentRow.length; c += 1) {
        const currentCell = currentRow[c];
        const baselineCell = baselineRow?.[c];
        const currentText = cellText(currentCell);
        const baselineText = cellText(baselineCell);
        if (currentText === baselineText) continue; // 편집되지 않음 (수식 셀 포함) — 원본 그대로 둔다
        if (r >= currentSheet.rowCount || c >= currentSheet.colCount) continue;
        sheet.getCell(r + 1, c + 1).value = currentText === '' ? null : coerceCellValue(currentText);
      }
    }
  }

  for (let sheetIdx = 0; sheetIdx < canonicalGrid.sheets.length; sheetIdx += 1) {
    const currentSheet = canonicalGrid.sheets[sheetIdx];
    const baselineSheet = baselineGrid.sheets[sheetIdx];
    const sheet = visibleSheets[sheetIdx];
    if (!sheet) continue;

    for (let r = 0; r < currentSheet.rows.length; r += 1) {
      const currentRow = currentSheet.rows[r];
      const baselineRow = baselineSheet?.rows[r];
      for (let c = 0; c < currentRow.length; c += 1) {
        const currentCell = currentRow[c];
        const baselineCell = baselineRow?.[c];
        const currentStyle = cellStyle(currentCell, canonicalGrid);
        const baselineStyle = cellStyle(baselineCell, baselineGrid);
        if (currentStyle === baselineStyle) continue;
        if (currentStyle && baselineStyle && JSON.stringify(currentStyle) === JSON.stringify(baselineStyle)) continue;
        if (r >= currentSheet.rowCount || c >= currentSheet.colCount) continue;
        applyStyleToCell(sheet.getCell(r + 1, c + 1), currentStyle ? toCellStyleAttrs(currentStyle) : undefined);
      }
    }
  }

  for (let i = 0; i < visibleSheets.length; i += 1) {
    if (wasProtected[i]) {
      await visibleSheets[i].protect('', {});
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/** 저장 직후 뷰어용 그리드(읽기전용 열람 캐시)를 새 xlsx 바이트에서 다시 뽑아낸다. */
export function reparseGridForViewer(newXlsxBytes: Uint8Array): Promise<GridDocument> {
  return parseXlsxToGrid(Buffer.from(newXlsxBytes));
}
