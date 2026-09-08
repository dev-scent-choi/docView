import ExcelJS from 'exceljs';
import * as Y from 'yjs';
import {
  GRID_BASELINE_MAP,
  GRID_CELLS_MAP,
  GRID_COMPUTED_MAP,
  GRID_MERGES_MAP,
  GRID_STRUCTURE_MAP,
  GRID_STYLES_BASELINE_MAP,
  GRID_STYLES_MAP,
  gridCellKey,
  parseGridCellKey,
  PAGE_SIZE_ROWS,
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

/**
 * toCellStyleAttrs의 역방향 — 클라이언트가 GRID_STYLES_MAP에 쓴 서식(CellStyleAttrs)을 canonical
 * 그리드의 공용 스타일 배열에 담을 모양(GridStyle)으로 바꾼다.
 *
 * 두 타입은 사실상 같은 서식을 서로 다른 층(Yjs 저장용 / canonical 저장용)에서 표현한 것이고,
 * 필드 이름도 하나(`wrap` ↔ `w`)만 빼고 전부 같다 — 그 한 곳이 바로 이 변환이 필요한 이유다.
 * 키를 항상 **같은 순서로** 넣는 것도 중요하다: findOrAddGridStyleIndex가 JSON 문자열 비교로
 * 동일한 서식을 찾기 때문에, 같은 서식이 키 순서만 달라 서로 다른 항목으로 중복 등록되면 스타일
 * 배열이 편집할 때마다 계속 늘어난다.
 */
export function fromCellStyleAttrs(attrs: CellStyleAttrs): GridStyle {
  const style: { -readonly [K in keyof GridStyle]?: GridStyle[K] } = {};
  if (attrs.b) style.b = 1;
  if (attrs.i) style.i = 1;
  if (attrs.u) style.u = 1;
  if (attrs.st) style.st = 1;
  if (attrs.fs) style.fs = attrs.fs;
  if (attrs.ff) style.ff = attrs.ff;
  if (attrs.col) style.col = attrs.col;
  if (attrs.bg) style.bg = attrs.bg;
  if (attrs.a) style.a = attrs.a;
  if (attrs.va) style.va = attrs.va;
  if (attrs.wrap) style.w = 1;
  if (attrs.nf) style.nf = attrs.nf;
  if (attrs.nd) style.nd = attrs.nd;
  if (attrs.lk) style.lk = 1;
  if (attrs.nu) style.nu = 1;
  if (attrs.bt) style.bt = attrs.bt;
  if (attrs.br) style.br = attrs.br;
  if (attrs.bb) style.bb = attrs.bb;
  if (attrs.bl) style.bl = attrs.bl;
  return style;
}

/** 두 GridStyle이 "같은 서식"인지 판정하기 위한 정규화 문자열 — 키 순서에 좌우되지 않도록 정렬한다
 * (parseXlsxToGrid의 StylePalette가 만든 기존 항목과 우리가 방금 만든 항목의 키 순서가 다를 수 있다). */
function gridStyleSignature(style: GridStyle): string {
  const entries = Object.entries(style as Record<string, unknown>).filter(([, v]) => v !== undefined);
  entries.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
  return JSON.stringify(entries);
}

/**
 * canonical 그리드의 공용 스타일 배열에서 이 서식과 **구조적으로 같은** 항목을 찾아 그 인덱스를
 * 돌려주고, 없으면 새로 추가한 뒤 그 인덱스를 돌려준다(셀은 서식 객체가 아니라 이 인덱스만 들고
 * 있다 — GridCell.s).
 *
 * 선형 탐색으로 충분하다: 이 배열은 "워크북이 실제로 쓰는 서로 다른 서식의 팔레트"라 보통 수십
 * 개 수준이고, 여기 오는 호출은 사용자가 서식 버튼을 누르는 빈도(초당 몇 번)에 불과하다.
 */
export function findOrAddGridStyleIndex(grid: GridDocument, style: GridStyle): number {
  const signature = gridStyleSignature(style);
  for (let i = 0; i < grid.styles.length; i += 1) {
    if (gridStyleSignature(grid.styles[i]) === signature) return i;
  }
  (grid.styles as GridStyle[]).push(style);
  return grid.styles.length - 1;
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

/**
 * 윈도잉이 만드는 Yjs 트랜잭션에 붙이는 origin 표식 — "이 변경은 사용자 편집이 아니라 서버가
 * 창(window)을 채우거나 비운 것"이라는 뜻이다.
 *
 * **이게 없으면 조용한 데이터 손상이 난다.** room.ts는 `cells.observe(...)`로 GRID_CELLS_MAP의
 * 모든 변경을 canonical 그리드에 되반영하는데(Task 2), 그 관찰자는 "키 삭제 = 사용자가 그 칸을
 * 지웠다"로 해석해 canonical을 0(빈 칸)으로 되돌린다. 페이지를 비우는 것(unseed)도 똑같이 키
 * 삭제이므로, 표식이 없으면 "스크롤해서 화면에서 사라진 500행"이 canonical에서 통째로 지워지고
 * 그 상태가 그대로 파일에 저장된다 — 사용자가 아무것도 지우지 않았는데 데이터가 날아간다.
 * 반대로 페이지를 채우는 것(seed)은 seedableCellValue로 서식이 벗겨진 편집용 텍스트를 쓰므로
 * (예: "1,000" → "1000"), 표식이 없으면 canonical의 올바른 표시 텍스트를 그 벗겨진 값으로
 * 덮어쓴다(seedRoom이 최초 시딩을 관찰 시작 "전"에 하는 것과 같은 이유의 문제 — 윈도잉에서는
 * 시딩이 관찰이 걸린 뒤에도 계속 일어나므로 시점으로는 피할 수 없다).
 *
 * 플래그 변수(`room.suppressCanonicalSync = true` 같은 것) 대신 트랜잭션 origin을 쓰는 이유:
 * 이 함수들은 청크 사이에 `await yieldToEventLoop()`으로 양보하므로, 그 틈에 도착한 **진짜**
 * 사용자 편집도 플래그가 켜져 있는 동안 처리되어 canonical 반영을 놓친다(자동저장 억제 플래그를
 * 함수 전체에 두르면 안 되는 것과 똑같은 클래스의 버그, Room.suppressAutosaveReschedule 주석
 * 참고). origin은 트랜잭션마다 개별적으로 붙으므로 그런 시간 창이 애초에 존재하지 않는다.
 */
export const GRID_WINDOW_SEED_ORIGIN = 'window-seed';

/**
 * 수식 계산값(GRID_COMPUTED_MAP) 쓰기에 붙이는 origin 표식 — GRID_WINDOW_SEED_ORIGIN과 같은
 * 취지("이건 사용자 편집이 아니라 서버가 만든 파생 캐시이므로 자동저장을 예약할 이유가 없다")지만
 * 다른 종류의 쓰기라 상수를 따로 둔다.
 *
 * **왜 자동저장을 예약하면 안 되는가**: 저장(exportXlsxFromGrid)은 canonicalGrid/baselineGrid만
 * 읽고 이 맵은 단 한 번도 보지 않는다 — 즉 이 맵에 무엇이 쓰여도 디스크에 나갈 바이트는 한 글자도
 * 달라지지 않는다. 그런데 표식이 없으면 **편집 없이 스크롤만 해도**(새 페이지 구독 → 그 페이지
 * 재계산 → computed 쓰기 → 'update') 자동저장이 걸려, 문서를 읽기만 하는 사용자가 8초 이상 멈출
 * 때마다 전체 xlsx export + 재파싱 + 뷰어 캐시 재생성이 한 번씩 돈다(flushToStore에는 dirty 검사가
 * 없다). 디바운스로 합쳐지긴 하지만 순수한 낭비다.
 *
 * broadcast는 그대로 해야 한다 — 그 페이지를 보고 있는 다른 클라이언트가 갱신된 계산값을 받는
 * 경로가 바로 그 broadcast다. 그래서 room.ts의 'update' 핸들러는 전파는 origin과 무관하게 하고
 * scheduleAutosave만 origin으로 건너뛴다(GRID_WINDOW_SEED_ORIGIN과 완전히 동일한 처리).
 */
export const GRID_COMPUTED_WRITE_ORIGIN = 'computed-write';

/**
 * 이 origin으로 태깅된 트랜잭션은 자동저장을 예약하지 않는다 — 둘 다 "서버가 만든 캐시/창 조작이라
 * 저장 대상(canonicalGrid)을 바꾸지 않는다"는 같은 이유다. room.ts의 `ydoc.on('update')`가 이
 * 판정을 쓴다(전파는 origin과 무관하게 항상 한다).
 */
export function isAutosaveExemptOrigin(origin: unknown): boolean {
  return origin === GRID_WINDOW_SEED_ORIGIN || origin === GRID_COMPUTED_WRITE_ORIGIN;
}

/**
 * 시트 하나의 페이지 하나(PAGE_SIZE_ROWS행)만 Yjs에 채운다 — seedGridFromXlsx의 내부 행 루프와
 * 같은 로직을 한 페이지 범위로 좁힌 것.
 *
 * 데이터 출처는 원본 파일이 아니라 **canonicalGrid**다: 사용자가 편집한 뒤 스크롤로 화면에서
 * 밀려났다가(unseed) 다시 돌아온 페이지도 편집된 값 그대로 되살아나야 한다(원본에서 다시 읽으면
 * 그 편집이 조용히 되돌려진다).
 *
 * **서식(GRID_STYLES_MAP/GRID_STYLES_BASELINE_MAP)은 여기서 채우지 않는다** — 서식은 구조
 * 메타데이터와 마찬가지로 절대 윈도잉하지 않고 룸을 열 때 전량 한 번만 채운다
 * (seedGridStylesFromXlsx 주석 참고).
 */
export async function seedGridPageFromXlsx(
  ydoc: Y.Doc,
  canonicalGrid: GridDocument,
  sheetIdx: number,
  page: number,
): Promise<void> {
  const sheet = canonicalGrid.sheets[sheetIdx];
  if (!sheet) return;
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const baseline = ydoc.getMap<string>(GRID_BASELINE_MAP);
  const covered = mergeCoveredCells(sheet.merges);
  const pageStart = page * PAGE_SIZE_ROWS;
  const pageEnd = Math.min(pageStart + PAGE_SIZE_ROWS, sheet.rows.length);

  // 페이지(500행)가 지금은 YIELD_EVERY_N_ROWS(2000행)보다 작아 실제로는 청크가 하나뿐이지만,
  // 두 상수 중 어느 쪽이 나중에 바뀌어도 "큰 시트가 이벤트 루프를 막지 않는다"는 보장이 유지되도록
  // 다른 대형 루프와 같은 청크+양보 구조를 그대로 쓴다(spreadsheet.ts yieldToEventLoop 주석 참고).
  for (let start = pageStart; start < pageEnd; start += YIELD_EVERY_N_ROWS) {
    const end = Math.min(start + YIELD_EVERY_N_ROWS, pageEnd);
    ydoc.transact(() => {
      for (let r = start; r < end; r += 1) {
        const row = sheet.rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c += 1) {
          const cell = row[c];
          if (cell === 0) continue;
          if (covered.has(`${r}:${c}`)) continue; // 병합에 덮인 칸 — 앵커의 중복 값이라 시딩하지 않는다
          const key = gridCellKey(sheetIdx, r, c);
          const style = cell.s !== undefined ? canonicalGrid.styles[cell.s] : undefined;
          if (cell.v !== '') {
            const value = seedableCellValue(cell.v, style);
            cells.set(key, value);
            // 이 baseline 맵은 이제 저장 기준선이 아니다 — 저장은 room.baselineGrid(룸 생성 시점의
            // 깊은 복사본)를 쓴다(Task 3). 그래서 편집된 셀이 재시딩될 때 여기에 "편집된 값"이
            // 들어가도 저장 결과에는 영향이 없다. 클라이언트가 참고할 수 있도록 채우는 것뿐.
            baseline.set(key, value);
          }
          // 서식은 여기서 쓰지 않는다 — 룸을 열 때 seedGridStylesFromXlsx가 전량 채워 두고 그
          // 뒤로는 절대 지우지 않으므로, 페이지가 (재)구독될 때 다시 쓸 것이 없다. 여기서 다시
          // 쓰면 사용자가 그 페이지를 보는 동안 적용한 서식이 원본 파일의 서식으로 되돌아간다.
        }
      }
    }, GRID_WINDOW_SEED_ORIGIN);
    await yieldToEventLoop();
  }
}

/**
 * 워크북 전체의 셀 서식을 한 번에 Yjs에 채운다 — 룸을 열자마자(seedGridStructureOnly 바로 뒤)
 * 한 번만 호출하고, 그 뒤로는 어떤 unseed 경로에서도 **절대 지우지 않는다**.
 *
 * **왜 서식은 윈도잉하지 않는가**: 클라이언트는 자기 UI 동작(굵게·색·표시 형식)으로
 * GRID_STYLES_MAP에 직접 쓰는데, 서버에는 그 서식 편집을 canonicalGrid로 되반영하는 관찰자가
 * 없다(값 편집만 cells.observe가 처리한다). 그래서 서식을 셀 값처럼 윈도잉하면 —
 * 페이지가 유예 시간 뒤 비워지고(서식 키 삭제) 나중에 다시 구독되어 canonical 기준으로 다시
 * 채워지는 순간 — 사용자가 그 페이지에서 적용해 둔 서식이 조용히 원본 서식으로 되돌아간다
 * (스크롤만 했는데 굵게가 풀리는 회귀). 구조 메타데이터(GRID_STRUCTURE_MAP/GRID_MERGES_MAP)와
 * 똑같이 "절대 윈도잉하지 않는 데이터"로 취급하면 애초에 지워지지 않으므로 되돌아갈 일이 없다.
 *
 * 전량을 한 번에 채워도 괜찮은 이유: 서식이 걸린 셀은 보통 전체 셀 수보다 훨씬 희소하다(대개
 * 머리글 행·합계 열 정도). 그래도 만약을 위해 다른 대형 루프와 같은 YIELD_EVERY_N_ROWS 청크 +
 * yieldToEventLoop() 구조를 그대로 쓴다 — 이 코드베이스에서 시트 행을 도는 모든 루프의 관례다.
 */
export async function seedGridStylesFromXlsx(ydoc: Y.Doc, grid: GridDocument): Promise<void> {
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const stylesBaseline = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_BASELINE_MAP);

  for (let sheetIdx = 0; sheetIdx < grid.sheets.length; sheetIdx += 1) {
    const sheet = grid.sheets[sheetIdx];
    const covered = mergeCoveredCells(sheet.merges);

    for (let start = 0; start < sheet.rows.length; start += YIELD_EVERY_N_ROWS) {
      const end = Math.min(start + YIELD_EVERY_N_ROWS, sheet.rows.length);
      ydoc.transact(() => {
        for (let r = start; r < end; r += 1) {
          const row = sheet.rows[r];
          if (!row) continue;
          for (let c = 0; c < row.length; c += 1) {
            const cell = row[c];
            if (cell === 0) continue;
            if (covered.has(`${r}:${c}`)) continue; // 병합에 덮인 칸 — 앵커의 중복이라 시딩하지 않는다
            const style = cell.s !== undefined ? grid.styles[cell.s] : undefined;
            if (!style) continue;
            const attrs = toCellStyleAttrs(style);
            if (Object.keys(attrs).length === 0) continue;
            const key = gridCellKey(sheetIdx, r, c);
            styles.set(key, attrs);
            stylesBaseline.set(key, attrs);
          }
        }
      }, GRID_WINDOW_SEED_ORIGIN);
      await yieldToEventLoop();
    }
  }
}

/**
 * 페이지 하나를 Yjs에서 지운다 — canonical 그리드에는 아무 영향도 없다(그쪽이 항상 진실이고,
 * 저장·다운로드·수식 계산은 모두 canonical만 본다). 지워진 키는 그 페이지를 아무도 보고 있지
 * 않다는 뜻이므로, 클라이언트 쪽 메모리도 같은 sync 프로토콜을 타고 함께 줄어든다.
 *
 * 행 범위로 좁히기 위해 맵 전체 키를 훑는다 — 윈도잉 이후 이 맵에는 "지금 구독 중인 페이지"의
 * 키만 들어 있어(그게 이 태스크의 요점) 훑는 양 자체가 작다.
 */
export async function unseedGridPageFromXlsx(
  ydoc: Y.Doc,
  sheetIdx: number,
  page: number,
  rowCount: number,
): Promise<void> {
  // 서식 맵(GRID_STYLES_MAP/GRID_STYLES_BASELINE_MAP)은 **일부러 빠져 있다** — 서식은 윈도잉
  // 대상이 아니고(seedGridStylesFromXlsx가 룸을 열 때 전량 채운다) 서버에는 클라이언트의 서식
  // 편집을 canonical로 되반영하는 경로가 없다. 여기서 지우면 나중에 그 페이지를 다시 구독할 때
  // 원본 서식으로 다시 채워져, 사용자가 적용해 둔 서식이 스크롤만으로 조용히 되돌아간다.
  const maps = [
    ydoc.getMap<string>(GRID_CELLS_MAP),
    ydoc.getMap<string>(GRID_BASELINE_MAP),
    ydoc.getMap<string>(GRID_COMPUTED_MAP),
  ] as ReadonlyArray<Y.Map<unknown>>;
  const pageStart = page * PAGE_SIZE_ROWS;
  const pageEnd = Math.min(pageStart + PAGE_SIZE_ROWS, rowCount);

  for (let start = pageStart; start < pageEnd; start += YIELD_EVERY_N_ROWS) {
    const end = Math.min(start + YIELD_EVERY_N_ROWS, pageEnd);
    ydoc.transact(() => {
      for (const map of maps) {
        for (const key of [...map.keys()]) {
          const parsed = parseGridCellKey(key);
          if (parsed && parsed.sheetIdx === sheetIdx && parsed.row >= start && parsed.row < end) {
            map.delete(key);
          }
        }
      }
    }, GRID_WINDOW_SEED_ORIGIN);
    await yieldToEventLoop();
  }
}

/**
 * 셀 데이터 없이 구조 메타데이터(시트 이름·행렬 수·틀고정·열 너비·병합)만 채운다 — 룸을 열자마자
 * 호출한다. 클라이언트는 아직 로드하지 않은 페이지의 스크롤 높이와 병합 경계를 미리 알아야
 * 가상 스크롤바를 올바른 길이로 그릴 수 있으므로, 이 두 맵은 **절대 윈도잉하지 않는다**(계획의
 * Global Constraints) — 구독 상태와 무관하게 모든 시트 분량을 항상 즉시 채운다.
 */
export async function seedGridStructureOnly(ydoc: Y.Doc, grid: GridDocument): Promise<void> {
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
    }, GRID_WINDOW_SEED_ORIGIN);
    await yieldToEventLoop();
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
 * 값·서식 배열만으로는 표현할 수 없는 구조 변경을 exportXlsxFromGrid에 전달하는 통로.
 *
 * **왜 canonicalGrid 안에 담지 않는가**: canonicalGrid.sheets의 **인덱스가 곧 시트 식별자**다
 * (Yjs cells 키 "시트:행:열"의 시트 번호, 수식 계산의 sheetIdx가 전부 이 번호를 쓴다). 그래서
 * 시트를 지웠다고 배열에서 splice해 버리면 그 뒤 모든 시트의 번호가 밀려 셀 데이터가 통째로 다른
 * 시트를 가리키게 된다 — 삭제는 배열에서 빼는 대신 "이 자리는 지워졌다"를 따로 알리는 방식으로만
 * 안전하게 표현할 수 있다. 탭 순서도 같은 이유로 배열 순서와 분리해 둔다.
 */
export interface XlsxExportOptions {
  /** 협업 중 삭제된 시트의 자리 번호(canonicalGrid.sheets의 인덱스). */
  readonly removedSheetIndices?: ReadonlySet<number>;
  /** 시트 자리 번호 → 화면(탭) 순서. 주지 않으면 원본 파일의 순서를 그대로 둔다. */
  readonly sheetDisplayOrder?: ReadonlyMap<number, number>;
}

const EMPTY_REMOVED_SHEETS: ReadonlySet<number> = new Set<number>();

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
  options: XlsxExportOptions = {},
): Promise<Uint8Array> {
  const removedSheets = options.removedSheetIndices ?? EMPTY_REMOVED_SHEETS;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalXlsxBytes as unknown as ArrayBuffer);
  const visibleSheets = workbook.worksheets.filter((sheet) => !sheet.state || sheet.state === 'visible');
  const originalSheetCount = visibleSheets.length;

  // 협업 중 삭제된 시트를 **가장 먼저** 워크북에서 제거한다 — 이름 바꾸기/시트 추가가 방금 지운
  // 시트의 이름을 그대로 재사용할 수 있기 때문이다(먼저 지워두지 않으면 이름이 충돌한다).
  // visibleSheets 배열 자체는 그대로 둔다: 이 배열의 인덱스가 곧 "원본에서의 시트 자리 번호"이고,
  // canonicalGrid.sheets의 인덱스(= Yjs cells 키의 시트 번호)와 1:1로 대응하는 유일한 연결고리다.
  for (const sheetIdx of removedSheets) {
    const sheet = visibleSheets[sheetIdx];
    if (sheet) workbook.removeWorksheet(sheet.id);
  }

  // ExcelJS는 시트 보호(sheetProtection)를 읽어오긴 하지만, 아무 것도 안 건드리고 그대로
  // writeBuffer()만 하면 그 정보를 재직렬화하지 않고 조용히 잃어버린다(ExcelJS 자체의 한계로 보임 —
  // protect()를 이번 세션에서 다시 명시적으로 불러야만 저장 시 살아남는다). 그래서 로드 직후
  // 상태를 기억해뒀다가 쓰기 직전에 다시 걸어준다 — 안 그러면 편집기 저장을 한 번이라도 거치고
  // 나면(사용자가 아무것도 안 건드려도 유휴 자동저장으로 매번 일어남) 잠금이 전부 풀려버린다.
  const wasProtected = visibleSheets.map((sheet) =>
    Boolean((sheet as unknown as { sheetProtection?: unknown }).sheetProtection),
  );

  for (let sheetIdx = 0; sheetIdx < canonicalGrid.sheets.length; sheetIdx += 1) {
    if (removedSheets.has(sheetIdx)) continue;
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

  // 값 diff 패스 — 이 아래 두 전체 그리드 스캔은 최대 시트 크기에서 각각 초 단위로 이어질 수
  // 있고(실측 ~2초), 자동저장과 다운로드 **매번** 돈다. 다른 대형 루프와 똑같이
  // YIELD_EVERY_N_ROWS행마다 실제 이벤트 루프에 양보한다(계획의 Global Constraint #1).
  for (let sheetIdx = 0; sheetIdx < canonicalGrid.sheets.length; sheetIdx += 1) {
    if (removedSheets.has(sheetIdx)) continue;
    const currentSheet = canonicalGrid.sheets[sheetIdx];
    const baselineSheet = baselineGrid.sheets[sheetIdx];
    const sheet = visibleSheets[sheetIdx];
    if (!sheet) continue;

    for (let r = 0; r < currentSheet.rows.length; r += 1) {
      const currentRow = currentSheet.rows[r];
      const baselineRow = baselineSheet?.rows[r];
      // 열 상한은 행 배열 길이가 아니라 시트의 colCount다 — 편집으로 열이 늘어난 시트에서는
      // (원본 시절 만들어진) 기존 행 배열이 새 colCount보다 짧을 수 있고, 그 새 열에 들어온 값이
      // 행 배열 길이 기준으로 돌면 통째로 빠진다. cellText는 없는 칸(undefined)을 빈 칸으로 본다.
      for (let c = 0; c < currentSheet.colCount; c += 1) {
        const currentCell = currentRow?.[c];
        const baselineCell = baselineRow?.[c];
        const currentText = cellText(currentCell);
        const baselineText = cellText(baselineCell);
        if (currentText === baselineText) continue; // 편집되지 않음 (수식 셀 포함) — 원본 그대로 둔다
        if (r >= currentSheet.rowCount) continue;
        sheet.getCell(r + 1, c + 1).value = currentText === '' ? null : coerceCellValue(currentText);
      }
      if ((r + 1) % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
    }
  }

  // 서식 diff 패스 — 위와 같은 이유로 같은 양보 구조를 쓴다.
  for (let sheetIdx = 0; sheetIdx < canonicalGrid.sheets.length; sheetIdx += 1) {
    if (removedSheets.has(sheetIdx)) continue;
    const currentSheet = canonicalGrid.sheets[sheetIdx];
    const baselineSheet = baselineGrid.sheets[sheetIdx];
    const sheet = visibleSheets[sheetIdx];
    if (!sheet) continue;

    for (let r = 0; r < currentSheet.rows.length; r += 1) {
      const currentRow = currentSheet.rows[r];
      const baselineRow = baselineSheet?.rows[r];
      for (let c = 0; c < currentSheet.colCount; c += 1) {
        const currentCell = currentRow?.[c];
        const baselineCell = baselineRow?.[c];
        const currentStyle = cellStyle(currentCell, canonicalGrid);
        const baselineStyle = cellStyle(baselineCell, baselineGrid);
        if (currentStyle === baselineStyle) continue;
        if (currentStyle && baselineStyle && JSON.stringify(currentStyle) === JSON.stringify(baselineStyle)) continue;
        if (r >= currentSheet.rowCount) continue;
        applyStyleToCell(sheet.getCell(r + 1, c + 1), currentStyle ? toCellStyleAttrs(currentStyle) : undefined);
      }
      if ((r + 1) % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
    }
  }

  for (let i = 0; i < visibleSheets.length; i += 1) {
    if (removedSheets.has(i)) continue; // 이미 워크북에서 제거된 시트 — 다시 잠글 대상이 아니다
    if (wasProtected[i]) {
      await visibleSheets[i].protect('', {});
    }
  }

  applySheetDisplayOrder(workbook, visibleSheets, canonicalGrid, removedSheets, options.sheetDisplayOrder);

  const buffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/**
 * 시트 탭 순서 변경(클라이언트의 SheetStructure.order)을 실제 파일의 시트 순서로 옮긴다.
 *
 * ExcelJS는 `workbook.worksheets` 게터에서 `orderNo`로 정렬하고 저장 모델도 그 게터를 쓰므로
 * (node_modules/exceljs/lib/doc/workbook.js), orderNo만 다시 매기면 저장 순서가 바뀐다. 이때
 * "지금 존재하는 orderNo 값들"을 그대로 재사용해 자리만 서로 바꾼다 — 새 숫자를 임의로 부여하면
 * 숨겨진 시트(visibleSheets에 없어 우리가 순서를 모르는 시트)와의 상대 위치가 흐트러진다.
 */
function applySheetDisplayOrder(
  workbook: ExcelJS.Workbook,
  visibleSheets: readonly ExcelJS.Worksheet[],
  canonicalGrid: GridDocument,
  removedSheets: ReadonlySet<number>,
  displayOrder: ReadonlyMap<number, number> | undefined,
): void {
  if (!displayOrder || displayOrder.size === 0) return;
  const survivors: Array<{ sheetIdx: number; worksheet: ExcelJS.Worksheet }> = [];
  for (let sheetIdx = 0; sheetIdx < canonicalGrid.sheets.length; sheetIdx += 1) {
    if (removedSheets.has(sheetIdx)) continue;
    const worksheet = visibleSheets[sheetIdx];
    if (worksheet && workbook.getWorksheet(worksheet.id)) survivors.push({ sheetIdx, worksheet });
  }
  if (survivors.length < 2) return;

  const slots = survivors
    .map(({ worksheet }) => (worksheet as unknown as { orderNo: number }).orderNo)
    .sort((a, b) => a - b);
  const ordered = [...survivors].sort(
    (a, b) => (displayOrder.get(a.sheetIdx) ?? a.sheetIdx) - (displayOrder.get(b.sheetIdx) ?? b.sheetIdx),
  );
  ordered.forEach(({ worksheet }, i) => {
    (worksheet as unknown as { orderNo: number }).orderNo = slots[i];
  });
}

/** 저장 직후 뷰어용 그리드(읽기전용 열람 캐시)를 새 xlsx 바이트에서 다시 뽑아낸다. */
export function reparseGridForViewer(newXlsxBytes: Uint8Array): Promise<GridDocument> {
  return parseXlsxToGrid(Buffer.from(newXlsxBytes));
}
