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
  parseGridCellKey,
  PAGE_SIZE_ROWS,
  WINDOWED_GRID_CELL_MAPS,
  type CellStyleAttrs,
  type GridMergeShape,
  type GridStructureOpRecord,
  type SheetStructure,
} from '../collab/gridShared.js';
import type { GridCell, GridDocument, GridStyle } from '../spreadsheet.js';
import { parseXlsxToGrid, yieldToEventLoop, YIELD_EVERY_N_ROWS } from '../spreadsheet.js';
import type { LazyGridSheetMeta, LazyGridStore } from './lazyGridStore.js';
import { getRow } from './lazyGridStore.js';


export function toCellStyleAttrs(style: GridStyle): CellStyleAttrs {
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
export function findOrAddGridStyleIndex(store: LazyGridStore, style: GridStyle): number {
  const signature = gridStyleSignature(style);
  for (let i = 0; i < store.styles.length; i += 1) {
    if (gridStyleSignature(store.styles[i]) === signature) return i;
  }
  store.styles.push(style);
  return store.styles.length - 1;
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
 * **왜 자동저장을 예약하면 안 되는가**: 저장(exportXlsxFromGrid)은 canonicalGrid/dirtyCells/
 * structuralOps만 읽고 이 맵은 단 한 번도 보지 않는다 — 즉 이 맵에 무엇이 쓰여도 디스크에 나갈 바이트는 한 글자도
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
 * LazyGridStore의 원본 행(getRow가 돌려주는, **원본 열 번호로 인덱싱된** 배열)에서 논리 좌표
 * (r, c)의 셀을 뽑는다 — lazyGridStore.ts의 getCell과 완전히 같은 규칙(더티 오버레이 우선 →
 * 논리 열을 원본 열로 변환해 인덱싱)이지만, 이미 getRow로 받아 둔 행을 재사용해 셀마다 다시
 * getRow를 호출하지 않는다(호출부가 청크 단위로 Promise.all 배치 읽기를 한 뒤 이 함수로 각 셀을
 * 뽑아 쓰는 구조 — Task 8 "구현 순서 메모" 참고).
 *
 * 이 조회가 dirtyOverlay를 확인하는 이유: seedGridPageFromXlsx는 나중에
 * (Task 11+) 실제 편집 중인 LazyGridStore에도 그대로 쓰이게 된다 — 편집된 칸은 디스크의 원본이
 * 아니라 dirtyOverlay에만 최신 값이 있으므로, 이걸 빼먹으면 창을 벗어났다 되돌아온 페이지에서
 * 사용자의 편집이 조용히 원본 값으로 되돌아간다(GRID_WINDOW_SEED_ORIGIN 주석이 경고하는 것과
 * 같은 클래스의 데이터 손실).
 */
function resolveGridCell(
  store: LazyGridStore,
  sheetIdx: number,
  row: number,
  col: number,
  originalRow: ReadonlyArray<GridCell | 0>,
): GridCell | 0 {
  const key = gridCellKey(sheetIdx, row, col);
  if (store.dirtyOverlay.has(key)) return store.dirtyOverlay.get(key) ?? 0;
  const originalCol = store.logicalToOriginalCol[sheetIdx]?.[col];
  if (originalCol === undefined || originalCol < 0) return 0;
  return originalRow[originalCol] ?? 0;
}

/**
 * 시트 하나의 페이지 하나(PAGE_SIZE_ROWS행)만 Yjs에 채운다 — seedGridFromXlsx의 내부 행 루프와
 * 같은 로직을 한 페이지 범위로 좁힌 것.
 *
 * 데이터 출처는 원본 파일이 아니라 **LazyGridStore**다(Task 3-7이 만든 지연 읽기 계층, Task 11부터
 * room.ts가 진짜 상주 인스턴스(`room.lazyGrid`)를 직접 넘긴다): 사용자가 편집한 뒤 스크롤로 화면에서 밀려났다가(unseed) 다시
 * 돌아온 페이지도 편집된 값 그대로 되살아나야 한다(원본에서 다시 읽으면 그 편집이 조용히
 * 되돌려진다) — resolveGridCell의 dirtyOverlay 확인이 그걸 보장한다.
 *
 * **서식(GRID_STYLES_MAP/GRID_STYLES_BASELINE_MAP)도 이제 여기서 콘텐츠와 함께 채운다.**
 * 예전에는 룸을 열 때 `seedGridStylesFromXlsx`가 워크북 전체(모든 시트 × 모든 행)를 한 번에 훑어
 * `getRow`를 호출해 서식을 전량 미리 채웠고, `joinEditSession`은 그 완료를 `room.ready`로
 * 기다렸다 — 실측 결과 이 전량 스캔이 룸을 여는 시간의 대부분을 차지했다(2만 행 파일에서 약 80%,
 * 5.9초 / 6만 행에서 약 81%, 18.4초 — 콘텐츠 인제스트나 수식 재계산보다 훨씬 큰 비용). 콘텐츠
 * (cells/baseline)는 이미 Task 11부터 "지금 구독된 페이지만" 채우는데, 서식만 룸이 열리자마자
 * 시트 전체를 선점적으로 읽는 게 비대칭이었고 그게 바로 이 비용의 정체였다.
 *
 * 고치는 방법은 콘텐츠가 이미 쓰는 패턴을 그대로 따르는 것이다: 이 함수는 어차피 페이지의 각 셀을
 * `resolveGridCell`로 읽으면서 `seedableCellValue`(쉼표 서식 판정) 때문에 그 셀의 style을 이미
 * 구해 두고 있었다 — 그 style을 그대로 GRID_STYLES_MAP에도 같이 쓰기만 하면, 서식도 콘텐츠와
 * **정확히 같은 시점에, 정확히 같은 창(구독된 페이지)만** 채워진다. 추가 디스크 읽기도, 추가
 * 청크 루프도 없다.
 *
 * 서식을 윈도잉해도 안전한 이유(예전엔 "서식은 절대 윈도잉하지 않는다"가 규칙이었다 — 아래 두
 * 가지가 그 규칙을 무너뜨리지 않고 대체한다):
 * 1) **화면 표시**: 페이지가 비워졌다(unseed) 다시 채워질 때, `resolveGridCell`은 항상
 *    `dirtyOverlay`를 먼저 확인한다. 사용자가 그 페이지에서 적용한 서식 편집은 room.ts의
 *    `styles.observe` → `applyStylesChangeToCanonical`이 이미 `setDirtyCell`로 `dirtyOverlay`에
 *    반영해 두므로, 재시딩은 편집된 서식을 그대로 복원한다 — "스크롤만 했는데 서식이 원본으로
 *    되돌아가는" 회귀는 콘텐츠와 마찬가지로 일어나지 않는다.
 * 2) **저장(export)**: `exportXlsxFromGrid`는 애초에 Yjs의 GRID_STYLES_MAP을 단 한 번도 읽지
 *    않는다 — `store.dirtyOverlay`(편집된 칸)와 원본 xlsx 바이트(편집 안 된 칸)만 본다. 즉 서식이
 *    Yjs에 아직 없거나(구독 전) 창 밖이라 비워져 있어도(unseed) 저장 결과에는 원래부터 아무
 *    영향이 없다 — 이 사실이 서식을 안전하게 윈도잉할 수 있는 진짜 근거다. GRID_STYLES_MAP은
 *    (GRID_BASELINE_MAP처럼) 저장 기준선이 아니라 순전히 "지금 그 페이지를 보는 클라이언트의
 *    화면 표시용" 캐시일 뿐이다.
 */
export async function seedGridPageFromXlsx(
  ydoc: Y.Doc,
  store: LazyGridStore,
  sheetIdx: number,
  page: number,
): Promise<void> {
  const meta = store.sheetMeta[sheetIdx];
  if (!meta) return;
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const baseline = ydoc.getMap<string>(GRID_BASELINE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const stylesBaseline = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_BASELINE_MAP);
  const covered = mergeCoveredCells(meta.merges);
  const pageStart = page * PAGE_SIZE_ROWS;
  const pageEnd = Math.min(pageStart + PAGE_SIZE_ROWS, meta.rowCount);

  // 페이지(500행)가 지금은 YIELD_EVERY_N_ROWS(2000행)보다 작아 실제로는 청크가 하나뿐이지만,
  // 두 상수 중 어느 쪽이 나중에 바뀌어도 "큰 시트가 이벤트 루프를 막지 않는다"는 보장이 유지되도록
  // 다른 대형 루프와 같은 청크+양보 구조를 그대로 쓴다(spreadsheet.ts yieldToEventLoop 주석 참고).
  for (let start = pageStart; start < pageEnd; start += YIELD_EVERY_N_ROWS) {
    const end = Math.min(start + YIELD_EVERY_N_ROWS, pageEnd);
    const rowNumbers: number[] = [];
    for (let r = start; r < end; r += 1) rowNumbers.push(r);
    // 청크 안의 행들을 한 번에 기다린다 — 이미 캐시가 데워진 행은 즉시 resolve되고, 캐시에 없던
    // 행만 실제 디스크 I/O가 나가되 그 I/O들이 병렬화된다(Task 8 "구현 순서 메모" 참고).
    const rows = await Promise.all(rowNumbers.map((r) => getRow(store, sheetIdx, r)));
    ydoc.transact(() => {
      for (let i = 0; i < rowNumbers.length; i += 1) {
        const r = rowNumbers[i];
        const originalRow = rows[i];
        for (let c = 0; c < meta.colCount; c += 1) {
          const cell = resolveGridCell(store, sheetIdx, r, c, originalRow);
          if (cell === 0) continue;
          if (covered.has(`${r}:${c}`)) continue; // 병합에 덮인 칸 — 앵커의 중복 값이라 시딩하지 않는다
          const key = gridCellKey(sheetIdx, r, c);
          const style = cell.s !== undefined ? store.styles[cell.s] : undefined;
          if (cell.v !== '') {
            const value = seedableCellValue(cell.v, style);
            cells.set(key, value);
            // 이 baseline 맵은 이제 저장 기준선이 아니다 — 저장은 room.dirtyCells/structuralOps로
            // "무엇이 편집됐는지"를 판단한다(Task 3 이후). 그래서 편집된 셀이 재시딩될 때 여기에
            // "편집된 값"이 들어가도 저장 결과에는 영향이 없다. 클라이언트가 참고할 수 있도록
            // 채우는 것뿐.
            baseline.set(key, value);
          }
          // 서식도 콘텐츠와 같은 창(이 페이지)만, 같은 시점에 채운다 — 함수 docblock의
          // "서식을 윈도잉해도 안전한 이유" 참고. GRID_STYLES_BASELINE_MAP도 GRID_BASELINE_MAP과
          // 똑같이 저장 기준선이 아니라 클라이언트 참고용일 뿐이다.
          if (style) {
            const attrs = toCellStyleAttrs(style);
            if (Object.keys(attrs).length > 0) {
              styles.set(key, attrs);
              stylesBaseline.set(key, attrs);
            }
          }
        }
      }
    }, GRID_WINDOW_SEED_ORIGIN);
    await yieldToEventLoop();
  }
}

/**
 * 행/열 삽입·삭제가 만드는 Yjs 트랜잭션에 붙이는 origin 표식.
 *
 * GRID_WINDOW_SEED_ORIGIN과 같은 취지("서버 자신이 쓴 것이니 관찰자가 사용자 편집으로 오해하면
 * 안 된다")지만 **자동저장 면제 대상은 아니다** — 창을 채우고 비우는 것과 달리 행/열 삽입은
 * canonicalGrid를 실제로 바꾸므로 반드시 저장돼야 한다. 그래서 isAutosaveExemptOrigin에는 넣지
 * 않고, room.ts의 관찰자들만 이 표식을 보고 canonical 되반영을 건너뛴다(canonical은 이미
 * gridShift.ts가 직접 옮겨 놓은 상태라, 관찰자가 또 반영하면 같은 이동이 두 번 적용된다).
 *
 * (서식 맵(GRID_STYLES_MAP)의 키 재배치는 예전엔 이 origin으로 표시된 별도 함수
 * `shiftGridStyleKeys`가 직접 옮겼다 — 서식이 절대 윈도잉되지 않던 시절, unseed/reseed로는 서식
 * 키가 옮겨지지 않았기 때문이다. 서식이 콘텐츠와 함께 윈도잉되는 지금은
 * `unseedGridPageFromXlsx`/`seedGridPageFromXlsx`의 통상적인 비우기·재시딩만으로 서식도 새 좌표에
 * 다시 채워지므로 그 함수는 삭제됐다 — room.ts의 `refreshSeededPagesAfterShift` 참고.)
 */
export const GRID_STRUCTURE_SHIFT_ORIGIN = 'structure-shift';

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
  // 서식 맵(GRID_STYLES_MAP/GRID_STYLES_BASELINE_MAP)도 이제 콘텐츠와 함께 비운다 — 서식이 더
  // 이상 "절대 윈도잉하지 않는" 예외가 아니다(seedGridPageFromXlsx의 docblock "서식을 윈도잉해도
  // 안전한 이유" 참고). 서버에는 클라이언트의 서식 편집을 canonical로 되반영하는 경로
  // (room.ts styles.observe → applyStylesChangeToCanonical → setDirtyCell)가 이미 있으므로,
  // 나중에 이 페이지가 다시 구독되면 seedGridPageFromXlsx가 resolveGridCell로 dirtyOverlay를
  // 먼저 확인해 편집된 서식을 그대로 복원한다 — 원본 서식으로 조용히 되돌아가는 일은 없다.
  // 비우는 맵 목록과 "지금 채워져 있는 페이지"를 세는 목록(room.ts의 seededPagesInYdoc)은 반드시
  // 같아야 하므로 공유 상수 하나를 함께 본다(WINDOWED_GRID_CELL_MAPS 주석 참고 — 리뷰 지적 I1).
  const maps = WINDOWED_GRID_CELL_MAPS.map((name) => ydoc.getMap<unknown>(name));
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
export async function seedGridStructureOnly(ydoc: Y.Doc, sheets: readonly LazyGridSheetMeta[]): Promise<void> {
  const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  const merges = ydoc.getMap<readonly GridMergeShape[]>(GRID_MERGES_MAP);

  for (let sheetIdx = 0; sheetIdx < sheets.length; sheetIdx += 1) {
    const sheet = sheets[sheetIdx];
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
export function coerceCellValue(raw: string): string | number | { formula: string } {
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

export function applyStyleToCell(cell: ExcelJS.Cell, style: CellStyleAttrs | undefined): void {
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

/** canonicalGrid 셀(GridCell | 0)에서 표시 텍스트만 뽑는다 — 빈 칸(0 sentinel)은 빈 문자열로
 * 취급한다. */
export function cellText(cell: GridCell | 0 | undefined): string {
  return cell === 0 || cell === undefined ? '' : cell.v;
}

/** 셀이 참조하는 스타일 인덱스를 실제 GridStyle로 해석한다 — LazyGridStore.styles는 공용 스타일
 * 팔레트라 셀은 인덱스(.s)만 들고 있다(Yjs styles 맵과 달리 셀마다 통짜 서식 객체를 직접 들고
 * 있지 않다). */
export function cellStyle(cell: GridCell | 0 | undefined, store: LazyGridStore): GridStyle | undefined {
  if (cell === 0 || cell === undefined || cell.s === undefined) return undefined;
  return store.styles[cell.s];
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
 * 실제로 사용자가 바꾼 칸(값·서식 각각)만 덮어써 재직렬화한다 — docx export와 같은
 * "바뀐 부분만 교체" 원칙. 한 번도 편집되지 않은 칸(수식 셀 포함)은 건드리지 않는다: 수식 셀은
 * 시딩 시 계산된 결괏값 문자열로만 채워지므로, 건드리지 않았는데도 dirtyOverlay에 있다고 전부
 * 써버리면 수식 자체가 정적 값으로 깨진다.
 *
 * 데이터 소스는 Yjs Y.Doc이 아니라 `LazyGridStore`(Task 3, 이제 `room.lazyGrid`)다 — 룸이 항상
 * 최신 상태로 유지하는 "단일 진실". "무엇이 바뀌었는지"는 셀 단위 비교로 재구성하지 않는다 —
 * 대신 두 가지 가벼운 기록으로 재구성한다:
 * - `baselineSheetExtents`: 룸이 열릴 때(=`ingestXlsxLazily` 직후, 아직 편집이 하나도 반영되기
 *   전) 각 시트의 (rowCount, colCount)만 남긴 스냅샷. **`store.sheetMeta[i].rowCount/colCount`로
 *   대체할 수 없다** — 그건 지금까지의 모든 편집(구조연산·직접 확장/축소)이 반영된 "현재" 값이라,
 *   구조연산 재생의 시작점(이 계획 문서가 원래 상정했던 것과 달리)으로 쓰면 재생 수식이 어긋난다.
 *   이 인자는 그래서 원래 계획 문서의 제안과 달리 계속 유지한다(아래 구현 노트 참고).
 * - `structuralOps`: 협업 중 적용된 행/열 삽입·삭제(MESSAGE_GRID_STRUCTURE_OP)의 순서 기록 —
 *   원본을 다시 로드한 이 워크시트에 **같은 순서로 재생**해 store와 물리적 위치를 맞춘다.
 *
 * 값 diff는 `store.dirtyOverlay`(gridCellKey -> 현재 값)를 직접 순회한다 — 이 오버레이의 키
 * 집합 자체가 이미 "실제로 편집된 칸"의 정보이고 값까지 담고 있으므로, 옛 `dirtyCells: Set<string>`
 * 인자처럼 별도로 canonical에서 값을 다시 찾아볼 필요가 없다.
 *
 * **알려진 트레이드오프**: dirtyOverlay는 한 번 편집된 칸을 계속 "편집됨"으로 기억한다 — 원래
 * 수식이었던 칸을 편집했다가 정확히 그 수식의 계산값과 같은 문자열로 되돌려도, 이 칸은 이제
 * 정적 값으로 저장된다(예전에는 값이 같으면 원본 수식이 그대로 보존됐다). 시트 전체를 매번
 * 비교하는 대신 실제 편집 건수에 비례하는 메모리만 쓰기 위해 받아들인 차이다.
 *
 * **구현 노트 — 계획 문서(2026-09-09 xlsx lazy canonical grid) Task 11 item 8과의 의도적 차이**:
 * 그 항목은 "baselineSheetExtents 인자는 완전히 제거한다"고 적었지만, 실제로 그렇게 하면
 * `tests/xlsx-collab.test.ts`의 "export does not call spliceRows/spliceColumns when nothing was
 * actually edited..." 회귀(ExcelJS의 raw rowCount/columnCount가 실제 값 있는 행/열 수보다 부풀려질
 * 수 있는 문제)가 되살아난다: `store.sheetMeta[i].rowCount`는 편집이 쌓이며 계속 바뀌는 "현재" 값이라
 * "구조연산을 재생하기 전 원래 몇 행/열이었는지"라는 별도의 고정된 기준선을 대신할 수 없다(직접
 * 확장/축소는 `structuralOps`에 기록되지 않으므로 현재 값에서 역산할 수도 없다). 그래서 이 인자는
 * `room.baselineSheetExtents`(seedRoom이 `ingestXlsxLazily` 직후 한 번 스냅샷)로 그대로 유지한다.
 */
export async function exportXlsxFromGrid(
  originalXlsxBytes: Uint8Array,
  store: LazyGridStore,
  baselineSheetExtents: ReadonlyArray<{ readonly rowCount: number; readonly colCount: number }>,
  structuralOps: readonly GridStructureOpRecord[],
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

  for (let sheetIdx = 0; sheetIdx < store.sheetMeta.length; sheetIdx += 1) {
    if (removedSheets.has(sheetIdx)) continue;
    const target = store.sheetMeta[sheetIdx];
    // 협업 중 새로 추가된 시트(webviewer 쪽 "+"/삽입 메뉴)는 원본 xlsx에 애초에 없어서
    // visibleSheets에 자리가 없다 — 지금 이름으로 새 워크시트를 만들어 그 자리에 채워 넣는다.
    let sheet = visibleSheets[sheetIdx];
    const isNewSheet = !sheet;
    if (!sheet) {
      sheet = workbook.addWorksheet(target.name);
      visibleSheets[sheetIdx] = sheet;
    } else if (sheet.name !== target.name) {
      // 시트 이름 변경(우클릭 "이름 바꾸기")도 지금까지 원본 워크시트 객체를 그대로 재사용하면서
      // name만 안 옮겨 써서 저장 후 되돌아가 있었다 — 여기서 실제로 반영한다.
      sheet.name = target.name;
    }

    if (isNewSheet) continue; // 원본에 없던 시트 — 재생할 구조 연산도, 맞출 원본 크기도 없다
    const baselineExtent = baselineSheetExtents[sheetIdx];
    if (!baselineExtent) continue;

    // 이 시트에 적용된 행/열 삽입·삭제를 **기록된 순서 그대로** 방금 원본을 다시 읽어들인
    // 이 워크시트에도 재생한다 — canonicalGrid에 있었던 것과 같은 splice를 적용해 두 쪽의 행/열
    // 물리적 위치를 맞춘다. ExcelJS의 spliceRows/spliceColumns는 재생되는 기존 행/열의 값·서식·
    // 행 높이 등을 그대로 옮겨 주므로, 편집되지 않은 칸은 이 재생만으로 이미 올바른 위치에 있다.
    let expectedRowCount = baselineExtent.rowCount;
    let expectedColCount = baselineExtent.colCount;
    for (const op of structuralOps) {
      if (op.sheet !== sheetIdx) continue;
      const onRows = op.kind === 'insertRows' || op.kind === 'deleteRows';
      const insert = op.kind === 'insertRows' || op.kind === 'insertCols';
      if (onRows) {
        if (insert) {
          sheet.spliceRows(op.at + 1, 0, ...Array.from({ length: op.count }, () => [] as unknown[]));
          expectedRowCount += op.count;
        } else {
          sheet.spliceRows(op.at + 1, op.count);
          expectedRowCount -= op.count;
        }
      } else if (insert) {
        sheet.spliceColumns(op.at + 1, 0, ...Array.from({ length: op.count }, () => [] as unknown[]));
        expectedColCount += op.count;
      } else {
        sheet.spliceColumns(op.at + 1, op.count);
        expectedColCount -= op.count;
      }
    }

    // 위 재생 이후에도 남는 차이는 삽입·삭제 연산 없이 rowCount/colCount만 직접 줄어든 경우다
    // (grid-editor.ts 툴바의 "행/열 삭제" 버튼 — 시트 끝에서부터만 지원하는 v1 단순화). 기준은
    // ExcelJS의 raw sheet.rowCount/columnCount가 아니라 위에서 계산한 expectedRowCount/
    // expectedColCount다 — raw 값은 서식만 있는 꼬리 행/열까지 세어 실제와 어긋날 수 있고, 이
    // 함수가 예전에 겪은 `Invalid array length` 버그의 원인이 정확히 그 어긋남이었다.
    if (target.rowCount < expectedRowCount) {
      const removeCount = expectedRowCount - target.rowCount;
      if (sheet.rowCount >= target.rowCount + removeCount) {
        sheet.spliceRows(target.rowCount + 1, removeCount);
      }
    }
    if (target.colCount < expectedColCount) {
      const removeCount = expectedColCount - target.colCount;
      if (sheet.columnCount >= target.colCount + removeCount) {
        sheet.spliceColumns(target.colCount + 1, removeCount);
      }
    }
  }

  // 반대로, 협업 중 삭제된 시트(원래 있었는데 지금 store.sheetMeta엔 그 인덱스가 없는 것 —
  // 배열이라 삭제는 항상 끝에서부터만 반영된다는 뜻, v1 스코프와 동일한 단순화)는 원본에서 그대로
  // 남아 재저장되고 있었다 — 실제로 워크북에서 제거한다.
  for (let sheetIdx = store.sheetMeta.length; sheetIdx < originalSheetCount; sheetIdx += 1) {
    const sheet = visibleSheets[sheetIdx];
    if (sheet) workbook.removeWorksheet(sheet.id);
  }

  // 편집된 칸의 좌표를 시트별로 미리 묶어 둔다 — dirtyOverlay 전체를 시트마다 다시 훑지 않기
  // 위함이다. 오버레이의 값 자체가 이미 "지금 이 칸이 뭐여야 하는지"라, 옛 dirtyCells(Set)처럼
  // canonical에서 값을 별도로 다시 찾아볼 필요가 없다 — 새로 추가된 시트든 원본에 있던 시트든
  // 똑같이 이 오버레이 하나로 다뤄진다(새 시트의 모든 칸은 원본이 없으므로 값이 있다면 반드시
  // 오버레이에 있다 — LazyGridStore.getCell의 "원본 매핑이 -1이면 0"과 같은 이유).
  const dirtyBySheet = new Map<number, Array<{ row: number; col: number; cell: GridCell | 0 }>>();
  for (const [key, cell] of store.dirtyOverlay) {
    const parsed = parseGridCellKey(key);
    if (!parsed || removedSheets.has(parsed.sheetIdx)) continue;
    let list = dirtyBySheet.get(parsed.sheetIdx);
    if (!list) {
      list = [];
      dirtyBySheet.set(parsed.sheetIdx, list);
    }
    list.push({ row: parsed.row, col: parsed.col, cell });
  }

  // 쓰기 패스 — dirtyBySheet에 담긴 좌표만 쓴다(새로 추가된 시트도 원본에 있던 시트도 동일하게).
  // 실제 편집 규모에 비례하므로(전체 시트 크기가 아니라), 대용량 파일에서 편집이 적으면 이 패스는
  // 예전의 전체 셀 스캔보다 훨씬 빠르다.
  let sinceYield = 0;
  const maybeYield = async () => {
    sinceYield += 1;
    if (sinceYield >= YIELD_EVERY_N_ROWS) {
      sinceYield = 0;
      await yieldToEventLoop();
    }
  };
  for (let sheetIdx = 0; sheetIdx < store.sheetMeta.length; sheetIdx += 1) {
    if (removedSheets.has(sheetIdx)) continue;
    const target = store.sheetMeta[sheetIdx];
    const sheet = visibleSheets[sheetIdx];
    if (!sheet) continue;

    const dirtyList = dirtyBySheet.get(sheetIdx);
    if (!dirtyList) continue;
    for (const { row: r, col: c, cell } of dirtyList) {
      if (r >= target.rowCount || c >= target.colCount) continue; // 이후 잘려나간 칸
      const text = cellText(cell);
      sheet.getCell(r + 1, c + 1).value = text === '' ? null : coerceCellValue(text);
      const style = cellStyle(cell, store);
      applyStyleToCell(sheet.getCell(r + 1, c + 1), style ? toCellStyleAttrs(style) : undefined);
      await maybeYield();
    }
  }

  for (let i = 0; i < visibleSheets.length; i += 1) {
    if (removedSheets.has(i)) continue; // 이미 워크북에서 제거된 시트 — 다시 잠글 대상이 아니다
    if (wasProtected[i]) {
      await visibleSheets[i].protect('', {});
    }
  }

  applySheetDisplayOrder(workbook, visibleSheets, store, removedSheets, options.sheetDisplayOrder);

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
  store: LazyGridStore,
  removedSheets: ReadonlySet<number>,
  displayOrder: ReadonlyMap<number, number> | undefined,
): void {
  if (!displayOrder || displayOrder.size === 0) return;
  const survivors: Array<{ sheetIdx: number; worksheet: ExcelJS.Worksheet }> = [];
  for (let sheetIdx = 0; sheetIdx < store.sheetMeta.length; sheetIdx += 1) {
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
