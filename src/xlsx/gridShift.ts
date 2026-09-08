import {
  yieldToEventLoop,
  YIELD_EVERY_N_ROWS,
  type GridCell,
  type GridMerge,
  type GridSheet,
} from '../spreadsheet.js';

/**
 * 행/열 삽입·삭제를 **canonical 그리드에서** 수행하는 원시 연산들.
 *
 * **왜 서버의 canonical에서 하는가**: 예전에는 클라이언트(grid-editor.ts의
 * shiftRowsAndMerges/shiftColsAndMerges)가 Yjs의 cells/styles 키를 직접 한 칸씩 옮겼다. 그
 * 접근은 "시트의 모든 셀이 Yjs에 들어 있다"를 전제하는데, 이 프로젝트가 Yjs를 "지금 구독 중인
 * 500행 페이지만 담는 창"으로 좁힌 뒤로는 그 전제가 깨졌다 — 창 밖의 행은 옮겨지지 않아 시트
 * 나머지 전체가 한 칸씩 어긋난다. 게다가 그 루프의 상한은 시트 실제 크기가 아니라 **하드코딩된
 * 1000**이라(30만 행을 지원하기 전 시대의 코드), 1001행이 넘는 시트에서는 행 1000의 데이터를
 * 덮어쓰고 1001행 이후는 아예 손대지 않았다.
 *
 * canonical 그리드는 구독 상태와 무관하게 **항상 시트 전체**를 담으므로(Room.canonicalGrid),
 * 여기서 옮기면 두 문제가 동시에 사라진다. 셀의 서식 인덱스(GridCell.s)는 셀 객체 안에 들어
 * 있으므로 행/열을 옮기면 서식도 함께 따라온다 — 별도 처리가 필요 없다.
 */

type MutableGridSheet = { -readonly [K in keyof GridSheet]: GridSheet[K] };
type MutableGridRow = Array<GridCell | 0>;

/**
 * 병합 하나의 한 축(행 또는 열)을 삽입/삭제에 맞춰 보정한다 — 1-based 시작점(엑셀 좌표계)과
 * 칸 수(span)를 그대로 받고 돌려준다.
 *
 * 삽입: 삽입 지점이 병합보다 앞이면 병합 전체가 밀리고, 병합 **안쪽**이면 병합이 그만큼 넓어진다
 * (엑셀과 같은 동작). 삭제: 삭제 범위와 겹치는 칸 수만큼 좁아지고, 병합보다 앞에서 지워진 칸
 * 수만큼 시작점이 당겨진다. span이 0 이하가 되면 그 병합은 통째로 사라진 것이다.
 */
function shiftMergeAxis(
  start1Based: number,
  span: number,
  at: number,
  count: number,
  insert: boolean,
): { start: number; span: number } {
  const start0 = start1Based - 1;
  if (insert) {
    if (start0 >= at) return { start: start1Based + count, span };
    if (at < start0 + span) return { start: start1Based, span: span + count };
    return { start: start1Based, span };
  }
  const deleteEnd = at + count;
  // 병합 범위 [start0, start0+span)과 삭제 범위 [at, at+count)가 실제로 겹치는 칸 수.
  const overlap = Math.max(0, Math.min(start0 + span, deleteEnd) - Math.max(start0, at));
  // 병합 시작점 **앞에서** 지워진 칸 수 — 그만큼 병합 전체가 당겨진다.
  const removedBefore = Math.max(0, Math.min(deleteEnd, start0) - at);
  return { start: start0 - removedBefore + 1, span: span - overlap };
}

/**
 * 시트의 병합 목록 전체를 행(또는 열) 삽입·삭제에 맞춰 다시 만든다. 완전히 사라진 병합과
 * 1×1로 쪼그라든 병합(= 더 이상 병합이 아니다)은 목록에서 뺀다.
 */
export function shiftMerges(
  merges: readonly GridMerge[],
  axis: 'row' | 'col',
  at: number,
  count: number,
  insert: boolean,
): GridMerge[] {
  const out: GridMerge[] = [];
  for (const merge of merges) {
    const shifted =
      axis === 'row'
        ? shiftMergeAxis(merge.r, merge.rs, at, count, insert)
        : shiftMergeAxis(merge.c, merge.cs, at, count, insert);
    const next: GridMerge =
      axis === 'row'
        ? { r: shifted.start, c: merge.c, rs: shifted.span, cs: merge.cs }
        : { r: merge.r, c: shifted.start, rs: merge.rs, cs: shifted.span };
    if (next.rs <= 0 || next.cs <= 0) continue;
    if (next.rs === 1 && next.cs === 1) continue; // 한 칸짜리는 병합이 아니다
    out.push(next);
  }
  return out;
}

/**
 * 행 삽입 — `at` 자리에 빈 행 `count`개를 끼워 넣고 그 아래를 전부 밀어 내린다.
 *
 * **양보가 필요 없는 유일한 방향**(계획의 Global Constraint #1을 어기는 게 아니라 애초에 해당하지
 * 않는다): 이건 행 배열 하나에 대한 splice 한 번, 즉 포인터 블록 복사다. 실측(이 환경, 30만 행 ×
 * 120열 = 힙 2.3GB짜리 시트): 삽입 지점과 무관하게 **0.002 ~ 0.97ms**. 반면 아래
 * insertCanonicalCols는 모든 행의 배열을 각각 건드려야 해서 같은 시트에서 197ms가 걸린다 —
 * 그래서 그쪽만 청크+양보로 돈다.
 */
export function insertCanonicalRows(sheet: GridSheet, at: number, count: number): void {
  const mutable = sheet as MutableGridSheet;
  const rows = mutable.rows as MutableGridRow[];
  const blanks: MutableGridRow[] = [];
  for (let i = 0; i < count; i += 1) blanks.push(new Array<GridCell | 0>(mutable.colCount).fill(0));
  // at이 rows.length보다 뒤일 수 있다(구조 메시지로 rowCount만 커지고 배열 패딩은 백그라운드
  // 체인에 밀려 있는 성긴 상태 — Room.canonicalGrowth 주석 참고). 그 경우 끼워 넣을 자리가
  // 어차피 전부 빈 행이므로 배열 끝에 붙여도 결과가 같다(splice 자신도 start를 length로 자른다).
  rows.splice(Math.min(at, rows.length), 0, ...blanks);
  mutable.rowCount = mutable.rowCount + count;
  mutable.merges = shiftMerges(mutable.merges, 'row', at, count, true);
}

/** 행 삭제 — `at`부터 `count`행을 빼고 그 아래를 전부 끌어올린다(위 삽입과 같은 이유로 동기). */
export function deleteCanonicalRows(sheet: GridSheet, at: number, count: number): void {
  const mutable = sheet as MutableGridSheet;
  const rows = mutable.rows as MutableGridRow[];
  if (at < rows.length) rows.splice(at, count);
  mutable.rowCount = Math.max(1, mutable.rowCount - count);
  mutable.merges = shiftMerges(mutable.merges, 'row', at, count, false);
}

/**
 * 열 삽입 — 모든 행 배열의 `at` 자리에 빈 칸 `count`개를 끼워 넣는다.
 *
 * 행 삽입과 달리 이건 **행 수만큼의 개별 배열 조작**이라 진짜로 O(rowCount)다(실측: 30만 행에서
 * 약 197ms의 통째 정지). 그래서 이 코드베이스의 다른 모든 대형 행 루프와 똑같이
 * YIELD_EVERY_N_ROWS행마다 실제로 이벤트 루프에 양보한다(spreadsheet.ts의 yieldToEventLoop 주석,
 * 계획의 Global Constraint #1).
 *
 * colCount·colWidths·merges는 행 루프가 **끝난 뒤** 갱신한다 — 그래야 루프 도중 양보 틈에 끼어든
 * 읽기(내보내기 등)가 "새 폭인데 아직 안 밀린 행"을 보는 대신 "옛 폭 + 일부만 밀린 행"을 본다.
 * 어느 쪽도 완전하진 않지만(청크 사이에 양보하는 모든 대형 연산이 공유하는 성질 —
 * Room.canonicalGrowth의 같은 설명 참고), 자동저장은 이 연산이 전부 끝난 뒤에야 예약되므로
 * 정상 경로에서 반쪽 상태가 파일로 나가지는 않는다.
 */
export async function insertCanonicalCols(sheet: GridSheet, at: number, count: number): Promise<void> {
  const mutable = sheet as MutableGridSheet;
  const rows = mutable.rows as MutableGridRow[];
  const blanks = new Array<GridCell | 0>(count).fill(0);
  for (let start = 0; start < rows.length; start += YIELD_EVERY_N_ROWS) {
    const end = Math.min(start + YIELD_EVERY_N_ROWS, rows.length);
    for (let r = start; r < end; r += 1) {
      const row = rows[r];
      // 폭이 at에 못 미치는 성긴 행은 건드리지 않는다 — 그 행의 at 이후는 이미 빈 칸이라
      // 끼워 넣어도 결과가 같고, 괜히 늘리면 성긴 표현의 이점(메모리)만 잃는다.
      if (row.length > at) row.splice(at, 0, ...blanks);
    }
    await yieldToEventLoop();
  }
  mutable.colCount = mutable.colCount + count;
  const widths = mutable.colWidths as number[];
  // 새 열의 너비는 밀려나는 그 열의 너비를 그대로 물려준다(엑셀과 같은 동작).
  if (widths.length > at) widths.splice(at, 0, ...new Array<number>(count).fill(widths[at]));
  mutable.merges = shiftMerges(mutable.merges, 'col', at, count, true);
}

/** 열 삭제 — 위 삽입과 같은 이유로 같은 청크+양보 구조를 쓴다. */
export async function deleteCanonicalCols(sheet: GridSheet, at: number, count: number): Promise<void> {
  const mutable = sheet as MutableGridSheet;
  const rows = mutable.rows as MutableGridRow[];
  for (let start = 0; start < rows.length; start += YIELD_EVERY_N_ROWS) {
    const end = Math.min(start + YIELD_EVERY_N_ROWS, rows.length);
    for (let r = start; r < end; r += 1) {
      const row = rows[r];
      if (row.length > at) row.splice(at, count);
    }
    await yieldToEventLoop();
  }
  mutable.colCount = Math.max(1, mutable.colCount - count);
  const widths = mutable.colWidths as number[];
  if (widths.length > at) widths.splice(at, count);
  mutable.merges = shiftMerges(mutable.merges, 'col', at, count, false);
}
