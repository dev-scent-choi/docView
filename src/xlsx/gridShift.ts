import type { GridMerge } from '../spreadsheet.js';

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
 *
 * room.ts의 applyGridStructureOp가 LazyGridStore의 logicalToOriginalRow/Col 매핑 배열을 스플라이스한
 * 뒤, 그와 별도로 관리되는 sheetMeta.merges를 같은 축으로 옮기는 데 쓴다(LazyGridStore.applyRowShift/
 * applyColShift 자체는 merges/colWidths를 건드리지 않는다 — lazyGridStore.ts 주석 참고).
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
