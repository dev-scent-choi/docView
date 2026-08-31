/**
 * xlsx 동시편집이 쓰는 Y.Doc 최상위 공유 타입 이름 — 서버(room.ts, src/xlsx/collab.ts)와
 * 클라이언트 번들(src/client/grid-editor.ts) 양쪽에서 그대로 가져다 쓴다(순수 상수라 어느
 * 쪽에서 임포트해도 Node 전용 의존성이 딸려오지 않는다).
 *
 * 셀은 "시트:행:열" 키의 평면 Y.Map으로 둔다 — 셀 하나하나가 독립된 CRDT 엔트리라
 * 서로 다른 셀을 동시에 편집해도 절대 충돌하지 않는다(같은 셀 동시 편집만 후속 쓰기가 이긴다).
 * 시트 구조(행/열 개수, 이름)는 별도 Y.Map에 시트 인덱스별로 통짜 JSON을 둔다 — 행/열
 * 추가·삭제 같은 구조 변경은 셀 편집보다 훨씬 드물어 세밀한 CRDT가 필요 없다.
 */
export const GRID_CELLS_MAP = 'cells';
export const GRID_STRUCTURE_MAP = 'structure';
/**
 * 시딩 시점의 셀 값을 그대로 담아두는 불변 스냅샷 — 클라이언트는 절대 쓰지 않는다.
 * 수식 셀은 계산된 결괏값 문자열로 시딩되는데, 저장 시 `cells`에 있는 모든 키를 원본에
 * 그대로 쓰면 사용자가 손대지 않은 수식까지 정적 값으로 덮어써 버린다. `cells`와 이 값을
 * 비교해 실제로 달라진(=사용자가 편집한) 셀만 원본에 반영하기 위한 기준선이다.
 */
export const GRID_BASELINE_MAP = 'baseline';

export function gridCellKey(sheetIdx: number, row: number, col: number): string {
  return `${sheetIdx}:${row}:${col}`;
}

export function parseGridCellKey(key: string): { sheetIdx: number; row: number; col: number } | null {
  const match = /^(\d+):(\d+):(\d+)$/.exec(key);
  if (!match) return null;
  return { sheetIdx: Number(match[1]), row: Number(match[2]), col: Number(match[3]) };
}

export interface SheetStructure {
  readonly name: string;
  readonly rowCount: number;
  readonly colCount: number;
  /** 엑셀 "틀 고정" — 위 N행/왼쪽 N열은 스크롤해도 계속 보인다. 없으면 둘 다 0. */
  readonly freezeRows?: number;
  readonly freezeCols?: number;
  /** 열별 실제 엑셀 너비(px, 시딩 시 spreadsheet.ts가 문자 단위를 변환해 넣는다) — 시트에 실제로
   * 있는 열 개수만큼만 있고, 그 뒤(편집기가 여백으로 더 그리는 빈 열들)는 없다. 없으면(구버전 문서
   * 등) 그리드 에디터가 기본 폭으로 그린다. */
  readonly colWidths?: readonly number[];
}

/**
 * 셀 값과 같은 "시트:행:열" 키로 병행 관리하는 서식 레이어 — 값(cells)과 분리해 두면
 * 값만 바뀐 셀과 서식만 바뀐 셀을 export 시 독립적으로 판단할 수 있다(둘 다 baseline과
 * 비교해 실제로 바뀐 것만 원본에 반영 — cells/baseline과 동일한 패턴).
 */
export const GRID_STYLES_MAP = 'styles';
export const GRID_STYLES_BASELINE_MAP = 'stylesBaseline';

/** 테두리 굵기 — 1=가는 선, 2=굵은 선(엑셀 thin/medium을 단순화). 색은 항상 기본(검정 계열)만 지원. */
export type BorderWeight = 1 | 2;

export interface CellStyleAttrs {
  readonly b?: 1;
  readonly i?: 1;
  readonly u?: 1;
  readonly st?: 1;
  readonly col?: string;
  readonly bg?: string;
  readonly a?: 'l' | 'c' | 'r';
  /** 시트 보호가 걸려 있고 이 셀이 잠겨 있음(=편집 불가) — 엑셀의 "시트 보호 + 셀 잠금" 그대로. */
  readonly lk?: 1;
  /** 엑셀 데이터 유효성 검사(정수/소수)가 걸려 있음 — 숫자가 아닌 값 입력을 막고 안내를 띄운다. */
  readonly nu?: 1;
  readonly bt?: BorderWeight;
  readonly br?: BorderWeight;
  readonly bb?: BorderWeight;
  readonly bl?: BorderWeight;
}

/**
 * 병합 셀 — 원본 xlsx의 병합 정보를 시트별로 통짜 배열로 둔다(구조 변경이므로 셀별 CRDT
 * 불필요, structure 맵과 같은 패턴). r/c는 1-based(엑셀 좌표계 그대로), rs/cs는 병합 범위의
 * 행/열 개수 — spreadsheet.ts의 GridMerge와 동일한 모양이지만, 이 파일은 클라이언트 번들에도
 * 들어가므로 ExcelJS를 끌어오는 spreadsheet.ts를 직접 import하지 않고 구조적으로만 맞춘다.
 */
export interface GridMergeShape {
  readonly r: number;
  readonly c: number;
  readonly rs: number;
  readonly cs: number;
}

export const GRID_MERGES_MAP = 'merges';

/** 편집 세션 하나가 감당할 수 있다고 보는 상한 — 넘으면 UI에 부드러운 경고만 띄운다(차단하지 않음). */
export const EDIT_SOFT_ROW_LIMIT = 500;
export const EDIT_SOFT_COL_LIMIT = 60;
