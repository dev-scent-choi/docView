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
/** 수식 셀의 계산된 표시값만 담는 Y.Map — "시트:행:열" 키(gridCellKey와 동일 형식). 수식이
 * 아닌 셀은 이 맵에 항목이 없다. cells는 항상 원본 텍스트(수식이면 수식 자체)만 담고, 여기에
 * 계산된 값을 절대 섞지 않는다 — 섞으면 저장 시 수식이 정적 값으로 깨진다. */
export const GRID_COMPUTED_MAP = 'computed';
export const GRID_STRUCTURE_MAP = 'structure';
/**
 * 문서 종류(xlsx/docx) 공통 — 서버가 실제로 디스크에 저장을 마친 시각(epoch ms)을 담아두는 작은
 * 맵. 일반 Y.Map이라 기존 sync 프로토콜을 그대로 타고 모든 클라이언트에 자동 전파되므로, 저장
 * 상태 표시를 위해 별도 커스텀 WebSocket 메시지를 만들 필요가 없다.
 */
export const DOC_META_MAP = 'docMeta';
export const DOC_META_SAVED_AT_KEY = 'savedAt';
/**
 * flushToStore가 저장에 실패했을 때(디스크 쓰기 예외, 내보내기 실패 등) 그 실패 시각(epoch ms)을
 * 기록하는 키. savedAt과 같은 맵에 두고, 클라이언트는 이 값이 savedAt보다 최신이면 "저장 실패"
 * 상태로 본다 — 별도 정리(clear) 없이도 다음 저장이 성공해 savedAt이 이 값을 앞지르는 순간
 * 자연스럽게 실패 표시가 사라진다. 이 값이 없으면(구버전 세션 등) 실패한 적 없는 것으로 본다.
 */
export const DOC_META_SAVE_ERROR_AT_KEY = 'saveErrorAt';
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
  /** 시트 탭이 표시되는 순서 — 시트를 가리키는 안정적인 키(cells/styles/merges가 전부 이 번호로
   * 참조된다)와 화면상 순서를 분리해 둬야, 탭을 드래그해 순서를 바꿔도 그 많은 셀 데이터를 전부
   * 다시 키잉하지 않고 이 작은 숫자 하나만 바꾸면 된다. 없으면(구버전 문서) 키 번호 자체를 순서로
   * 쓴다. */
  readonly order?: number;
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

/** 숫자 표시 형식 — 실제 엑셀 서식 코드를 그대로 저장하지 않고 자주 쓰는 몇 가지로 단순화한다
 * (툴바의 회계/백분율/쉼표 버튼과 1:1 대응). 원본 값(cells)은 항상 그대로고, 이건 "어떻게
 * 보여줄지"만 바꾼다 — 수식 계산 결과 표시와 같은 방식(값과 표시를 분리). */
export type NumberFormatKind = 'number' | 'currency' | 'percent';

export interface CellStyleAttrs {
  readonly b?: 1;
  readonly i?: 1;
  readonly u?: 1;
  readonly st?: 1;
  readonly col?: string;
  readonly bg?: string;
  readonly a?: 'l' | 'c' | 'r';
  /** 세로 정렬 — 위/가운데/아래. 없으면 엑셀 기본(숫자는 아래, 대개 아래 정렬)과 동일하게 둔다. */
  readonly va?: 't' | 'm' | 'b';
  /** 자동 줄 바꿈 — 셀 안에서 내용이 길면 줄을 바꿔 여러 줄로 보여준다. */
  readonly wrap?: 1;
  /** 글꼴 이름. 기본 글꼴과 같으면 굳이 저장하지 않는다. */
  readonly ff?: string;
  /** 글자 크기(px). */
  readonly fs?: number;
  /** 숫자 표시 형식 종류(쉼표/통화/백분율) — 없으면 "일반"(원본 텍스트 그대로). */
  readonly nf?: NumberFormatKind;
  /** nf가 있을 때만 의미 있는 소수 자릿수. 없으면 0. */
  readonly nd?: number;
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

/**
 * 윈도잉(페이지 단위 로딩)의 페이지 크기 — 페이지 N은 0-based로 행 [N*500, (N+1)*500)이다.
 *
 * 서버(room.ts, src/xlsx/collab.ts)와 클라이언트 번들(src/client/grid-editor.ts)이 **반드시 같은
 * 값**을 봐야 한다: 클라이언트는 지금 보이는 행 범위를 이 값으로 나눠 "페이지 번호"를 구해 구독
 * 메시지를 보내고, 서버는 같은 계산으로 그 페이지의 행 범위를 Yjs에 채운다 — 두 쪽이 어긋나면
 * 클라이언트가 요청한 곳과 서버가 채운 곳이 달라져 화면에 빈 칸만 보인다. 그래서 서버 전용
 * 의존성이 딸려오는 xlsx/collab.ts가 아니라 이 공유 상수 파일에 둔다(GRID_CELLS_MAP 등이 여기
 * 있는 것과 같은 이유).
 */
export const PAGE_SIZE_ROWS = 500;

/** 편집 세션 하나가 감당할 수 있다고 보는 상한 — 넘으면 UI에 부드러운 경고만 띄운다(차단하지 않음). */
export const EDIT_SOFT_ROW_LIMIT = 500;
export const EDIT_SOFT_COL_LIMIT = 60;
