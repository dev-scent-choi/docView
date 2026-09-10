/**
 * pptx 동시편집이 쓰는 Y.Doc 최상위 공유 타입 이름 — 서버(room.ts, src/pptx/collab.ts)와
 * 클라이언트 번들(src/client/editor-slides.ts) 양쪽에서 그대로 가져다 쓴다. 슬라이드 "내용"은
 * 절대 Yjs에 담지 않는다 — 이 배열은 오직 "어떤 슬라이드가 몇 번째냐"만 나타낸다.
 */
export type SlideEntryKind = 'original' | 'duplicate' | 'blank';

export interface SlideEntry {
  readonly id: string; // 이 배열 항목의 고유 UUID (Yjs 식별용, 슬라이드 파일 번호와 무관)
  readonly kind: SlideEntryKind;
  /** kind가 'original' | 'duplicate'일 때만 존재 — 원본 zip의 실제 슬라이드 파트 이름
   *  (예: 'slide3.xml', 확장자 포함, 'ppt/slides/' 접두어 없음). 'duplicate'도 원본과 같은
   *  sourceSlidePart를 가리킨다(내보낼 때 그 파트를 복제해서 새 번호를 붙인다). */
  readonly sourceSlidePart: string | null;
  /** kind가 'original' | 'duplicate'일 때만 존재 — 업로드 시점에 만들어 둔 "바로보기"용 PDF에서
   *  이 슬라이드가 몇 번째 페이지였는지(1-based). 썸네일을 그릴 때 `/api/view/pages/:page`에
   *  그대로 쓴다 — 편집 중에는 이 PDF를 다시 만들지 않으므로 항상 유효한 페이지 번호다. */
  readonly sourcePageNumber: number | null;
}

export const PPTX_SLIDES_ARRAY = 'slides';

/**
 * slideContents.get(slideEntryId) -> Y.Map(shapeId -> Y.Map({ meta: SlideShapeMeta, content: Y.XmlFragment })).
 * 슬라이드 "내용"은 이제 여기 담긴다 — 위 PPTX_SLIDES_ARRAY 자체는 여전히 순서/포인터만
 * 나타낸다(이 파일 위쪽 주석은 "슬라이드 배열 자체"에는 내용을 안 담는다는 뜻으로 남겨 둔다).
 */
export const PPTX_SLIDE_CONTENTS_MAP = 'slideContents';

/** slideContents 안, 텍스트박스 하나의 위치 메타데이터. content(Y.XmlFragment)는 같은
 * 텍스트박스의 Y.Map 안에 'content' 키로 별도로 들어간다(이 인터페이스에는 없음 — Yjs
 * 공유 타입은 plain interface로 표현할 수 없어서 분리한다). */
export interface SlideShapeMeta {
  readonly shapeId: string;
  /** EMU(1pt = 12700 EMU) 단위, 원본 슬라이드 XML의 <a:off>/<a:ext>와 같은 단위.
   * null = 원본 XML에 명시적 위치가 없어(레이아웃 상속 등) 기본값을 썼다는 뜻. */
  readonly x: number | null;
  readonly y: number | null;
  readonly w: number | null;
  readonly h: number | null;
}

/** 새로 추가한(원본에 없던) 텍스트박스의 기본 위치/크기 — 슬라이드 중앙에 가로 60%, 세로 20%.
 * 표준 16:9 슬라이드 크기(12192000 x 6858000 EMU) 기준으로 잡되, 실제 슬라이드 크기가
 * 다르면 호출부가 비율로 다시 계산한다(이 상수는 "기본 비율"의 근거로만 쓴다). */
export const NEW_TEXTBOX_WIDTH_RATIO = 0.6;
export const NEW_TEXTBOX_HEIGHT_RATIO = 0.2;
