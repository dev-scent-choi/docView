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

/**
 * slideVisuals.get(slideEntryId) -> readonly SlideVisual[] (일반 배열 — Yjs 공유 타입이 아니라
 * 항상 통째로 교체되는 순수 값이다). 이미지·단순 도형(색만 있고 글자 없음)·표·차트/스마트아트
 * 자리표시자처럼 "편집 대상이 아니라 화면에 보여주기만 하는" 요소들 — slideContents(텍스트박스,
 * 협업 편집 대상)와 별도 맵에 두는 이유는 duplicateSlideContent/exportPptxFromSlides가 이미
 * "모든 항목은 {meta, content} 모양의 도형이다"를 전제로 짜여 있어서, 그 전제를 깨지 않고
 * 완전히 병렬로 추가하기 위함이다. 내보낼 때 이 값을 쓰지 않는다 — 원본 XML을 그대로 복사하는
 * 기존 export 경로가 이미 이 요소들을 바이트 그대로 보존한다. */
export const PPTX_SLIDE_VISUALS_MAP = 'slideVisuals';

export interface SlideVisual {
  readonly kind: 'image' | 'shape' | 'table' | 'placeholder';
  readonly x: number | null;
  readonly y: number | null;
  readonly w: number | null;
  readonly h: number | null;
  /** kind:'image' — data: URI(원본 미디어 바이트를 base64로 그대로 인라인) */
  readonly dataUri?: string;
  /** kind:'shape' — 단색 채우기(#RRGGBB), 없으면 채우기 없는 도형 */
  readonly fill?: string;
  /** kind:'table' — 행 단위 셀 텍스트(행렬, 서식 없이 텍스트만) */
  readonly rows?: readonly (readonly string[])[];
  /** kind:'placeholder' — 표시할 라벨(예: "차트", "스마트아트") */
  readonly label?: string;
}

/** 새로 추가한(원본에 없던) 텍스트박스의 기본 위치/크기 — 슬라이드 중앙에 가로 60%, 세로 20%.
 * 표준 16:9 슬라이드 크기(12192000 x 6858000 EMU) 기준으로 잡되, 실제 슬라이드 크기가
 * 다르면 호출부가 비율로 다시 계산한다(이 상수는 "기본 비율"의 근거로만 쓴다). */
export const NEW_TEXTBOX_WIDTH_RATIO = 0.6;
export const NEW_TEXTBOX_HEIGHT_RATIO = 0.2;

/**
 * src/client/editor-slides.ts의 "텍스트 추가" 버튼 핸들러가 쓰는 가드(FINAL REVIEW C1 / Fix 1)의
 * boolean 공식을 그대로 뽑아낸 순수 함수. 원래 이 조건은 그 핸들러 안에 인라인으로만 있었는데,
 * editor-slides.ts는 모듈 최상단에서 `document`/`location`을 곧바로 읽어(예: `document.getElementById`)
 * import 시점에 이미 브라우저 전역을 요구하므로 — 이 저장소엔 아직 jsdom 테스트 환경이 없어서
 * (plain vitest, node 환경) 그 파일 자체를 테스트에서 직접 import할 수 없다. 이 함수는 그 가드의
 * "판단" 부분만 떼어내 이 파일(순수 타입/상수만 있고 DOM이나 서버 전용 의존성이 없어 서버·클라이언트
 * 양쪽에서 이미 그대로 import해 쓰는 공유 모듈)로 옮긴 것이라 plain vitest로 직접 단위 테스트할 수
 * 있다 — tests/editorSlidesAddTextboxGuard.test.ts 참고. 실제 DOM 이벤트 연결(버튼 클릭, Yjs 트랜잭션
 * 등)은 여전히 editor-slides.ts에 남아 있고 여기서 검증하지 않는다.
 *
 * @param entryKind 텍스트박스를 추가하려는 슬라이드(SlideEntry)의 kind.
 * @param isAlreadySeeded 이 슬라이드의 slideContents 항목이 이미 채워졌는지
 *   (`slideContents.has(entryId)`).
 * @returns true면 "텍스트 추가"를 막아야 한다 — kind가 'blank'가 아닌데 아직 시딩되지 않은 경우뿐이다.
 *   (원본/복제 슬라이드인데 서버의 시딩 응답을 아직 못 받은 상태 — 여기서 새 perSlide를 만들면
 *   다음 export가 "baseline에는 있는데 perSlide엔 없다 = 삭제됐다"로 오해해 원본 도형을 전부
 *   지워버린다. kind가 'blank'면 애초에 원본이라 부를 게 없으므로 막을 이유가 없다.)
 */
export function shouldBlockAddTextbox(entryKind: SlideEntryKind, isAlreadySeeded: boolean): boolean {
  return !isAlreadySeeded && entryKind !== 'blank';
}
