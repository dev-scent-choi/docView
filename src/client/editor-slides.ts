import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import * as encoding from 'lib0/encoding';
import { EditorState, type Command } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { baseKeymap, toggleMark } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import { ySyncPlugin, yCursorPlugin, yUndoPlugin, ySyncPluginKey, undo, redo } from 'y-prosemirror';
import {
  PPTX_SLIDES_ARRAY,
  PPTX_SLIDE_CONTENTS_MAP,
  NEW_TEXTBOX_WIDTH_RATIO,
  NEW_TEXTBOX_HEIGHT_RATIO,
  type SlideEntry,
  type SlideShapeMeta,
} from '../collab/pptxShared.js';
import { DOC_META_MAP, DOC_META_SAVED_AT_KEY, DOC_META_SAVE_ERROR_AT_KEY } from '../collab/gridShared.js';
import { pickPresenceColorForName, renderPresenceAvatars } from './presence.js';
import { pptxTextSchema } from '../pptx/textSchema.js';

/** src/pptx/collab.ts가 서버 쪽에서 선언한 것과 같은 모양이지만, 그 파일은 node:crypto/fflate 등
 * 서버 전용 의존성을 끌고 있어 클라이언트 번들에 그대로 import할 수 없다 — Yjs 공유 타입 자체는
 * 구조적 타이핑이라 여기서 같은 모양을 독립적으로 다시 선언해도 완전히 호환된다. */
type ShapeContentMap = Y.Map<unknown>;
type PerSlideContentsMap = Y.Map<ShapeContentMap>;

interface Permissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
  readonly canEdit: boolean;
}

interface ViewInfo {
  readonly docId: string;
  readonly name: string;
  readonly viewer: string;
  readonly permissions: Permissions;
}

const token = new URLSearchParams(location.search).get('token');
const titleEl = document.getElementById('doc-title')!;
const connStatusEl = document.getElementById('conn-status')!;
const railEl = document.getElementById('slides-rail') as HTMLElement;
const canvasEl = document.getElementById('slide-canvas') as HTMLElement;
const cardTemplate = document.getElementById('slide-card-template') as HTMLTemplateElement;
const presenceListEl = document.getElementById('presence-list') as HTMLElement;

// 어떤 슬라이드가 오른쪽 캔버스에 표시 중인지 — renderCanvasForSlide()가 이 값을 읽어 실제
// 콘텐츠를 그린다.
let activeSlideId: string | null = null;
// Task 8이 남겨 둔 훅 포인트 — main()에서 renderCanvasForSlide로 연결한다(아래 참고).
let onActiveSlideChange: ((entryId: string) => void) | null = null;

function showFatalError(message: string): void {
  document.getElementById('error-message')!.textContent = message;
  document.getElementById('error-screen')!.hidden = false;
}

async function fetchViewInfo(): Promise<ViewInfo> {
  const res = await fetch(`/api/view/info?token=${encodeURIComponent(token ?? '')}`);
  const json = (await res.json()) as { success: boolean; data?: ViewInfo; error?: string };
  if (!res.ok || !json.success || !json.data) throw new Error(json.error ?? '문서 정보를 불러오지 못했습니다.');
  return json.data;
}

let canEdit = false;
let slidesArray: Y.Array<SlideEntry>;
let dragFromIndex: number | null = null;
// main()에서 대입된다(slidesArray와 같은 패턴) — subscribeSlideContent/renderCanvasForSlide가
// 모듈 스코프 함수라 provider/slideContents도 모듈 스코프에 있어야 접근할 수 있다.
let provider: WebsocketProvider;
let slideContents: Y.Map<PerSlideContentsMap>;

/**
 * 슬라이드 크기(EMU)를 캔버스 픽셀로 스케일한다. `/api/view/info`(ViewInfo)에는 슬라이드 크기
 * 필드가 없고, room.ts/viewer.ts 어디에도 sldSz류 값을 따로 내려주는 경로가 없다(grep으로 확인
 * 완료) — 표준 16:9 슬라이드 크기(12192000 x 6858000 EMU)를 기본값으로 쓴다. 실제 프레젠테이션의
 * 진짜 크기를 반영하는 것은 이후 개선 과제로 남긴다(이번 태스크에서 서버 API를 넓히지 않는다).
 */
const SLIDE_W_EMU = 12192000;
const SLIDE_H_EMU = 6858000;
const EMU_PER_PX = SLIDE_W_EMU / 960; // 캔버스 폭을 960px 기준으로 스케일(반응형은 이후 과제)

function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

/** room.ts의 MESSAGE_SLIDE_CONTENT_SUBSCRIBE와 **같은 값이어야 한다**(grid-editor.ts의
 * MESSAGE_GRID_STRUCTURE_OP 주석과 같은 이유 — 0~3은 y-websocket 표준 프로토콜이 이미 쓰고
 * 있어 새 값은 그 범위를 벗어나게 둔다). */
const MESSAGE_SLIDE_CONTENT_SUBSCRIBE = 5;

/**
 * 슬라이드 하나를 "지금 연다"고 서버에 알린다 — 서버가 slideContents.get(entryId)를 아직 안
 * 채웠으면 원본 XML에서 지연 시딩한다(room.ts의 seedSlideContentIfNeeded). 전송 방식은
 * grid-editor.ts의 sendPageSubscription과 완전히 같은, 이미 검증된 패턴이다: WebsocketProvider가
 * 공개 필드로 들고 있는 실제 WebSocket 인스턴스(node_modules/y-websocket/dist/src/y-websocket.d.ts:
 * `ws: WebSocket | null`)에 varUint 메시지 타입 + varString JSON 페이로드를 그대로 실어 보낸다 —
 * y-websocket 자신도 내부적으로 `this.ws.send(encoding.toUint8Array(encoder))`로 syncStep1을
 * 보내는 것과 같은 경로다.
 */
// 서버에 아직 시딩 응답을 못 받은(=slideContents에 그 키가 아직 없는) 슬라이드 id 집합 —
// 같은 슬라이드 카드를 빠르게 두 번 클릭해도 두 번째 요청은 보내지 않기 위한 것이다(서버는
// 요청마다 디스크 읽기+unzip+XML 파싱을 새로 하므로, 첫 요청이 아직 응답하기 전의 중복
// 요청은 그 비용을 그대로 낭비하고 두 번째는 결국 조용히 no-op된다). slideContents.observe가
// 실제로 그 슬라이드 키가 채워진 것을 감지하면(main() 참고) 이 집합에서 제거한다.
//
// 다만 서버 쪽 시딩이 어떤 이유로든(에러 throw, 읽기 전용 연결, 잘못된 payload, 존재하지 않는
// 슬라이드 엔트리 등) 실패하면 slideContents에는 아무것도 쓰이지 않으므로 observe가 영영 안
// 불리고, 이 집합에 남은 entryId는 세션이 끝날 때까지 제거되지 않는다 — 그러면 그 슬라이드는
// 이후 몇 번을 다시 클릭해도 이 함수 맨 위 가드에서 조용히 no-op되어 영구히 빈 캔버스로 남는다.
// 이를 막기 위해 아래 subscribeSlideContent에서 항목을 추가할 때마다 유한한 수명(타임아웃)을
// 함께 건다 — 정상적으로 시딩되면 observe가 먼저 지우고, 타임아웃이 나중에 발동해도
// Set.delete는 이미 없는 값에 대해 안전한 no-op이라 아무 문제도 없다. 시딩이 실패한 경우엔
// 타임아웃이 먼저 발동해 이 집합에서 빠지고, 사용자가 그 슬라이드를 다시 클릭하면 재시도된다.
const pendingSlideSubscriptions = new Set<string>();
// pendingSlideSubscriptions 항목 하나의 최대 수명(ms) — 정확할 필요는 없고, 정상적인 왕복
// 시간보다 넉넉히 길어서 성공 케이스에서는 절대 발동하지 않을 정도면 된다.
const PENDING_SUBSCRIPTION_TIMEOUT_MS = 9000;

// 복제(duplicate)로 만들어진 슬라이드 entryId -> 그 원본(source) 슬라이드 entryId. 복제 버튼
// 핸들러가 선제적으로 duplicateOfId를 실어 구독 요청을 한 번 보내지만(아래 setupCardActions
// 참고), 그 시점에 WebSocket이 OPEN이 아니면(재연결 중 등) subscribeSlideContent는 아무것도
// 보내지 않고 조용히 리턴한다 — 그 경우 그 복제 슬라이드에 대해 duplicateOfId를 실어 보내는
// 유일한 시도가 사라져 버리고, 이후 renderCanvasForSlide의 "아직 안 열렸다" 폴백 경로가 보통의
// (duplicateOfId 없는) 구독만 보내게 되어 서버가 소스의 살아있는 편집 내용 대신 오래된 디스크
// XML로 시딩해 버린다. 이 맵은 그 간극을 구조적으로 막는다 — subscribeSlideContent가
// duplicateOfId 없이 불릴 때마다 여기서 먼저 찾아본다(아래 참고). 실제로 시딩이 끝나면(=
// slideContents.observe가 그 키를 감지하면) 더 이상 필요 없으니 지운다.
const pendingDuplicateSources = new Map<string, string>();

function subscribeSlideContent(entryId: string, duplicateOfId?: string): void {
  if (!provider.ws || provider.ws.readyState !== WebSocket.OPEN) return;
  if (pendingSlideSubscriptions.has(entryId)) return; // 이미 이 슬라이드에 대한 요청이 응답을 기다리는 중
  // 이 호출이 duplicateOfId를 명시적으로 안 받았다면(예: renderCanvasForSlide의 폴백 경로,
  // 또는 위 타임아웃으로 풀려난 뒤의 재시도) 이 슬라이드가 복제본으로 알려져 있는지 먼저
  // 확인한다 — 있다면 그 원본 id를 그대로 duplicateOfId로 써서, 어느 경로로 결국 요청이
  // 성공하든 항상 올바른 소스에서 시딩되게 한다.
  const effectiveDuplicateOfId = duplicateOfId ?? pendingDuplicateSources.get(entryId);
  pendingSlideSubscriptions.add(entryId);
  setTimeout(() => pendingSlideSubscriptions.delete(entryId), PENDING_SUBSCRIPTION_TIMEOUT_MS);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SLIDE_CONTENT_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ slideEntryId: entryId, duplicateOfId: effectiveDuplicateOfId }));
  provider.ws.send(encoding.toUint8Array(encoder));
}

// entryId(슬라이드) -> 그 슬라이드의 shapeId -> 마운트된 EditorView. 슬라이드를 벗어나면(다른
// 슬라이드를 클릭하면) 여기 담긴 뷰를 전부 destroy()한 뒤 비운다 — 그러지 않으면 이미 화면에서
// 사라진 슬라이드의 EditorView가 계속 ySyncPlugin으로 Y.Doc을 관찰하며 메모리/리스너가 샌다.
const mountedViews = new Map<string, EditorView>();

// 지금 포커스를 들고 있는(=가장 최근에 focus 이벤트를 받은) EditorView. 서식 툴바(굵게/기울임/
// 밑줄/색)가 "어느 텍스트박스에 적용할지" 알아야 하는데, 슬라이드 하나에 텍스트박스(=EditorView)가
// 여러 개 동시에 마운트돼 있어 이 값 없이는 알 수 없다 — 각 view 생성 시 handleDOMEvents.focus로
// 갱신한다(아래 renderCanvasForSlide 참고). 툴바 버튼을 클릭하면 그 클릭이 현재 포커스된
// contentEditable을 blur시키지만, ProseMirror의 state.selection은 DOM blur와 무관하게 그대로
// 남아 있으므로(editor.ts의 bindToolbarButton과 동일한, 이미 검증된 가정) blur 이벤트에서 이 값을
// null로 되돌리지 않는다 — 그러면 버튼 클릭 시점엔 이미 focusedView가 null이 되어 아무 데도 적용이
// 안 된다. 유일한 예외는 renderCanvasForSlide가 이 view를 destroy할 때(아래에서 별도로 null 처리).
let focusedView: EditorView | null = null;

// 지금 활성 슬라이드의 perSlide(Y.Map<shapeId, shapeMap>)에 붙여 둔 구조적 변경 구독 해제 함수.
// perSlide 자기 자신에 대한 얕은(shallow) observe는 shapeId 키가 추가/삭제될 때만(=텍스트박스
// 추가/삭제) 올라오고, 텍스트박스 "안" 내용 편집(타이핑)은 그보다 훨씬 깊은 Y.XmlFragment
// 레벨에서 일어나 여기로 전혀 올라오지 않는다 — 그래서 이 구독을 renderCanvasForSlide 재호출
// 트리거로 써도 키 입력마다 EditorView가 destroy/재생성되는 일은 없다. slideContents 자체의
// 최상위 observe(main() 참고)는 "슬라이드가 아직 시딩 전이냐"만 보므로, 이미 시딩된 슬라이드
// 안에서 텍스트박스가 추가/삭제되는 것(이 태스크의 핵심 기능)은 그 최상위 observe로는 잡히지
// 않는다 — 그래서 이 별도의 perSlide 레벨 구독이 필요하다. renderCanvasForSlide가 불릴 때마다
// (슬라이드 전환이든, 이 구독 자신이 재호출한 것이든) 맨 위에서 이전 구독을 먼저 해제하고 새로
// 붙이므로, 활성 슬라이드가 아닌 슬라이드에 대한 구독이 남아있는 일은 없다.
let activePerSlideUnobserve: (() => void) | null = null;

/**
 * y-prosemirror의 undo/redo를 실제 키(Ctrl+Z/Ctrl+Shift+Z, Ctrl+Y)에 연결한다. 이 keymap
 * 없이는 yUndoPlugin()이 만드는 Y.UndoManager가 순전히 죽은 코드다 — 어떤 키도 그 undo/redo
 * 커맨드를 부르지 않으니 Ctrl+Z를 누르면 브라우저 자신의 contentEditable 네이티브 undo가
 * 대신 실행되고, 그 결과 되돌아간 DOM을 ProseMirror가 "새 입력"으로 다시 읽어들여 모든
 * 협업자에게 진짜 편집처럼 전파해 버린다 — 그 과정에서 다른 협업자가 동시에 넣은, 전혀 무관한
 * 편집을 조용히 되돌릴 수 있다. editor.ts(docx 편집기)의 buildKeymap과 같은 패턴 — baseKeymap
 * 위에 Mod-z/Mod-y/Mod-Shift-z만 y-prosemirror의 undo/redo로 덮어쓴다(이 스키마엔 표/목록/
 * 제목이 없어 그 외 docx 전용 바인딩은 옮기지 않는다).
 */
function buildSlideKeymap(): Record<string, Command> {
  return {
    ...baseKeymap,
    'Mod-z': undo,
    'Mod-y': redo,
    'Mod-Shift-z': redo,
  };
}

/**
 * 읽기 전용 열람자가 "아직 아무도 열어본 적 없는" 슬라이드를 볼 때 캔버스에 보여줄 안내.
 * (Task 9 리뷰에서 지적된 항목 — canEdit 세션은 subscribeSlideContent로 서버 지연 시딩을
 * 요청할 수 있지만, 읽기 전용 연결은 room.ts가 그 요청 자체를 서버에서 거부하므로 응답이 영영
 * 오지 않는다. 그 상태를 아무 설명 없는 빈 캔버스로 방치하지 않고, 편집자가 먼저 열어야 한다는
 * 것을 명시적으로 알린다. 원본 슬라이드 내용/이미지를 렌더링하는 것은 이 기능 전체의 범위 밖이라
 * 시도하지 않는다.) canvasEl.replaceChildren()이 이미 이전 내용을 비운 뒤 호출되므로 여기서는
 * 추가만 한다.
 */
function showSlideUnavailablePlaceholder(): void {
  const placeholder = document.createElement('div');
  placeholder.className = 'slide-canvas-placeholder';
  placeholder.textContent = '이 슬라이드는 아직 미리보기를 사용할 수 없습니다 (편집자가 먼저 열어야 합니다).';
  canvasEl.append(placeholder);
}

/**
 * 오른쪽 캔버스에 activeSlideId 슬라이드의 텍스트박스를 실제 위치로 배치하고, 각 텍스트박스에
 * pptxTextSchema 기반 ProseMirror EditorView를 마운트한다(y-prosemirror 바인딩은 editor.ts의
 * 기존 docx 편집기 마운트 코드와 동일한 패턴 — ySyncPlugin/yCursorPlugin/yUndoPlugin).
 *
 * 슬라이드가 아직 한 번도 열린 적이 없으면(slideContents에 그 entryId 키가 없으면) 서버에 지연
 * 시딩을 요청만 하고 빈 캔버스로 돌아온다 — 실제 콘텐츠는 slideContents.observe가 시딩 완료를
 * 감지해 이 함수를 다시 불러줄 때 그려진다(main()의 observe 콜백 참고).
 */
function renderCanvasForSlide(entryId: string): void {
  // y-prosemirror(1.3.7)의 ySyncPlugin/yCursorPlugin은 원격 갱신·커서 이동이 있을 때마다
  // "이 view에 적용할 setMeta"를 모듈 전역 큐에 넣고 setTimeout(0)으로 나중에 한꺼번에
  // flush한다(y-prosemirror 소스의 updateMetas). updateMetas는 flush 직전
  // `syncState.binding.isDestroyed`를 확인하긴 하지만, 라이브러리 자신은 destroy() 안
  // 어디서도 그 플래그를 true로 설정하지 않는다(node_modules/y-prosemirror/dist/
  // y-prosemirror.cjs의 ProsemirrorBinding.prototype.destroy와 lib.js:29 확인 완료 — 순수
  // 업스트림 버그, 죽은 코드나 마찬가지다). 그래서 destroy()만 부르면, 이미 예약돼 있던 flush가
  // 나중에 실행되며 죽은 view에 대고 dispatch를 시도해 "Cannot read properties of null
  // (reading 'matchesNode')" 예외를 던진다 — 두 탭이 동시에 타이핑하는 중 슬라이드를 전환하는
  // 시나리오로 재현 확인했다. destroy() 호출을 한 틱 미루는 방식은 그 창을 좁힐 뿐 완전히
  // 닫지는 못한다 — 미뤄둔 그 짧은 구간에도 원격 awareness 갱신(커서 이동 등)이 도착하면
  // setMeta가 다시 예약되어 같은 충돌이 그대로 재현될 수 있다. 실제 수정은 라이브러리가
  // 빠뜨린 그 플래그를 우리가 destroy() 호출 "직전"에 직접 true로 설정해 두는 것이다 —
  // view.state(그리고 그 안의 ySyncPlugin state가 들고 있는 binding 객체)는 destroy() 이후에도
  // 참조가 살아 있으므로, 이렇게 하면 언제 예약된 flush든 실행 시점에 항상 이 가드에 걸려
  // 안전하게 no-op된다(타이밍에 기대지 않는 결정적 해결).
  const previousViews = [...mountedViews.values()];
  for (const view of previousViews) {
    const syncState = ySyncPluginKey.getState(view.state) as { binding?: { isDestroyed?: boolean } } | undefined;
    if (syncState?.binding) syncState.binding.isDestroyed = true;
    // focusedView가 지금 destroy하려는 view를 가리키고 있다면 반드시 여기서 null로 되돌린다 —
    // 그러지 않으면 (예: 이 텍스트박스에 커서를 두고 있다가 다른 슬라이드 카드를 클릭한 경우)
    // 서식 툴바가 이미 destroy된(= docView가 사라진) view에 대고 dispatch를 시도해 예외를 던진다.
    if (focusedView === view) focusedView = null;
    view.destroy();
  }
  mountedViews.clear();
  // 이전에 활성이던 슬라이드의 perSlide 구조 구독을 해제한다 — 그러지 않으면 다른 슬라이드로
  // 옮긴 뒤에도 그 슬라이드에 텍스트박스가 추가/삭제될 때마다(원격 협업자가 하는 것 포함) 이미
  // 화면에 없는 슬라이드를 위해 불필요하게 이 함수가 다시 불린다.
  activePerSlideUnobserve?.();
  activePerSlideUnobserve = null;
  canvasEl.style.width = `${emuToPx(SLIDE_W_EMU)}px`;
  canvasEl.style.height = `${emuToPx(SLIDE_H_EMU)}px`;
  canvasEl.replaceChildren();

  const perSlide = slideContents.get(entryId);
  if (!perSlide) {
    if (canEdit) {
      subscribeSlideContent(entryId); // 아직 안 열렸다 — 요청하고 observe가 다시 불러줄 때까지 대기
    } else {
      // 읽기 전용 연결은 시딩(=Y.Doc 쓰기)을 트리거하는 구독 요청을 room.ts가 서버에서 직접
      // 거부한다(MESSAGE_SLIDE_CONTENT_SUBSCRIBE 핸들러의 readOnlyConns 체크) — 그러니 보내봤자
      // 응답이 오지 않을 요청을 굳이 보내지 않는다. 편집자가 먼저 이 슬라이드를 열어 시딩해 줄
      // 때까지 이 뷰어는 이유를 알 수 있는 안내를 본다 — 나중에 다른 세션의 편집자가 이 슬라이드를
      // 열면 slideContents.observe(main() 참고)가 이 함수를 다시 불러 정상적으로 대체된다.
      showSlideUnavailablePlaceholder();
    }
    return;
  }
  // perSlide는 이미 존재하지만(슬라이드가 한 번은 시딩됨) 그 안에서 텍스트박스가 나중에
  // 추가/삭제될 수 있다("텍스트 추가" 버튼, 삭제 버튼, 또는 원격 협업자의 같은 조작) — 그 구조적
  // 변경(=perSlide 자신의 키 추가/삭제)을 감지해 다시 그린다. 얕은 observe라 텍스트박스 "안"
  // 내용 편집(타이핑)에는 반응하지 않는다(위 activePerSlideUnobserve 선언부 주석 참고).
  const onPerSlideStructureChange = () => renderCanvasForSlide(entryId);
  perSlide.observe(onPerSlideStructureChange);
  activePerSlideUnobserve = () => perSlide.unobserve(onPerSlideStructureChange);

  for (const [shapeId, shapeMap] of perSlide.entries()) {
    const meta = shapeMap.get('meta') as SlideShapeMeta;
    const box = document.createElement('div');
    box.className = 'slide-textbox';
    box.dataset.shapeId = shapeId; // 디버깅/테스트에서 개별 텍스트박스를 식별하기 위함(동작에는 영향 없음)
    box.style.left = `${emuToPx(meta.x ?? SLIDE_W_EMU * 0.2)}px`;
    box.style.top = `${emuToPx(meta.y ?? SLIDE_H_EMU * 0.4)}px`;
    box.style.width = `${emuToPx(meta.w ?? SLIDE_W_EMU * NEW_TEXTBOX_WIDTH_RATIO)}px`;
    box.style.height = `${emuToPx(meta.h ?? SLIDE_H_EMU * NEW_TEXTBOX_HEIGHT_RATIO)}px`;
    canvasEl.append(box);

    if (canEdit) {
      // 호버 시에만 보이는 작은 × 삭제 버튼(CSS: .slide-textbox-delete). perSlide.delete는
      // 위에서 이 슬라이드에 붙여 둔 onPerSlideStructureChange 구독을 트리거해 renderCanvasForSlide를
      // 다시 부르므로, 남은 텍스트박스 전체가 destroy→재생성되는 이미 검증된 기존 경로(이 함수 맨
      // 위의 previousViews 정리)를 그대로 타 삭제된 shapeId의 EditorView도 함께 정리된다 — 별도
      // cleanup 로직을 새로 만들 필요가 없다.
      const deleteBtn = document.createElement('button');
      deleteBtn.type = 'button';
      deleteBtn.className = 'slide-textbox-delete';
      deleteBtn.textContent = '×';
      deleteBtn.title = '텍스트 상자 삭제';
      deleteBtn.setAttribute('aria-label', '텍스트 상자 삭제');
      deleteBtn.addEventListener('click', (event) => {
        event.stopPropagation();
        const ydoc = slidesArray.doc!;
        ydoc.transact(() => {
          perSlide.delete(shapeId);
        });
      });
      box.append(deleteBtn);
    }

    const fragment = shapeMap.get('content') as Y.XmlFragment;
    const view = new EditorView(box, {
      state: EditorState.create({
        schema: pptxTextSchema,
        plugins: [ySyncPlugin(fragment), yCursorPlugin(provider.awareness), yUndoPlugin(), keymap(buildSlideKeymap())],
      }),
      editable: () => canEdit,
      // 서식 툴바가 "지금 어느 텍스트박스를 편집 중인지" 알 수 있는 유일한 경로 — handleDOMEvents의
      // 반환값 false는 "이 이벤트를 ProseMirror가 가로채 막지 말고 브라우저 기본 동작(포커스 이동)을
      // 그대로 진행시켜라"는 뜻이다(true를 반환하면 기본 동작이 막힌다). focus 이벤트는 이 view가
      // DOM에 붙은 뒤 나중에(사용자가 클릭하거나 명시적으로 .focus()를 부를 때)만 발생하므로, 이
      // 콜백 안에서 아직 대입 중인 const view를 참조해도(클로저) 실제 호출 시점엔 이미 대입이 끝나
      // 있어 안전하다 — editor.ts가 dispatchTransaction에서 겪은 "생성자 안에서 view 참조" 문제와
      // 달리, focus는 생성자 실행 중 동기적으로 발생하지 않는다.
      handleDOMEvents: {
        focus: () => {
          focusedView = view;
          return false;
        },
      },
    });
    mountedViews.set(shapeId, view);
  }
}

function cardFor(entry: SlideEntry, index: number): HTMLElement {
  const fragment = cardTemplate.content.cloneNode(true) as DocumentFragment;
  const card = fragment.querySelector('.slide-card') as HTMLElement;
  card.dataset.index = String(index);
  card.dataset.id = entry.id;
  const img = card.querySelector('img') as HTMLImageElement;
  img.src =
    entry.kind === 'blank'
      ? '/img/blank-slide.svg'
      : `/api/view/pages/${entry.sourcePageNumber}?token=${encodeURIComponent(token ?? '')}`;
  (card.querySelector('.slide-card-number') as HTMLElement).textContent = String(index + 1);
  if (!canEdit) {
    card.querySelector('.slide-card-actions')?.remove();
    card.draggable = false;
  }
  return card;
}

function renderSlides(): void {
  const entries = slidesArray.toArray();
  railEl.replaceChildren(...entries.map((entry, index) => cardFor(entry, index)));
  // 아직 선택된 슬라이드가 없거나(최초 로드), 선택돼 있던 슬라이드가 방금 삭제됐다면 첫 번째
  // 슬라이드로 선택을 옮긴다 — 그러지 않으면 삭제 후 레일에 활성 표시가 하나도 안 남는다.
  const activeStillExists = entries.some((entry) => entry.id === activeSlideId);
  if (!activeStillExists) {
    if (entries.length > 0) {
      selectSlide(entries[0].id);
      return; // selectSlide()가 아래와 동일한 active 클래스 갱신을 이미 하고 끝났다
    }
    activeSlideId = null;
  }
  for (const card of railEl.querySelectorAll('.slide-card')) {
    card.classList.toggle('active', (card as HTMLElement).dataset.id === activeSlideId);
  }
}

// 슬라이드 카드 클릭(또는 렌더 직후 기본 선택)으로 오른쪽 캔버스에 표시할 슬라이드를 바꾼다.
// 실제 콘텐츠 렌더링은 onActiveSlideChange 콜백(renderCanvasForSlide)이 처리한다 — 여기서는 상태/표시만.
function selectSlide(entryId: string): void {
  // 이미 활성 상태인 슬라이드를 다시 클릭한 경우 — 아무것도 바뀌지 않았으니 다시 그리지 않는다.
  // renderCanvasForSlide는 호출될 때마다 그 슬라이드의 모든 EditorView를 destroy→재생성하므로,
  // 이 가드가 없으면 재클릭할 때마다 편집 중이던 커서 위치/선택 영역이 이유 없이 날아간다.
  if (entryId === activeSlideId) return;
  activeSlideId = entryId;
  for (const card of railEl.querySelectorAll('.slide-card')) {
    card.classList.toggle('active', (card as HTMLElement).dataset.id === entryId);
  }
  onActiveSlideChange?.(entryId);
}

function setupDragAndDrop(): void {
  railEl.addEventListener('dragstart', (event) => {
    const card = (event.target as HTMLElement).closest<HTMLElement>('.slide-card');
    if (!card) return;
    dragFromIndex = Number(card.dataset.index);
    event.dataTransfer?.setData('text/plain', card.dataset.id ?? '');
  });
  railEl.addEventListener('dragover', (event) => {
    if (dragFromIndex !== null) event.preventDefault();
  });
  railEl.addEventListener('drop', (event) => {
    event.preventDefault();
    const target = (event.target as HTMLElement).closest<HTMLElement>('.slide-card');
    if (!target || dragFromIndex === null) return;
    const toIndex = Number(target.dataset.index);
    if (toIndex === dragFromIndex) return;
    const ydoc = slidesArray.doc!;
    ydoc.transact(() => {
      const [moved] = slidesArray.slice(dragFromIndex!, dragFromIndex! + 1);
      slidesArray.delete(dragFromIndex!, 1);
      // 삭제 후 배열이 한 칸 줄어들므로, 드롭 대상이 삭제 지점보다 뒤에 있었다면 그 대상의
      // 실제(삭제 후) 위치는 원래 toIndex보다 1 작다 — 항상 "대상 카드 바로 앞"에 꽂히도록
      // 두 경우 모두 대상의 삭제-후 위치에 삽입한다.
      const insertAt = toIndex > dragFromIndex! ? toIndex - 1 : toIndex;
      slidesArray.insert(insertAt, [moved]);
    });
    dragFromIndex = null;
  });
}

function setupCardActions(): void {
  railEl.addEventListener('click', (event) => {
    const target = event.target as HTMLElement;
    const card = target.closest<HTMLElement>('.slide-card');
    if (!card) return;
    // 복제/삭제 버튼이 아니라 카드 자체(썸네일 등)를 클릭했다면 선택만 하고 끝난다 — 읽기
    // 전용 열람자도 슬라이드를 넘겨 보되(canEdit 게이트 없이), 복제/삭제 버튼은 어차피
    // cardFor()에서 이미 DOM에서 제거돼 있어 아래 분기로는 내려오지 않는다.
    if (!target.closest('.slide-card-duplicate') && !target.closest('.slide-card-delete')) {
      selectSlide(card.dataset.id!);
      return;
    }
    if (!canEdit) return;
    const index = Number(card.dataset.index);
    const entry = slidesArray.get(index);
    if (target.closest('.slide-card-duplicate')) {
      const newEntryId = crypto.randomUUID();
      // 아래 proactive subscribeSlideContent 호출이 그 순간 WebSocket이 OPEN이 아니라서(재연결
      // 중 등) 아무것도 안 보내고 조용히 리턴하더라도, 이 매핑은 이미 존재하게 만들어 둔다 —
      // 그래야 나중에 어떤 경로로든(renderCanvasForSlide의 폴백, 또는 pendingSlideSubscriptions
      // 타임아웃 이후의 재시도) 이 슬라이드에 대한 구독이 결국 시도될 때 duplicateOfId 없는
      // "평범한" 요청이 나가는 대신 항상 이 원본 id를 함께 실어 보낼 수 있다.
      pendingDuplicateSources.set(newEntryId, entry.id);
      const ydoc = slidesArray.doc!;
      ydoc.transact(() => {
        slidesArray.insert(index + 1, [
          {
            id: newEntryId,
            kind: 'duplicate',
            sourceSlidePart: entry.sourceSlidePart,
            sourcePageNumber: null,
          },
        ]);
      });
      // 서버가 원본 파일의(오래된) 디스크 XML이 아니라 소스 슬라이드의 "지금 세션에서 편집된"
      // 콘텐츠를 그대로 깊은 복사해 시딩하도록, 사용자가 실제로 이 복제본을 열기 전에(활성
      // 슬라이드가 되는 것과 무관하게) 미리 duplicateOfId를 실어 구독 요청을 보내 둔다. 이 전송
      // 자체가 드롭돼도(WS가 그 순간 OPEN이 아니었던 경우) 위에서 이미 채워둔
      // pendingDuplicateSources 덕분에, 이후 이 슬라이드에 대해 실제로 나가는 첫 구독 요청은
      // subscribeSlideContent 내부에서 이 매핑을 찾아 duplicateOfId를 채워 보낸다.
      subscribeSlideContent(newEntryId, entry.id);
    } else if (target.closest('.slide-card-delete')) {
      if (slidesArray.length <= 1) return; // 마지막 남은 슬라이드는 지울 수 없다
      slidesArray.delete(index, 1);
    }
  });
  document.getElementById('btn-add-blank')?.addEventListener('click', () => {
    if (!canEdit) return;
    slidesArray.insert(slidesArray.length, [
      { id: crypto.randomUUID(), kind: 'blank', sourceSlidePart: null, sourcePageNumber: null },
    ]);
  });
  document.getElementById('btn-add-textbox')?.addEventListener('click', () => {
    if (!canEdit || !activeSlideId) return;
    const entryId = activeSlideId;
    const ydoc = slidesArray.doc!;
    ydoc.transact(() => {
      // 슬라이드가 완전히 빈 상태(원본이 blank로 만들어졌거나, 아직 시딩되지 않은 것과는 다르게
      // canEdit 세션이라면 애초에 subscribeSlideContent로 이미 시딩됐어야 하지만, 혹시라도 정말
      // shapeless한 슬라이드라면) slideContents에 이 슬라이드의 perSlide 항목 자체가 없을 수
      // 있다 — 그 경우 여기서 새로 만든다.
      let perSlide = slideContents.get(entryId);
      if (!perSlide) {
        perSlide = new Y.Map<ShapeContentMap>();
        slideContents.set(entryId, perSlide);
      }
      const shapeId = `new-${crypto.randomUUID()}`;
      const shapeMap: ShapeContentMap = new Y.Map();
      const meta: SlideShapeMeta = {
        shapeId,
        x: Math.round((SLIDE_W_EMU * (1 - NEW_TEXTBOX_WIDTH_RATIO)) / 2),
        y: Math.round((SLIDE_H_EMU * (1 - NEW_TEXTBOX_HEIGHT_RATIO)) / 2),
        w: Math.round(SLIDE_W_EMU * NEW_TEXTBOX_WIDTH_RATIO),
        h: Math.round(SLIDE_H_EMU * NEW_TEXTBOX_HEIGHT_RATIO),
      };
      shapeMap.set('meta', meta);
      shapeMap.set('content', new Y.XmlFragment());
      perSlide.set(shapeId, shapeMap);
    });
    // perSlide가 이미 있었다면 위 변경은 slideContents 자신의 키가 아니라 그 안의 perSlide
    // Y.Map에 대한 나머지 변경이라 최상위 slideContents.observe(main() 참고)로는 안 올라온다 —
    // 하지만 이 슬라이드가 지금 활성 상태였다면(대부분의 경우) renderCanvasForSlide가 이미 이
    // perSlide에 onPerSlideStructureChange 구독을 붙여 뒀으므로(위 renderCanvasForSlide 참고)
    // 그 구독이 알아서 다시 그려준다 — perSlide를 방금 새로 만든 경우(원래 shapeless)만
    // slideContents.set이 최상위 키 변경이라 그 observe가 대신 다시 그려준다. 두 경로 모두
    // 여기서 별도로 renderCanvasForSlide를 부를 필요가 없다.
  });
}

/**
 * 서식 툴바(굵게/기울임/밑줄/색) 버튼을 지금 focusedView에 잡혀 있는 EditorView에 연결한다.
 * editor.ts(docx 편집기)의 bindToolbarButton/font-color 배선과 같은 패턴 — 일반 click 이벤트를
 * 쓰고(mousedown+preventDefault가 아니다) 커맨드 적용 뒤 view.focus()로 포커스를 되돌린다. 이
 * 패턴이 안전한 이유: 버튼 mousedown이 지금 포커스된 contentEditable을 blur시키더라도
 * ProseMirror의 EditorState.selection은 DOM 포커스와 별개로 유지되는 순수 JS 상태라 click
 * 핸들러가 실행되는 시점에도 그대로 남아 있다 — focusedView는 focus 이벤트에서만 갱신하고 blur
 * 에서는 절대 건드리지 않으므로(renderCanvasForSlide 위쪽 주석 참고) 이 사이에도 값을 잃지 않는다.
 */
function bindFormatButton(id: string, markName: 'strong' | 'em' | 'underline'): void {
  document.getElementById(id)?.addEventListener('click', () => {
    if (!focusedView) return; // 포커스된 텍스트박스가 없으면 조용히 아무 것도 하지 않는다.
    toggleMark(pptxTextSchema.marks[markName])(focusedView.state, focusedView.dispatch);
    focusedView.focus();
  });
}

function setupFormatToolbar(): void {
  bindFormatButton('btn-bold', 'strong');
  bindFormatButton('btn-italic', 'em');
  bindFormatButton('btn-underline', 'underline');

  const colorInput = document.getElementById('text-color-picker') as HTMLInputElement | null;
  const colorBtn = document.getElementById('text-color-btn') as HTMLElement | null;
  // .color-btn-bar는 --btn-color 커스텀 프로퍼티로 자기 배경색을 그리므로(style.css), 입력의
  // 현재 값과 항상 맞춰 둔다 — 초기값(HTML의 value="#1b1a17")과 이후 사용자가 고른 색 둘 다.
  colorBtn?.style.setProperty('--btn-color', colorInput?.value ?? '#1b1a17');
  colorInput?.addEventListener('input', () => {
    if (!focusedView) return;
    colorBtn?.style.setProperty('--btn-color', colorInput.value);
    const { from, to } = focusedView.state.selection;
    // from === to(선택 영역 없이 커서만 있는 경우)여도 addMark(from, to, ...)는 그냥 빈 범위에
    // 아무것도 안 하는 no-op이다 — toggleMark류처럼 "커서 위치의 다음 입력에 적용"하는 stored-mark
    // 동작까지는 흉내 내지 않지만, 이 태스크 범위에서는 브리프가 그린 그대로(선택 영역에 직접
    // addMark) 충분하다.
    const tr = focusedView.state.tr.addMark(from, to, pptxTextSchema.marks.textColor.create({ color: colorInput.value }));
    focusedView.dispatch(tr);
    focusedView.focus();
  });
}

async function main(): Promise<void> {
  if (!token) {
    showFatalError('열람 토큰이 없습니다.');
    return;
  }
  let info: ViewInfo;
  try {
    info = await fetchViewInfo();
  } catch (error) {
    showFatalError(error instanceof Error ? error.message : '문서 정보를 불러오지 못했습니다.');
    return;
  }
  titleEl.textContent = info.name;
  document.title = `${info.name} — DocView 편집`;
  canEdit = info.permissions.canEdit;
  // 읽기 전용 토큰이어도 이 화면 자체는 그대로 보여주고 쓰기만 막는다 — grid-editor.ts와 동일한
  // 패턴(canEdit:false로 fatal error를 띄우면 "바로보기"가 동시편집 화면을 못 씀). 실제 쓰기
  // 차단은 room.ts의 readOnlyConns(서버 쪽)가 이미 하고 있고, 여기서는 UI만 맞춘다.
  if (!canEdit) {
    document.getElementById('btn-add-blank')?.setAttribute('hidden', '');
    document.getElementById('btn-add-textbox')?.setAttribute('hidden', '');
    document.getElementById('format-toolbar')?.setAttribute('hidden', '');
    connStatusEl.insertAdjacentHTML('afterend', '<span class="conn-status" data-status="readonly">읽기 전용</span>');
  }

  const ydoc = new Y.Doc();
  slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
  slideContents = ydoc.getMap<PerSlideContentsMap>(PPTX_SLIDE_CONTENTS_MAP);
  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, {
    params: { token },
  });

  const myColor = pickPresenceColorForName(info.viewer);
  provider.awareness.setLocalStateField('user', { name: info.viewer, color: myColor });

  provider.on('status', ({ status }: { status: string }) => {
    connStatusEl.dataset.status = status;
    connStatusEl.textContent = status === 'connected' ? '연결됨' : '연결 중…';
  });
  // 서버가 편집 세션 자체를 거부한 경우(edit.ts의 socket.close(4xxx, 이유)) — 4400대는 재시도해도
  // 같은 이유로 다시 거부될 뿐인 영구 거부다(잘못된 토큰, 문서 없음 등). 기본 동작(지수 백오프로
  // 계속 재연결 시도)을 멈추고 서버가 준 이유를 그대로 보여준다 — grid-editor.ts와 동일한 처리.
  provider.on('connection-close', (event: CloseEvent | null) => {
    if (event && event.code >= 4400) {
      provider.shouldConnect = false;
      showFatalError(event.reason || '서버가 연결을 거부했습니다.');
    }
  });
  provider.awareness.on('change', () => renderPresenceAvatars(presenceListEl, provider.awareness));

  // slideContents는 슬라이드별 키가 지연 시딩될 때만 top-level 변경이 일어난다(텍스트박스 안
  // 내용 편집은 그 nested Y.Map/Y.XmlFragment 안에서만 바뀌므로 여기로 이벤트가 올라오지 않는다
  // — 그건 각 텍스트박스에 마운트된 ySyncPlugin이 알아서 반영한다). event.keysChanged로 지금
  // 보고 있는 슬라이드의 키가 실제로 추가된 경우에만 다시 그린다 — 그러지 않고 무조건 다시
  // 그리면, 다른 슬라이드가 시딩될 때마다(다른 참여자가 다른 슬라이드를 여는 경우 등) 지금 이미
  // 마운트돼 편집 중인 EditorView까지 destroy→재생성돼 커서/포커스가 끊기고 만다.
  slideContents.observe((event) => {
    // 새로 채워진(=서버가 시딩을 끝낸) 슬라이드는 더 이상 "응답 대기 중"이 아니므로 in-flight
    // 집합에서 지운다 — 그러지 않으면 그 슬라이드에 대한 이후의 정당한 subscribeSlideContent
    // 호출이 (실제로는 이미 끝난) 이전 요청 때문에 영원히 무시된다. 같은 이유로
    // pendingDuplicateSources에서도 지운다 — 시딩이 실제로 끝난 이상 그 매핑은 더 이상 쓸모가
    // 없고(서버는 이미 채워진 슬라이드를 다시 덮어쓰지 않으므로 재시딩은 원래 일어나지 않지만),
    // 남겨두면 훗날 어떤 경로로 같은 entryId가 재사용되는 예외적인 경우에 오래된 매핑이 다시
    // 쓰일 위험이 있어 위생상 정리한다.
    for (const key of event.keysChanged) {
      pendingSlideSubscriptions.delete(key);
      pendingDuplicateSources.delete(key);
    }
    if (activeSlideId && event.keysChanged.has(activeSlideId)) renderCanvasForSlide(activeSlideId);
  });
  // Task 8이 남겨 둔 훅 — 슬라이드를 고를 때마다 오른쪽 캔버스를 그 슬라이드 내용으로 채운다.
  onActiveSlideChange = renderCanvasForSlide;

  slidesArray.observe(renderSlides);
  renderSlides();
  setupDragAndDrop();
  setupCardActions();
  setupFormatToolbar();

  // ---------- 저장 상태 표시(Office 365 스타일 "저장 중… / 저장됨") + Ctrl+S ----------
  // grid-editor.ts와 완전히 동일한 패턴 — docMeta.savedAt은 서버의 flushToStore가 실제로
  // 디스크에 쓴 뒤에만 갱신하는 일반 Y.Map 값이라(gridShared.ts, 문서 종류 공통), 이 pptx
  // 편집기도 별도 서버 변경 없이 그대로 재사용할 수 있다.
  const docMeta = ydoc.getMap<number>(DOC_META_MAP);
  const saveStatusEl = document.getElementById('save-status') as HTMLElement | null;
  const saveStatusTextEl = saveStatusEl?.querySelector('.save-status-text') as HTMLElement | null;
  let lastKnownSavedAt = docMeta.get(DOC_META_SAVED_AT_KEY) ?? 0;
  let lastKnownSaveErrorAt = docMeta.get(DOC_META_SAVE_ERROR_AT_KEY) ?? 0;
  function setSaveStatus(status: 'saving' | 'saved' | 'error'): void {
    if (!saveStatusEl) return;
    saveStatusEl.hidden = false;
    saveStatusEl.dataset.status = status;
    if (saveStatusTextEl) {
      saveStatusTextEl.textContent =
        status === 'saving' ? '저장 중…' : status === 'saved' ? '모든 변경 내용이 저장됨' : '저장 실패 — 편집을 계속하면 다시 시도합니다';
    }
  }
  if (canEdit) {
    ydoc.on('afterTransaction', (tx) => {
      if (!tx.local || tx.changed.size === 0) return;
      setSaveStatus('saving');
    });
  }
  // saveErrorAt이 savedAt보다 최신이면 아직 실패가 해소되지 않은 것 — 서버가 다음 저장에
  // 성공해 savedAt이 다시 앞서는 순간 별도 정리 없이도 자연스럽게 실패 표시가 사라진다.
  docMeta.observe(() => {
    const savedAt = docMeta.get(DOC_META_SAVED_AT_KEY) ?? 0;
    const saveErrorAt = docMeta.get(DOC_META_SAVE_ERROR_AT_KEY) ?? 0;
    if (saveErrorAt > savedAt && saveErrorAt > lastKnownSaveErrorAt) {
      lastKnownSaveErrorAt = saveErrorAt;
      setSaveStatus('error');
    } else if (savedAt > lastKnownSavedAt) {
      lastKnownSavedAt = savedAt;
      setSaveStatus('saved');
    }
  });
  if (canEdit) {
    window.addEventListener('keydown', (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        setSaveStatus('saving');
        fetch(`/api/view/save?token=${encodeURIComponent(token)}`, { method: 'POST' }).catch(() => {
          // 네트워크 오류 등으로 실패해도 8초 유휴 자동저장이 알아서 다시 시도하므로 별도 안내는 생략한다.
        });
      }
    });
  }
}

void main();
