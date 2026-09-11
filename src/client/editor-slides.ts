import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import * as encoding from 'lib0/encoding';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { ySyncPlugin, yCursorPlugin, yUndoPlugin } from 'y-prosemirror';
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
function subscribeSlideContent(entryId: string, duplicateOfId?: string): void {
  if (!provider.ws || provider.ws.readyState !== WebSocket.OPEN) return;
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SLIDE_CONTENT_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ slideEntryId: entryId, duplicateOfId }));
  provider.ws.send(encoding.toUint8Array(encoder));
}

// entryId(슬라이드) -> 그 슬라이드의 shapeId -> 마운트된 EditorView. 슬라이드를 벗어나면(다른
// 슬라이드를 클릭하면) 여기 담긴 뷰를 전부 destroy()한 뒤 비운다 — 그러지 않으면 이미 화면에서
// 사라진 슬라이드의 EditorView가 계속 ySyncPlugin으로 Y.Doc을 관찰하며 메모리/리스너가 샌다.
const mountedViews = new Map<string, EditorView>();

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
  // 곧바로 view.destroy()를 부르지 않고 한 틱(매크로태스크) 미룬다. y-prosemirror의
  // ySyncPlugin/yCursorPlugin은 원격 갱신·커서 이동이 있을 때마다 "이 view에 적용할 setMeta"를
  // 모듈 전역 큐에 넣고 setTimeout(0)으로 나중에 한꺼번에 flush한다(y-prosemirror 소스의
  // updateMetas). 그 큐가 이미 이 view를 예약해 둔 상태에서 우리가 view.destroy()를 동기로
  // 불러버리면(destroy()는 view.docView를 null로 만든다), 나중에 실행되는 그 flush가 이미
  // 죽은 view에 대고 dispatch를 시도해 "Cannot read properties of null (reading 'matchesNode')"
  // 예외를 던진다 — 실제로 두 탭이 동시에 타이핑하는 중 슬라이드를 전환하는 시나리오로 재현
  // 확인했다(y-prosemirror 1.3.7 updateMetas가 binding.isDestroyed를 확인하긴 하지만, 그
  // 라이브러리 자신이 그 플래그를 어디서도 true로 설정하지 않아 가드가 항상 무력하다 — 업스트림
  // 결함이라 우리 쪽 코드에서 그 플래그에 기댈 수 없다). destroy()를 setTimeout(0)으로 미루면,
  // 클릭 이전에 이미 예약돼 있던 flush가 (아직 살아있는 view에 대해) 먼저 실행되어 정상적으로
  // 끝난 뒤에야 view가 파괴된다 — DOM에서는 아래에서 즉시 교체하므로 화면 전환 자체는 지연되지
  // 않는다(detached된 DOM 위에서 view.destroy()를 부르는 것 자체는 안전하다).
  const previousViews = [...mountedViews.values()];
  if (previousViews.length > 0) {
    setTimeout(() => {
      for (const view of previousViews) view.destroy();
    }, 0);
  }
  mountedViews.clear();
  canvasEl.style.width = `${emuToPx(SLIDE_W_EMU)}px`;
  canvasEl.style.height = `${emuToPx(SLIDE_H_EMU)}px`;
  canvasEl.replaceChildren();

  const perSlide = slideContents.get(entryId);
  if (!perSlide) {
    subscribeSlideContent(entryId); // 아직 안 열렸다 — 요청하고 observe가 다시 불러줄 때까지 대기
    return;
  }
  for (const [shapeId, shapeMap] of perSlide.entries()) {
    const meta = shapeMap.get('meta') as SlideShapeMeta;
    const box = document.createElement('div');
    box.className = 'slide-textbox';
    box.style.left = `${emuToPx(meta.x ?? SLIDE_W_EMU * 0.2)}px`;
    box.style.top = `${emuToPx(meta.y ?? SLIDE_H_EMU * 0.4)}px`;
    box.style.width = `${emuToPx(meta.w ?? SLIDE_W_EMU * NEW_TEXTBOX_WIDTH_RATIO)}px`;
    box.style.height = `${emuToPx(meta.h ?? SLIDE_H_EMU * NEW_TEXTBOX_HEIGHT_RATIO)}px`;
    canvasEl.append(box);

    const fragment = shapeMap.get('content') as Y.XmlFragment;
    const view = new EditorView(box, {
      state: EditorState.create({
        schema: pptxTextSchema,
        plugins: [ySyncPlugin(fragment), yCursorPlugin(provider.awareness), yUndoPlugin()],
      }),
      editable: () => canEdit,
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
      const ydoc = slidesArray.doc!;
      ydoc.transact(() => {
        slidesArray.insert(index + 1, [
          {
            id: crypto.randomUUID(),
            kind: 'duplicate',
            sourceSlidePart: entry.sourceSlidePart,
            sourcePageNumber: null,
          },
        ]);
      });
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
    if (activeSlideId && event.keysChanged.has(activeSlideId)) renderCanvasForSlide(activeSlideId);
  });
  // Task 8이 남겨 둔 훅 — 슬라이드를 고를 때마다 오른쪽 캔버스를 그 슬라이드 내용으로 채운다.
  onActiveSlideChange = renderCanvasForSlide;

  slidesArray.observe(renderSlides);
  renderSlides();
  setupDragAndDrop();
  setupCardActions();

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
