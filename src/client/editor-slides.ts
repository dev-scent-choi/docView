import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from '../collab/pptxShared.js';
import { DOC_META_MAP, DOC_META_SAVED_AT_KEY, DOC_META_SAVE_ERROR_AT_KEY } from '../collab/gridShared.js';
import { pickPresenceColorForName, renderPresenceAvatars } from './presence.js';

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

// 어떤 슬라이드가 오른쪽 캔버스에 표시 중인지 — Task 9가 이 값을 읽어 실제 콘텐츠를 그린다.
let activeSlideId: string | null = null;
// Task 9가 여기에 실제 렌더링 로직을 걸어 넣는 훅 포인트. 이 태스크에서는 상태 갱신만 하고
// 아직 아무것도 구독하지 않는다.
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
// 실제 콘텐츠 렌더링은 Task 9가 onActiveSlideChange 콜백에서 처리한다 — 여기서는 상태/표시만.
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
  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, {
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
