import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { PPTX_SLIDES_ARRAY, type SlideEntry } from '../collab/pptxShared.js';
import { DOC_META_MAP, DOC_META_SAVED_AT_KEY } from '../collab/gridShared.js';
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
const gridEl = document.getElementById('slides-grid') as HTMLElement;
const cardTemplate = document.getElementById('slide-card-template') as HTMLTemplateElement;
const presenceListEl = document.getElementById('presence-list') as HTMLElement;

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
  gridEl.replaceChildren(...entries.map((entry, index) => cardFor(entry, index)));
}

function setupDragAndDrop(): void {
  gridEl.addEventListener('dragstart', (event) => {
    const card = (event.target as HTMLElement).closest<HTMLElement>('.slide-card');
    if (!card) return;
    dragFromIndex = Number(card.dataset.index);
    event.dataTransfer?.setData('text/plain', card.dataset.id ?? '');
  });
  gridEl.addEventListener('dragover', (event) => {
    if (dragFromIndex !== null) event.preventDefault();
  });
  gridEl.addEventListener('drop', (event) => {
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
  gridEl.addEventListener('click', (event) => {
    if (!canEdit) return;
    const target = event.target as HTMLElement;
    const card = target.closest<HTMLElement>('.slide-card');
    if (!card) return;
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
  if (!info.permissions.canEdit) {
    showFatalError('이 링크에는 동시편집 권한이 없습니다. 뷰어에서 열람만 가능합니다.');
    return;
  }

  titleEl.textContent = info.name;
  document.title = `${info.name} — DocView 편집`;
  canEdit = info.permissions.canEdit;

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
  function setSaveStatus(status: 'saving' | 'saved'): void {
    if (!saveStatusEl) return;
    saveStatusEl.hidden = false;
    saveStatusEl.dataset.status = status;
    if (saveStatusTextEl) saveStatusTextEl.textContent = status === 'saving' ? '저장 중…' : '모든 변경 내용이 저장됨';
  }
  if (canEdit) {
    ydoc.on('afterTransaction', (tx) => {
      if (!tx.local || tx.changed.size === 0) return;
      setSaveStatus('saving');
    });
  }
  docMeta.observe(() => {
    const savedAt = docMeta.get(DOC_META_SAVED_AT_KEY) ?? 0;
    if (savedAt > lastKnownSavedAt) {
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
