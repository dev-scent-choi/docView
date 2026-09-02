import { baseKeymap, chainCommands, setBlockType, toggleMark } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import { EditorState, Plugin, PluginKey, TextSelection, type Command } from 'prosemirror-state';
import { Decoration, DecorationSet, EditorView } from 'prosemirror-view';
import { liftListItem, sinkListItem, splitListItem, wrapInList } from 'prosemirror-schema-list';
import {
  addColumnAfter,
  addRowAfter,
  columnResizing,
  deleteColumn,
  deleteRow,
  deleteTable,
  goToNextCell,
  tableEditing,
} from 'prosemirror-tables';
import type { Node as PMNode } from 'prosemirror-model';
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { redo, undo, ySyncPlugin, yCursorPlugin, yUndoPlugin } from 'y-prosemirror';
import { docxSchema } from '../docx/schema.js';
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
  readonly editHasUnsupportedContent: boolean;
}

const params = new URLSearchParams(location.search);
const token = params.get('token') ?? '';

const titleEl = document.getElementById('doc-title') as HTMLElement;
const statusEl = document.getElementById('conn-status') as HTMLElement;
const presenceEl = document.getElementById('presence-list') as HTMLElement;
const warningEl = document.getElementById('unsupported-warning') as HTMLElement;
const editorHost = document.getElementById('editor') as HTMLElement;
const errorScreen = document.getElementById('error-screen') as HTMLElement;
const errorMessage = document.getElementById('error-message') as HTMLElement;
const toolbar = document.getElementById('toolbar') as HTMLElement;

function showFatalError(message: string): void {
  editorHost.hidden = true;
  toolbar.hidden = true;
  errorMessage.textContent = message;
  errorScreen.hidden = false;
}

async function fetchViewInfo(): Promise<ViewInfo> {
  const res = await fetch(`/api/view/info?token=${encodeURIComponent(token)}`);
  const json = (await res.json()) as { success: boolean; data?: ViewInfo; error?: string };
  if (!res.ok || !json.success || !json.data) {
    throw new Error(json.error || '문서 정보를 불러오지 못했습니다.');
  }
  return json.data;
}

function buildKeymap(): Record<string, Command> {
  const nodes = docxSchema.nodes;
  const marks = docxSchema.marks;
  return {
    ...baseKeymap,
    'Mod-z': undo,
    'Mod-y': redo,
    'Mod-Shift-z': redo,
    'Mod-b': toggleMark(marks.strong),
    'Mod-i': toggleMark(marks.em),
    'Mod-u': toggleMark(marks.underline),
    'Mod-Alt-1': setBlockType(nodes.heading, { level: 1 }),
    'Mod-Alt-2': setBlockType(nodes.heading, { level: 2 }),
    'Mod-Alt-3': setBlockType(nodes.heading, { level: 3 }),
    'Mod-Alt-0': setBlockType(nodes.paragraph),
    Tab: chainCommands(sinkListItem(nodes.list_item), goToNextCell(1)),
    'Shift-Tab': chainCommands(liftListItem(nodes.list_item), goToNextCell(-1)),
    Enter: chainCommands(splitListItem(nodes.list_item), baseKeymap.Enter),
  };
}

/** toggleMark과 달리 항상 "덮어쓰기"만 한다 — 색/크기처럼 값이 여러 개인 마크는 토글하면
 * 다른 값으로 바꾸려 할 때 그냥 지워져 버리므로 별도 커맨드가 필요하다. */
function setMark(markType: (typeof docxSchema.marks)[keyof typeof docxSchema.marks], attrs: Record<string, unknown>): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) return false;
    if (dispatch) dispatch(state.tr.addMark(from, to, markType.create(attrs)));
    return true;
  };
}

function removeMark(markType: (typeof docxSchema.marks)[keyof typeof docxSchema.marks]): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) return false;
    if (dispatch) dispatch(state.tr.removeMark(from, to, markType));
    return true;
  };
}

const ALIGNABLE_NODE_NAMES = new Set(['paragraph', 'heading']);

function setAlign(align: 'left' | 'center' | 'right'): Command {
  return (state, dispatch) => {
    const { from, to } = state.selection;
    let applicable = false;
    state.doc.nodesBetween(from, to, (node) => {
      if (ALIGNABLE_NODE_NAMES.has(node.type.name)) applicable = true;
    });
    if (!applicable) return false;
    if (dispatch) {
      const tr = state.tr;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (ALIGNABLE_NODE_NAMES.has(node.type.name)) {
          tr.setNodeMarkup(pos, undefined, { ...node.attrs, align: align === 'left' ? null : align });
        }
      });
      // "전체 선택"(Ctrl+A)의 to는 doc.content.size인데, 그 위치는 문단 "밖"(depth 0=doc)에서
      // 끝난다 — TextSelection.create(tr.doc, from, to)는 그 경계에서 그대로 예외를 던진다.
      // TextSelection.between은 같은 상황에서 예외 대신 가장 가까운 유효한 텍스트 위치로
      // 스냅해 주므로, "전체 선택 후 서식 적용" 흐름에서 안전하게 선택을 복원할 수 있다.
      tr.setSelection(TextSelection.between(tr.doc.resolve(from), tr.doc.resolve(Math.min(to, tr.doc.content.size))));
      dispatch(tr);
    }
    return true;
  };
}

interface SearchMatch {
  readonly from: number;
  readonly to: number;
}

interface SearchPluginState {
  readonly matches: readonly SearchMatch[];
  readonly current: number;
}

const searchPluginKey = new PluginKey<SearchPluginState>('search');

function findMatches(doc: PMNode, query: string): SearchMatch[] {
  if (!query) return [];
  const lower = query.toLowerCase();
  const matches: SearchMatch[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    const text = node.text.toLowerCase();
    let idx = text.indexOf(lower);
    while (idx !== -1) {
      matches.push({ from: pos + idx, to: pos + idx + query.length });
      idx = text.indexOf(lower, idx + 1);
    }
  });
  return matches;
}

function buildSearchPlugin(): Plugin<SearchPluginState> {
  return new Plugin<SearchPluginState>({
    key: searchPluginKey,
    state: {
      init: () => ({ matches: [], current: -1 }),
      apply(tr, prev) {
        const meta = tr.getMeta(searchPluginKey) as SearchPluginState | undefined;
        if (meta) return meta;
        if (tr.docChanged) {
          const matches = prev.matches.map((m) => ({ from: tr.mapping.map(m.from), to: tr.mapping.map(m.to) }));
          return { matches, current: prev.current };
        }
        return prev;
      },
    },
    props: {
      decorations(state) {
        const { matches, current } = searchPluginKey.getState(state) ?? { matches: [], current: -1 };
        if (matches.length === 0) return DecorationSet.empty;
        const decorations = matches.map((match, index) =>
          Decoration.inline(match.from, match.to, {
            class: index === current ? 'editor-search-hit editor-search-hit-current' : 'editor-search-hit',
          }),
        );
        return DecorationSet.create(state.doc, decorations);
      },
    },
  });
}

/** 툴바 버튼에 현재 선택 영역의 서식을 반영한다(오피스처럼 굵게/정렬 등이 눌린 상태로 보이게). */
function markActive(state: EditorState, markType: (typeof docxSchema.marks)[keyof typeof docxSchema.marks]): boolean {
  const { from, to, empty, $from } = state.selection;
  return empty ? !!markType.isInSet(state.storedMarks ?? $from.marks()) : state.doc.rangeHasMark(from, to, markType);
}

function currentAlign(state: EditorState): string | null {
  const attrs = state.selection.$from.parent.attrs as { align?: string | null };
  return attrs.align ?? null;
}

function isBlockActive(
  state: EditorState,
  nodeType: (typeof docxSchema.nodes)[keyof typeof docxSchema.nodes],
  attrs: Record<string, unknown> = {},
): boolean {
  const parent = state.selection.$from.parent;
  if (parent.type !== nodeType) return false;
  return Object.entries(attrs).every(([key, value]) => parent.attrs[key] === value);
}

function setPressed(id: string, pressed: boolean): void {
  document.getElementById(id)?.setAttribute('aria-pressed', String(pressed));
}

function syncToolbarState(state: EditorState): void {
  setPressed('btn-bold', markActive(state, docxSchema.marks.strong));
  setPressed('btn-italic', markActive(state, docxSchema.marks.em));
  setPressed('btn-underline', markActive(state, docxSchema.marks.underline));
  setPressed('btn-strike', markActive(state, docxSchema.marks.strike));
  const align = currentAlign(state);
  setPressed('btn-align-left', !align || align === 'left');
  setPressed('btn-align-center', align === 'center');
  setPressed('btn-align-right', align === 'right');
  setPressed('btn-paragraph', isBlockActive(state, docxSchema.nodes.paragraph));
  setPressed('btn-h1', isBlockActive(state, docxSchema.nodes.heading, { level: 1 }));
  setPressed('btn-h2', isBlockActive(state, docxSchema.nodes.heading, { level: 2 }));
  setPressed('btn-h3', isBlockActive(state, docxSchema.nodes.heading, { level: 3 }));
}

function insertTable(view: EditorView): void {
  const { table, table_row: row, table_cell: cell, paragraph } = docxSchema.nodes;
  const makeCell = () => cell.create(null, paragraph.create());
  const makeRow = () => row.create(null, [makeCell(), makeCell()]);
  const node = table.create(null, [makeRow(), makeRow()]);
  const tr = view.state.tr.replaceSelectionWith(node);
  view.dispatch(tr);
  view.focus();
}

function bindToolbarButton(
  view: EditorView,
  id: string,
  command: Command,
): void {
  const button = document.getElementById(id);
  if (!button) return;
  button.addEventListener('click', () => {
    command(view.state, view.dispatch, view);
    view.focus();
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
  if (info.editHasUnsupportedContent) {
    warningEl.hidden = false;
  }

  const ydoc = new Y.Doc();
  const fragment = ydoc.getXmlFragment('prosemirror');
  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, {
    params: { token },
  });

  // 이름 기준으로 색을 정한다(ydoc.clientID는 접속마다 무작위라 같은 사람이 다른 브라우저로
  // 들어오면 매번 색이 달라져 버린다) — 같은 이름이면 항상 같은 색.
  const localColor = pickPresenceColorForName(info.viewer);
  provider.awareness.setLocalStateField('user', { name: info.viewer, color: localColor });

  provider.on('status', (event: { status: string }) => {
    statusEl.textContent =
      event.status === 'connected' ? '연결됨' : event.status === 'disconnected' ? '연결 끊김 — 재연결 중…' : event.status;
    statusEl.dataset.status = event.status;
  });
  provider.on('sync', (isSynced: boolean) => {
    if (isSynced) statusEl.dataset.status = 'synced';
  });
  provider.awareness.on('change', () => renderPresenceAvatars(presenceEl, provider.awareness));
  renderPresenceAvatars(presenceEl, provider.awareness);

  const state = EditorState.create({
    schema: docxSchema,
    plugins: [
      ySyncPlugin(fragment),
      yCursorPlugin(provider.awareness),
      yUndoPlugin(),
      keymap(buildKeymap()),
      columnResizing(),
      tableEditing(),
      buildSearchPlugin(),
    ],
  });

  // dispatchTransaction 콜백은 view를 참조하므로, new EditorView(...) 생성자 실행 도중(예:
  // ySyncPlugin의 초기 동기화가 생성자 안에서 동기적으로 트랜잭션을 dispatch하는 경우) 콜백이
  // 곧바로 불리면 아직 대입되지 않은 view를 참조해 예외가 난다 — 그러면 main()의 catch가
  // showFatalError를 호출해 이미 마운트된 .ProseMirror까지 통째로 숨겨버린다. 그래서 생성자는
  // 기본 동작으로 완료시키고, view가 확실히 대입된 뒤 setProps로 커스텀 콜백을 나중에 건다.
  const view = new EditorView(editorHost, { state });
  view.setProps({
    dispatchTransaction(tr) {
      const newState = view.state.apply(tr);
      view.updateState(newState);
      syncToolbarState(newState);
    },
  });
  syncToolbarState(view.state);

  bindToolbarButton(view, 'btn-bold', toggleMark(docxSchema.marks.strong));
  bindToolbarButton(view, 'btn-italic', toggleMark(docxSchema.marks.em));
  bindToolbarButton(view, 'btn-underline', toggleMark(docxSchema.marks.underline));
  bindToolbarButton(view, 'btn-strike', toggleMark(docxSchema.marks.strike));
  bindToolbarButton(view, 'btn-h1', setBlockType(docxSchema.nodes.heading, { level: 1 }));
  bindToolbarButton(view, 'btn-h2', setBlockType(docxSchema.nodes.heading, { level: 2 }));
  bindToolbarButton(view, 'btn-h3', setBlockType(docxSchema.nodes.heading, { level: 3 }));
  bindToolbarButton(view, 'btn-paragraph', setBlockType(docxSchema.nodes.paragraph));
  bindToolbarButton(view, 'btn-bullet-list', wrapInList(docxSchema.nodes.bullet_list));
  bindToolbarButton(view, 'btn-ordered-list', wrapInList(docxSchema.nodes.ordered_list));
  bindToolbarButton(view, 'btn-add-row', addRowAfter);
  bindToolbarButton(view, 'btn-add-col', addColumnAfter);
  bindToolbarButton(view, 'btn-del-row', deleteRow);
  bindToolbarButton(view, 'btn-del-col', deleteColumn);
  bindToolbarButton(view, 'btn-del-table', deleteTable);
  document.getElementById('btn-table')?.addEventListener('click', () => insertTable(view));
  document.getElementById('btn-undo')?.addEventListener('click', () => {
    undo(view.state);
    view.focus();
  });
  document.getElementById('btn-redo')?.addEventListener('click', () => {
    redo(view.state);
    view.focus();
  });

  const fontColorInput = document.getElementById('font-color') as HTMLInputElement | null;
  const fontColorBtn = document.getElementById('font-color-btn') as HTMLElement | null;
  fontColorBtn?.style.setProperty('--btn-color', fontColorInput?.value ?? '#1b1a17');
  fontColorInput?.addEventListener('input', () => {
    fontColorBtn?.style.setProperty('--btn-color', fontColorInput.value);
    setMark(docxSchema.marks.textColor, { color: fontColorInput.value })(view.state, view.dispatch);
    view.focus();
  });
  bindToolbarButton(view, 'btn-clear-color', removeMark(docxSchema.marks.textColor));

  const fontSizeSelect = document.getElementById('font-size') as HTMLSelectElement | null;
  fontSizeSelect?.addEventListener('change', () => {
    const size = Number(fontSizeSelect.value);
    if (size > 0) setMark(docxSchema.marks.fontSize, { size })(view.state, view.dispatch);
    fontSizeSelect.value = '';
    view.focus();
  });

  bindToolbarButton(view, 'btn-align-left', setAlign('left'));
  bindToolbarButton(view, 'btn-align-center', setAlign('center'));
  bindToolbarButton(view, 'btn-align-right', setAlign('right'));

  // ---------- 확대/축소 ----------
  const zoomHost = document.getElementById('editor-zoom-host') as HTMLElement;
  const zoomLevelEl = document.getElementById('zoom-level') as HTMLElement;
  const ZOOM_MIN = 0.6;
  const ZOOM_MAX = 2;
  let zoom = 1;
  function applyZoom(): void {
    zoomHost.style.zoom = String(zoom);
    zoomLevelEl.textContent = `${Math.round(zoom * 100)}%`;
  }
  document.getElementById('zoom-in')?.addEventListener('click', () => {
    zoom = Math.min(ZOOM_MAX, Math.round((zoom + 0.1) * 10) / 10);
    applyZoom();
  });
  document.getElementById('zoom-out')?.addEventListener('click', () => {
    zoom = Math.max(ZOOM_MIN, Math.round((zoom - 0.1) * 10) / 10);
    applyZoom();
  });

  // ---------- 문서 내 검색 ----------
  const searchBar = document.getElementById('search-bar') as HTMLElement;
  const searchInput = document.getElementById('search-input') as HTMLInputElement;
  const searchCount = document.getElementById('search-count') as HTMLElement;
  const searchPrevBtn = document.getElementById('search-prev') as HTMLButtonElement;
  const searchNextBtn = document.getElementById('search-next') as HTMLButtonElement;
  let searchMatches: SearchMatch[] = [];
  let searchCurrent = -1;

  function applySearch(jumpTo: number): void {
    searchMatches = findMatches(view.state.doc, searchInput.value);
    searchCurrent = searchMatches.length > 0 ? ((jumpTo % searchMatches.length) + searchMatches.length) % searchMatches.length : -1;
    let tr = view.state.tr.setMeta(searchPluginKey, { matches: searchMatches, current: searchCurrent });
    if (searchCurrent >= 0) {
      const match = searchMatches[searchCurrent];
      tr = tr.setSelection(TextSelection.near(tr.doc.resolve(match.from))).scrollIntoView();
    }
    view.dispatch(tr);
    searchCount.textContent = searchMatches.length > 0 ? `${searchCurrent + 1}/${searchMatches.length}` : searchInput.value ? '0/0' : '';
    searchPrevBtn.disabled = searchMatches.length === 0;
    searchNextBtn.disabled = searchMatches.length === 0;
  }
  function openSearch(): void {
    searchBar.hidden = false;
    searchInput.focus();
    searchInput.select();
  }
  function closeSearch(): void {
    searchBar.hidden = true;
    searchInput.value = '';
    applySearch(0);
  }
  document.getElementById('search-btn')?.addEventListener('click', () => {
    if (searchBar.hidden) openSearch();
    else closeSearch();
  });
  document.getElementById('search-close')?.addEventListener('click', closeSearch);
  searchInput.addEventListener('input', () => applySearch(0));
  searchInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') applySearch(searchCurrent + (event.shiftKey ? -1 : 1));
    if (event.key === 'Escape') closeSearch();
  });
  searchPrevBtn.addEventListener('click', () => applySearch(searchCurrent - 1));
  searchNextBtn.addEventListener('click', () => applySearch(searchCurrent + 1));
  window.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      openSearch();
    }
  });

  window.addEventListener('beforeunload', () => {
    provider.disconnect();
  });
}

main().catch((error) => {
  console.error(error);
  showFatalError('편집기를 시작하지 못했습니다.');
});
