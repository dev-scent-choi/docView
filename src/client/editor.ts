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
  TableView,
  tableEditing,
} from 'prosemirror-tables';
import type { Node as PMNode } from 'prosemirror-model';
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { redo, undo, ySyncPlugin, yCursorPlugin, yUndoPlugin } from 'y-prosemirror';
import { docxSchema, type PageSetup } from '../docx/schema.js';
import { pickPresenceColorForName, renderPresenceAvatars } from './presence.js';

interface Permissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
  readonly canEdit: boolean;
}

interface ScriptFontFamily {
  readonly eastAsia: string | null;
  readonly ascii: string | null;
}

interface ViewInfo {
  readonly docId: string;
  readonly name: string;
  readonly viewer: string;
  readonly permissions: Permissions;
  readonly editHasUnsupportedContent: boolean;
  readonly pageSetup: PageSetup | null;
  readonly defaultFontSizePt: number | null;
  readonly defaultFontFamily: ScriptFontFamily | null;
}

const params = new URLSearchParams(location.search);
const token = params.get('token') ?? '';

/** false면 이 문서는 읽기 전용이다(동시편집 권한 없는 링크) — grid-editor.ts의 같은 이름 변수와
 * 같은 역할. main()에서 서버가 내려준 토큰 권한으로 설정된다. ProseMirror의 EditorView가
 * `editable: () => canEdit`로 이 값을 직접 참조해 키보드/붙여넣기/드롭 등 DOM에서 오는 편집
 * 이벤트 전체를 스스로 거부하고, 여기서는 그 위에 툴바(서식 리본) 자체를 숨겨 애초에 사용자가
 * 고칠 수 있다고 착각하지 않게 한다. 실제 쓰기 차단은 여전히 서버(collab/room.ts의
 * readOnlyConns)가 독립적으로 수행한다.
 *
 * 주의: `editable:false`는 ProseMirror 코어의 `editHandlers` 목록(keydown/paste/cut/drop/
 * composition류)에 있는 DOM 이벤트만 막는다 — view.dispatch()를 직접 부르는 경로(툴바 버튼 등)는
 * 이 값을 아예 안 본다(그래서 리본을 숨기는 게 장식이 아니라 실질적 방어다). 표 열 너비 드래그처럼
 * editHandlers 밖의 handleDOMEvents 경로는 현재 prosemirror-tables 자신이 내부에서 view.editable을
 * 확인해 막아주고 있을 뿐, ProseMirror 코어가 보장하는 동작이 아니다 — 이 라이브러리를 업그레이드
 * 하거나 handleDOMEvents를 쓰는 새 플러그인을 추가할 때는 그 플러그인이 editable을 스스로
 * 확인하는지 다시 확인할 것. 어느 경우든 서버 쪽 readOnlyConns 차단이 최종 방어선이라 이 가정이
 * 깨져도 실제 쓰기가 반영되지는 않는다. */
let canEdit = true;

const titleEl = document.getElementById('doc-title') as HTMLElement;
const statusEl = document.getElementById('conn-status') as HTMLElement;
const readonlyBadgeEl = document.getElementById('readonly-badge') as HTMLElement;
const presenceEl = document.getElementById('presence-list') as HTMLElement;
const warningEl = document.getElementById('unsupported-warning') as HTMLElement;
const editorHost = document.getElementById('editor') as HTMLElement;
const errorScreen = document.getElementById('error-screen') as HTMLElement;
const errorMessage = document.getElementById('error-message') as HTMLElement;
const toolbar = document.getElementById('toolbar') as HTMLElement;

/**
 * Word 원본의 실제 페이지 크기·여백으로 편집 화면의 "종이" 크기를 그린다 — 기존엔 항상
 * 고정된 CSS 값(860px, 65ch 등 문서 크기와 무관한 임의 수치)을 썼는데, 그 임의값이 곧 "용지
 * 폭이 실제보다 좁아 보인다"/"셀 너비가 원본과 다르다"는 문제의 원인이었다. 페이지 전체 폭을
 * 그대로 쓰고 여백을 padding으로 재현하면, 안쪽 콘텐츠 폭(=페이지 폭-여백)이 Word의 실제
 * 텍스트 영역 폭과 정확히 같아진다 — 표가 그 안에서 백분율이 아니라 자기 실제 폭(px)으로
 * 그려지므로(schema.ts의 table toDOM) 셀 너비도 함께 정확해진다.
 */
function applyPageSetup(pageSetup: PageSetup | null): void {
  if (!pageSetup) return; // 원본에서 못 읽으면(비정상 문서) 기존 CSS 기본값을 그대로 둔다
  editorHost.style.width = `${pageSetup.widthPx}px`;
  editorHost.style.maxWidth = `${pageSetup.widthPx}px`;
  editorHost.style.padding =
    `${pageSetup.marginTopPx}px ${pageSetup.marginRightPx}px ${pageSetup.marginBottomPx}px ${pageSetup.marginLeftPx}px`;
  // .ProseMirror의 임의 max-width:65ch 캡을 풀어서 실제 여백만큼의 콘텐츠 폭을 그대로 쓰게 한다
  editorHost.classList.add('page-sized');
}

/**
 * 직접 서식도 문단 스타일도 없는 실행문은 Word에서 word/styles.xml의 문서 기본값(대개
 * 10~11pt)을 상속받아 보여지는데, 웹뷰어는 그 상속을 계산하지 않고 브라우저 기본값(16px)을
 * 그대로 썼다 — 그래서 직접 서식이 거의 없는 문서 전체가 원본보다 커 보였다. editorHost에
 * 상속 가능한 기본 font-size로 걸어두면, 직접 서식이 있는 실행문(fontSize 마크가 만든
 * 인라인 style)은 이 상속값을 그대로 덮어써 영향받지 않는다.
 */
function applyDefaultFontSize(pt: number | null): void {
  if (!pt) return;
  editorHost.style.fontSize = `${(pt * 4) / 3}px`; // 1pt = 96/72px = 4/3px
}

function escapeXmlText(value: string): string {
  return value.replace(/[<>&'"]/g, (ch) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' })[ch] as string);
}

/** "2026-09-06 23:30:11" — toISOString은 UTC라 화면을 보는 사람의 실제 시계와 어긋나 보이므로
 * 로컬 시각 그대로 초 단위까지 조립한다. viewer.js의 동명 함수와 같은 포맷(별도 번들이라 공유 불가). */
function formatWatermarkTimestamp(date: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ` +
    `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`
  );
}

/** 뷰어(viewer.js)와 같은 화면 열람용 워터마크를 동시편집 화면에도 띄운다 — 토큰에 서명된(위조
 * 불가) viewer 이름 + 이 화면을 연 시각을 편집 화면 전체에 옅게 반복해 깐다. */
function applyWatermark(viewerName: string): void {
  if (!viewerName) return;
  const text = `${viewerName}  ${formatWatermarkTimestamp(new Date())}`;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="300" height="180">` +
    `<text x="150" y="100" transform="rotate(-28 150 100)" text-anchor="middle" ` +
    `font-family="sans-serif" font-size="14" fill="rgba(0,0,0,0.16)">${escapeXmlText(text)}</text>` +
    `</svg>`;
  document.documentElement.style.setProperty('--watermark-bg', `url("data:image/svg+xml,${encodeURIComponent(svg)}")`);
}

/**
 * 한글(Hangul) 코드 영역 — 이 범위의 글자만 eastAsia 서체(예: 굴림)를 쓰게 강제하고, 나머지
 * (숫자·영문·기호 등)는 ascii 서체(예: Arial)로 넘긴다. 그냥 "굴림, Arial" 순서로만 나열하면
 * 브라우저는 "이 글자를 그릴 수 있는 첫 폰트"로 폴백하는데, 굴림도 숫자 글리프를 갖고 있어
 * Arial까지 안 내려가 버린다 — 그래서 날짜("2026. 07. 24") 같은 숫자가 Word에서 보이는 것보다
 * 눈에 띄게 얇게 보였다(실제 문서로 확인한 결과). unicode-range로 스크립트별 폰트를 강제로
 * 나눠야 Word처럼 숫자는 ascii 서체로, 한글은 eastAsia 서체로 각각 그려진다.
 */
const HANGUL_UNICODE_RANGE = 'U+1100-11FF, U+3130-318F, U+A960-A97F, U+AC00-D7A3, U+D7B0-D7FF';

/**
 * 문서의 실제 기본 서체(한글 eastAsia + 영문/숫자 ascii, word/theme/theme1.xml로 해석된 값)를
 * 편집 화면 기본 폰트로 적용한다. 두 서체가 다르면 위 unicode-range로 스크립트별로 갈라 쓰는
 * 합성 폰트를 만들어 쓰고, 하나만 있거나 둘이 같으면 그냥 그 서체 하나만 쓴다.
 */
function applyDefaultFontFamily(family: ScriptFontFamily | null): void {
  if (!family || (!family.eastAsia && !family.ascii)) return; // 이게 없으면 이 앱 UI용 폰트로 그려져 Word와 아예 다른 서체가 된다
  const { eastAsia, ascii } = family;
  // .ProseMirror 자신에게 이미 font-family 선언이 있어(style.css) 상위 요소의 inline
  // font-family는 상속되지 않는다 — 그 선언이 참조하는 CSS 변수로 값을 흘려보낸다.
  if (eastAsia && ascii && eastAsia !== ascii) {
    const styleEl = document.createElement('style');
    styleEl.textContent =
      `@font-face { font-family: "doc-mixed-font"; src: local("${eastAsia}"); unicode-range: ${HANGUL_UNICODE_RANGE}; }` +
      `@font-face { font-family: "doc-mixed-font"; src: local("${ascii}"); }`;
    document.head.appendChild(styleEl);
    editorHost.style.setProperty(
      '--doc-font-family',
      `"doc-mixed-font", "${eastAsia}", "${ascii}", var(--font-sans)`,
    );
  } else {
    editorHost.style.setProperty('--doc-font-family', `"${eastAsia ?? ascii}", var(--font-sans)`);
  }
}

/**
 * columnResizing()의 기본 TableView는 <table>을 직접 document.createElement로 만들고
 * schema.ts의 table.toDOM은 아예 호출하지 않는다 — 그래서 표 자체의 정렬(w:tblPr/w:jc,
 * 셀 안 문단 정렬과는 별개)이 스키마에만 있고 화면엔 전혀 반영되지 않았다. columnResizing이
 * 공식으로 제공하는 View 옵션으로 이 클래스를 대신 넣어 정렬 스타일만 덧붙인다.
 */
class AlignedTableView extends TableView {
  constructor(node: PMNode, defaultCellMinWidth: number) {
    super(node, defaultCellMinWidth);
    this.applyAlign(node);
  }

  private applyAlign(node: PMNode): void {
    const align = node.attrs.align as string | null;
    if (align === 'center') {
      this.table.style.marginLeft = 'auto';
      this.table.style.marginRight = 'auto';
    } else if (align === 'right') {
      this.table.style.marginLeft = 'auto';
      this.table.style.marginRight = '0';
    } else {
      this.table.style.marginLeft = '';
      this.table.style.marginRight = '';
    }
  }

  override update(node: PMNode): boolean {
    const ok = super.update(node);
    if (ok) this.applyAlign(node);
    return ok;
  }
}

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

  // canEdit:false여도 이 편집기 화면 자체는 그대로 보여준다 — grid-editor.ts(xlsx)/editor-slides.ts
  // (pptx)와 같은 정책: 실제 쓰기 차단은 서버가 읽기 전용 연결로 처리하므로(collab/room.ts), 여기선
  // 서식 도구(리본) 그룹만 숨겨 애초에 사용자가 고칠 수 있다고 착각하지 않게 한다 — 뒤로가기/제목/
  // 검색/참여자 같은 상단 줄은 도구가 아니라 계속 보여야 하므로 toolbar 전체가 아니라
  // .toolbar-format만 숨긴다.
  canEdit = info.permissions.canEdit;
  if (!canEdit) {
    toolbar.querySelector('.toolbar-format')?.setAttribute('hidden', '');
    readonlyBadgeEl.hidden = false;
  }

  titleEl.textContent = info.name;
  document.title = `${info.name} — DocView 편집`;
  if (info.editHasUnsupportedContent) {
    warningEl.hidden = false;
  }
  applyPageSetup(info.pageSetup);
  applyDefaultFontSize(info.defaultFontSizePt);
  applyDefaultFontFamily(info.defaultFontFamily);
  applyWatermark(info.viewer);

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
  // 서버가 편집 세션 자체를 거부한 경우(edit.ts의 socket.close(4xxx, 이유)) — 4400대는 재시도해도
  // 같은 이유로 다시 거부될 뿐인 영구 거부다(잘못된 토큰, 문서 없음 등). 기본 동작(지수 백오프로
  // 계속 재연결 시도)을 멈추고 서버가 준 이유를 그대로 보여준다 — grid-editor.ts와 동일한 처리.
  provider.on('connection-close', (event: CloseEvent | null) => {
    if (event && event.code >= 4400) {
      provider.shouldConnect = false;
      showFatalError(event.reason || '서버가 연결을 거부했습니다.');
    }
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
      columnResizing({ View: AlignedTableView }),
      tableEditing(),
      buildSearchPlugin(),
    ],
  });

  // dispatchTransaction 콜백은 view를 참조하므로, new EditorView(...) 생성자 실행 도중(예:
  // ySyncPlugin의 초기 동기화가 생성자 안에서 동기적으로 트랜잭션을 dispatch하는 경우) 콜백이
  // 곧바로 불리면 아직 대입되지 않은 view를 참조해 예외가 난다 — 그러면 main()의 catch가
  // showFatalError를 호출해 이미 마운트된 .ProseMirror까지 통째로 숨겨버린다. 그래서 생성자는
  // 기본 동작으로 완료시키고, view가 확실히 대입된 뒤 setProps로 커스텀 콜백을 나중에 건다.
  // editable: 프로세미러 표준 EditorView 옵션 — false면 DOM contentEditable을 끄고, 키보드
  // 입력·IME 조합·붙여넣기·드롭·자르기 등 "DOM에서 오는" 편집 이벤트 전체를 프로세미러 자신이
  // 무시한다(prosemirror-view의 editHandlers가 view.editable을 직접 검사). 다만 이 게이트는
  // 어디까지나 DOM 이벤트 경로에만 걸리고 view.dispatch를 직접 호출하는 코드(예: 아래 툴바 버튼
  // 클릭 핸들러)는 우회한다 — 그래서 이 옵션만으로는 부족하고, 툴바(.toolbar-format) 자체를
  // 위에서 이미 숨겨 그 우회로를 원천 차단했다.
  const view = new EditorView(editorHost, { state, editable: () => canEdit });
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
  // 뷰어("바로 보기")는 페이지를 실제 96dpi 크기가 아니라 .pages-container(최대 1120px,
  // 좌우 1rem 패딩 → 실사용 1088px)에 맞춰 확대해서 보여준다 — 편집기는 지금까지 페이지를
  // 항상 실제 크기(zoom=100%)로만 그려서, 같은 문서인데 "종이 자체가 더 작아 보인다"는
  // 차이가 났다. 뷰어와 같은 컨테이너 폭 기준으로 초기 줌을 계산해 화면에 보이는 크기를 맞춘다.
  const VIEWER_PAGE_CONTAINER_MAX_PX = 1120;
  const VIEWER_PAGE_CONTAINER_PADDING_PX = 32; // 뷰어 .pages-container의 좌우 1rem(16px)×2
  function computeInitialZoom(pageWidthPx: number): number {
    const availableWidthPx = Math.min(VIEWER_PAGE_CONTAINER_MAX_PX, window.innerWidth) - VIEWER_PAGE_CONTAINER_PADDING_PX;
    const fitZoom = availableWidthPx / pageWidthPx;
    return Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(fitZoom * 10) / 10));
  }
  let zoom = info.pageSetup ? computeInitialZoom(info.pageSetup.widthPx) : 1;
  function applyZoom(): void {
    zoomHost.style.zoom = String(zoom);
    zoomLevelEl.textContent = `${Math.round(zoom * 100)}%`;
  }
  applyZoom();
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
  // applySearch가 매번 findMatches로 문서 전체를 다시 훑으므로(문서가 크면 체감되는 비용), 입력
  // 중에는 짧게 모아서(디바운스) 실행한다 — Enter/이전·다음 버튼 등 명시적 액션은 즉시 반응해야
  // 하므로 이 디바운스를 적용하지 않는다.
  let searchInputDebounce: ReturnType<typeof setTimeout> | undefined;
  searchInput.addEventListener('input', () => {
    clearTimeout(searchInputDebounce);
    searchInputDebounce = setTimeout(() => applySearch(0), 150);
  });
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
