import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import * as encoding from 'lib0/encoding';
import {
  DOC_META_MAP,
  DOC_META_SAVED_AT_KEY,
  DOC_META_SAVE_ERROR_AT_KEY,
  EDIT_SOFT_COL_LIMIT,
  EDIT_SOFT_ROW_LIMIT,
  GRID_CELLS_MAP,
  GRID_COMPUTED_MAP,
  GRID_MERGES_MAP,
  GRID_STRUCTURE_MAP,
  GRID_STYLES_MAP,
  gridCellKey,
  parseGridCellKey,
  PAGE_SIZE_ROWS,
  type CellStyleAttrs,
  type GridMergeShape,
  type SheetStructure,
} from '../collab/gridShared.js';
import { dedupeByName, pickPresenceColorForName, renderPresenceAvatars, type AwarenessUserState } from './presence.js';

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

const ALIGN_TO_CSS: Record<'l' | 'c' | 'r', string> = { l: 'left', c: 'center', r: 'right' };
/** 글꼴 드롭다운에서 고른 이름이 이 컴퓨터/브라우저에 없을 때를 위한 대체 글꼴 목록. */
const DEFAULT_CELL_FONT_STACK = "'맑은 고딕', Malgun Gothic, Calibri, sans-serif";
/** CellStyleAttrs.fs는 항상 px로 저장한다(기존 관례) — 툴바 입력창은 엑셀처럼 pt로 보여주므로
 * 여기서만 서로 변환한다. */
const PT_TO_PX = 96 / 72;
const PX_TO_PT = 72 / 96;

const params = new URLSearchParams(location.search);
const token = params.get('token') ?? '';

/** false면 이 문서 전체가 읽기 전용이다(예: 과거 연도 접속현황) — 잠금 여부(lk)와 무관하게 모든
 * 칸이 편집 불가여야 한다. main()에서 서버가 내려준 토큰 권한으로 설정된다. */
let canEdit = true;

/** renderTableVirtualized()가 다시 그릴 때마다 하나씩 늘어난다. 붙여넣기로 시트 범위가 넓어지거나
 * 스크롤로 가상 렌더가 다시 그리면 renderTableVirtualized()가 트랜잭션 도중 동기적으로 표 전체를
 * (또는 화면에 보이는 범위를) 다시 그리는데, 이때 포커스가 있던 셀이
 * DOM에서 떨어져 나가며 브라우저가 blur를 보낸다 — 그 blur 콜백이 실행되는 시점엔 이미 그 요소가
 * 실제로는 못 쓰게 됐어도 `isConnected`가 아직 true로 보이는 타이밍이 브라우저마다 있어(제거
 * 시점과 blur 디스패치 시점이 항상 완전히 동기적이지 않음) 그 값만으로는 안전하게 걸러낼 수 없다.
 * 대신 "이 칸이 만들어졌을 때의 세대"와 "지금 세대"를 비교하는 논리적 판정을 쓴다 — 사이에 표가
 * 다시 그려졌으면(세대가 달라졌으면) 그 blur는 이미 낡은 것이니 커밋하지 않는다. */
let renderGeneration = 0;

const titleEl = document.getElementById('doc-title') as HTMLElement;
const statusEl = document.getElementById('conn-status') as HTMLElement;
const presenceEl = document.getElementById('presence-list') as HTMLElement;
const warningEl = document.getElementById('unsupported-warning') as HTMLElement;
const editorHost = document.getElementById('grid-editor') as HTMLElement;
const zoomHost = document.getElementById('grid-zoom-host') as HTMLElement;
const tableEl = document.getElementById('grid-edit-table') as HTMLTableElement;
const scrollHostEl = document.querySelector('.grid-editor-scroll') as HTMLElement;
const sheetTabsEl = document.getElementById('sheet-tabs') as HTMLElement;
const cellRefBoxEl = document.getElementById('cell-ref-box') as HTMLElement;
const formulaInputEl = document.getElementById('formula-input') as HTMLInputElement;
const toolbar = document.getElementById('toolbar') as HTMLElement;
const errorScreen = document.getElementById('error-screen') as HTMLElement;
const errorMessage = document.getElementById('error-message') as HTMLElement;
const loadingScreen = document.getElementById('loading-screen') as HTMLElement;
const loadingPercentEl = document.getElementById('loading-percent') as HTMLElement;
const loadingGaugeFillEl = document.getElementById('loading-gauge-fill') as HTMLElement;

// 다른 화면에 iframe으로 임베드된 경우엔 부모 창이 이미 자기 스피너를 쓰고 있으므로(위
// g-render-pending 관련 주석 참고) 이 화면의 로딩 표시는 끈다 — 독립 접속일 때만 보인다.
const showsOwnLoadingScreen = window.parent === window;
if (!showsOwnLoadingScreen) {
  loadingScreen.hidden = true;
}

/**
 * 최초 동기화(서버가 canonical 그리드를 다시 구성하는 과정)는 진행률을 알려줄 방법이 없다 —
 * ExcelJS 자체의 압축 해제·XML 파싱이 대부분의 시간을 차지하는데 그 내부 진행 상황은 이
 * 라이브러리가 알려주지 않는다. 그래서 index.html의 업로드 변환 게이지(main.js의
 * CONVERT_CREEP_CAP/RATE)와 똑같은 방식 — 시간이 지날수록 상한(92%)에 점근하도록 부드럽게
 * 채워, "멈춘 게 아니라 진행 중"이라는 신호만 준다. 실제 완료(provider의 sync 이벤트)가
 * 오면 그 즉시 100%로 스냅한다.
 */
const LOADING_CREEP_CAP = 92;
const LOADING_CREEP_RATE = 0.04;
const LOADING_CREEP_INTERVAL_MS = 300;
let loadingDisplayed = 0;
let loadingCreepTimer: ReturnType<typeof setInterval> | null = null;
function setLoadingProgress(percent: number): void {
  loadingPercentEl.textContent = `${Math.round(percent)}%`;
  loadingGaugeFillEl.style.transform = `scaleX(${percent / 100})`;
}
if (showsOwnLoadingScreen) {
  loadingCreepTimer = setInterval(() => {
    loadingDisplayed = Math.min(LOADING_CREEP_CAP, loadingDisplayed + (LOADING_CREEP_CAP - loadingDisplayed) * LOADING_CREEP_RATE);
    setLoadingProgress(loadingDisplayed);
  }, LOADING_CREEP_INTERVAL_MS);
}
function stopLoadingCreep(): void {
  if (loadingCreepTimer !== null) {
    clearInterval(loadingCreepTimer);
    loadingCreepTimer = null;
  }
  setLoadingProgress(100);
}

function showFatalError(message: string): void {
  stopLoadingCreep();
  loadingScreen.hidden = true;
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

function colLetter(index: number): string {
  let n = index;
  let label = '';
  do {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return label;
}

function cellRefLabel(row: number, col: number): string {
  return `${colLetter(col)}${row + 1}`;
}

/** 엑셀처럼 숫자로 보이는 값은 서식이 따로 없으면 기본적으로 오른쪽 정렬로 보여준다. */
const NUMERIC_VALUE = /^-?\d+(\.\d+)?$/;

/** 엑셀 테두리 굵기(1=가는 선, 2=굵은 선)를 CSS로. 지정이 없으면 빈 문자열 —
 * 그러면 스타일시트의 기본 점선 격자(.grid-edit-table td)가 그대로 보인다. */
function borderCss(weight: 1 | 2 | undefined): string {
  return weight ? `${weight}px solid var(--v-text)` : '';
}

const VALIGN_TO_CSS: Record<'t' | 'm' | 'b', string> = { t: 'top', m: 'middle', b: 'bottom' };

function applyStyleToTd(td: HTMLTableCellElement, style: CellStyleAttrs | undefined): void {
  td.style.fontWeight = style?.b ? '700' : '400';
  td.style.fontStyle = style?.i ? 'italic' : 'normal';
  td.style.textDecoration = [style?.u ? 'underline' : '', style?.st ? 'line-through' : ''].filter(Boolean).join(' ');
  td.style.color = style?.col ?? '';
  td.style.background = style?.bg ?? '';
  td.style.fontFamily = style?.ff ? `"${style.ff}", ${DEFAULT_CELL_FONT_STACK}` : '';
  td.style.fontSize = style?.fs ? `${style.fs}px` : '';
  td.style.verticalAlign = style?.va ? VALIGN_TO_CSS[style.va] : '';
  // 자동 줄 바꿈 — 셀 높이는 고정이라(틀고정 sticky 좌표 계산이 고정 행 높이를 전제로 함) 줄이
  // 너무 많아지면 아래가 잘릴 수 있다. 그래도 한 줄에 다 몰아넣고 말줄임표로 자르는 것보다는
  // 실제로 줄바꿈되는 쪽이 "자동 줄 바꿈"이라는 기능 자체엔 맞다.
  td.style.whiteSpace = style?.wrap ? 'normal' : '';
  if (style?.a) {
    td.style.textAlign = ALIGN_TO_CSS[style.a];
  } else if (style?.nf) {
    td.style.textAlign = 'right'; // 숫자 서식이 있으면 서식이 적용된 텍스트(₩1,000 등)라 정규식 판정이 안 먹는다
  } else {
    td.style.textAlign = NUMERIC_VALUE.test((td.textContent ?? '').trim()) ? 'right' : '';
  }
  td.style.borderTop = borderCss(style?.bt);
  td.style.borderRight = borderCss(style?.br);
  td.style.borderBottom = borderCss(style?.bb);
  td.style.borderLeft = borderCss(style?.bl);
  td.classList.toggle('g-locked', style?.lk === 1);
}

function escapeXmlText(value: string): string {
  return value.replace(/[<>&'"]/g, (ch) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' })[ch] as string);
}

/** "2026-09-06 23:30:11" — toISOString은 UTC라 화면을 보는 사람의 실제 시계와 어긋나 보이므로
 * 로컬 시각 그대로 초 단위까지 조립한다. viewer.js/editor.ts의 동명 함수와 같은 포맷(별도
 * 번들이라 공유 불가). */
function formatWatermarkTimestamp(date: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ` +
    `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`
  );
}

/** 뷰어/docx 편집기와 같은 화면 열람용 워터마크를 xlsx 동시편집 화면에도 띄운다. */
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

  // canEdit:false여도 이 편집기 화면 자체는 그대로 보여준다(과거 연도 접속현황처럼 "현재 연도와
  // 똑같은 화면으로, 다만 손댈 수는 없게" 보여줘야 하는 경우가 있다) — 실제 쓰기 차단은 서버가
  // 읽기 전용 연결로 처리한다(collab/room.ts). 여기서는 서식 도구(리본)만 꺼서 애초에 사용자가
  // 고칠 수 있다고 착각하지 않게 한다 — 뒤로가기/제목/검색/참여자 같은 상단 줄은 도구가 아니라
  // 계속 보여야 하므로 toolbar 전체가 아니라 .toolbar-format만 숨긴다.
  canEdit = info.permissions.canEdit;
  if (!canEdit) {
    toolbar.querySelector('.toolbar-format')?.setAttribute('hidden', '');
    formulaInputEl.readOnly = true;
  }

  titleEl.textContent = info.name;
  document.title = `${info.name} — DocView 편집`;
  applyWatermark(info.viewer);

  const ydoc = new Y.Doc();
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const computed = ydoc.getMap<string>(GRID_COMPUTED_MAP);
  const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const merges = ydoc.getMap<readonly GridMergeShape[]>(GRID_MERGES_MAP);

  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, {
    params: { token },
  });

  // 이름 기준으로 색을 정한다(ydoc.clientID는 접속마다 무작위라 같은 사람이 다른 브라우저로
  // 들어오면 매번 색이 달라져 버린다) — 같은 이름이면 항상 같은 색.
  const localColor = pickPresenceColorForName(info.viewer);
  provider.awareness.setLocalStateField('user', { name: info.viewer, color: localColor });
  // 그리드 전체의 기본 --remote-color를 내 색으로 깔아 둔다 — 그러면 칸을 선택만 한 순간부터도
  // (편집모드에 들어가기 전부터) 이미 내 아바타 색 테두리로 보이고, 실제로 편집모드에 들어가도
  // 색이 안 바뀌니 따로 트랜지션을 신경 쓸 필요가 없다. 다른 참여자가 편집 중인 칸은 그 칸에 직접
  // 걸리는 인라인 --remote-color(renderRemoteCellHighlights)가 이 기본값보다 항상 우선한다.
  tableEl.style.setProperty('--remote-color', localColor);

  provider.on('status', (event: { status: string }) => {
    statusEl.textContent =
      event.status === 'connected' ? '연결됨' : event.status === 'disconnected' ? '연결 끊김 — 재연결 중…' : event.status;
    statusEl.dataset.status = event.status;
  });
  // 서버가 편집 세션 자체를 거부한 경우(edit.ts의 socket.close(4xxx, 이유)) — 4400대는 재시도해도
  // 같은 이유로 다시 거부될 뿐인 영구 거부다(잘못된 토큰, 문서 없음, 동시편집 룸 상한 등). 기본
  // 동작(지수 백오프로 계속 재연결 시도)을 멈추고 서버가 준 이유를 그대로 보여준다 — 특히 룸 상한
  // 거부("사용자가 많습니다…")는 그냥 두면 사용자가 원인을 알 길 없이 "연결 끊김 — 재연결 중…"만
  // 계속 보게 된다.
  provider.on('connection-close', (event: CloseEvent | null) => {
    if (event && event.code >= 4400) {
      provider.shouldConnect = false;
      showFatalError(event.reason || '서버가 연결을 거부했습니다.');
    }
  });
  provider.on('sync', (isSynced: boolean) => {
    if (isSynced) {
      statusEl.dataset.status = 'synced';
      renderSheetTabs();
      renderTableVirtualized();
      // 동기화 직후 structure.observe가 한 번 더 renderTableVirtualized()를 트리거해 폭이 살짝 바뀌는 경우가
      // 있다 — 그 재렌더까지 끝날 시간(두 번의 rAF, 즉 최소 한 번의 페인트)을 준 뒤에야 화면에
      // 드러내, 사용자가 "넓었다가 줄어드는" 깜빡임을 보지 않게 한다.
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          editorHost.classList.remove('g-render-pending');
          stopLoadingCreep();
          loadingScreen.hidden = true;
          // 부모 창(iframe으로 이 문서를 띄운 화면)이 자신의 로딩 스피너를 이 시점까지 유지하다가
          // 여기서 iframe으로 교체할 수 있도록 알려준다 — webviewer 자신은 스피너 UI를 새로 만들지
          // 않는다(부모가 이미 쓰고 있는 스피너를 그대로 재사용하는 게 목적).
          if (window.parent !== window) {
            window.parent.postMessage({ type: 'docview:ready' }, '*');
          }
        });
      });
    }
  });

  // ---------- 저장 상태 표시(Office 365 스타일 "저장 중… / 저장됨") + Ctrl+S ----------
  // docMeta.savedAt은 서버의 flushToStore가 실제로 디스크에 쓴 뒤에만 갱신하는 일반 Y.Map
  // 값이라, 이 클라이언트뿐 아니라 같이 편집 중인 다른 사람 화면에도 똑같이 반영된다.
  // 엑셀 하단 상태 표시줄의 "준비/입력/편집" 텍스트 — 지금 이 세션의 편집 동작 기준이라
  // 다른 참여자에게는 전파하지 않는다(각자 자기 화면 기준 상태).
  const gridStatusTextEl = document.getElementById('grid-status-text') as HTMLElement | null;
  function setGridStatusText(text: string): void {
    if (gridStatusTextEl) gridStatusTextEl.textContent = text;
  }
  if (!canEdit && gridStatusTextEl) gridStatusTextEl.textContent = '읽기 전용';

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

  /** 이 문서를 iframe으로 임베드한 부모 창(예: 업체별 접속현황 탭)이 자기 UI 안에 접속자 아바타를
   * 직접 그릴 수 있도록, webviewer 자신의 툴바(chrome=hidden이면 안 보임)와는 별개로 접속자 목록을
   * 그대로 postMessage로도 흘려보낸다 — webviewer는 그 목록을 "어디에" 보여줄지 전혀 몰라도 된다. */
  function postPresenceToParent(): void {
    if (window.parent === window) return;
    const states = [...provider.awareness.getStates().entries()] as Array<[number, AwarenessUserState]>;
    const users = dedupeByName(states).map(([, state]) => ({
      name: state.user!.name!,
      color: state.user!.color ?? pickPresenceColorForName(state.user!.name!),
    }));
    window.parent.postMessage({ type: 'docview:presence', users }, '*');
  }
  provider.awareness.on('change', () => {
    renderPresenceAvatars(presenceEl, provider.awareness);
    renderRemoteCellHighlights();
    postPresenceToParent();
  });
  renderPresenceAvatars(presenceEl, provider.awareness);
  postPresenceToParent();

  const undoManager = new Y.UndoManager([cells, structure, styles]);

  /**
   * 시트 자리 번호 → 마지막으로 관찰한 행/열 수. structure.observe가 "행/열 삽입·삭제가 실제로
   * 일어났다"를 시트 이름 변경·탭 순서 변경 같은 다른 구조 갱신과 구별하는 유일한 기준이다
   * (리뷰 지적 I-4의 undo 스택 무효화 참고).
   */
  const lastSeenSheetExtents = new Map<string, { rowCount: number; colCount: number }>();
  structure.forEach((s, key) => lastSeenSheetExtents.set(key, { rowCount: s.rowCount, colCount: s.colCount }));

  /**
   * 이번 structure 이벤트에서 기존 시트의 행/열 **수**가 바뀌었는지 — 확인하면서 기준값을 갱신한다.
   *
   * 처음 보는 키(협업 중 새로 추가된 시트)는 "바뀐 것"으로 세지 않는다: 시트는 항상 뒤에 붙고
   * 자리 번호가 밀리지 않으므로(room.ts의 removedSheetIndices 주석) 기존 셀 키가 가리키는 대상이
   * 달라지지 않는다.
   */
  function consumeSheetExtentChange(): boolean {
    let changed = false;
    structure.forEach((s, key) => {
      const prev = lastSeenSheetExtents.get(key);
      if (prev && (prev.rowCount !== s.rowCount || prev.colCount !== s.colCount)) changed = true;
      lastSeenSheetExtents.set(key, { rowCount: s.rowCount, colCount: s.colCount });
    });
    return changed;
  }

  let activeSheet = 0;
  /** 마지막으로 선택(focus)된 칸 — 편집 중이 아니어도 계속 유지된다(툴바 서식 버튼이 "마지막
   * 선택 칸"에 적용되어야 하므로). */
  let activeCellKey: string | null = null;
  /** 지금 실제로 편집 중(contenteditable=true, 텍스트 커서 보임)인 칸. 엑셀처럼 칸을 클릭만 하면
   * "선택"만 되고, F2·더블클릭·직접 타이핑을 해야 이 값이 채워지며 편집모드로 들어간다. */
  let editingKey: string | null = null;
  /** 편집 중이라 건너뛴 구조 변경 재렌더가 밀려 있는지 — structure.observe/flushPendingStructureRender 참고. */
  let pendingStructureRender = false;
  /** 다중 셀 드래그 선택의 두 꼭짓점 — 드래그 안 하고 그냥 칸 하나만 클릭/이동했으면 앵커=포커스로
   * 같아서(1×1 범위) 별도 배경 없이 기존 :focus 테두리만 보인다. */
  let selAnchorRow = 0;
  let selAnchorCol = 0;
  let selFocusRow = 0;
  let selFocusCol = 0;
  let isRangeDragging = false;

  /** 셀의 표시용 값을 얻는다 — 수식 셀이면 서버가 계산해 둔 값(computed 맵)을, 아니면 원본
   * 텍스트(cells 맵)를 그대로 쓴다. 클라이언트는 더 이상 수식을 직접 계산하지 않는다(서버가
   * canonical 전체 데이터 기준으로 계산해야 화면 밖 참조도 정확하기 때문 — 설계 문서의 핵심
   * 결정 참고). */
  function computeCellDisplay(sheetIdx: number, row: number, col: number): unknown {
    const key = gridCellKey(sheetIdx, row, col);
    const computedValue = computed.get(key);
    if (computedValue !== undefined) return computedValue;
    const raw = cells.get(key) ?? '';
    if (raw.startsWith('=')) return ''; // 서버 계산이 아직 도착 전(그 페이지가 막 로드된 직후) — 잠시 빈 값
    if (raw.trim() === '') return '';
    const num = Number(raw);
    return Number.isFinite(num) ? num : raw;
  }

  function formatComputedValue(value: unknown): string {
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (typeof value !== 'number') return value === null || value === undefined ? '' : String(value);
    // 부동소수점 오차(0.1+0.2=0.30000000000000004류)를 다듬는다
    return String(Math.round(value * 1e10) / 1e10);
  }

  /** 원본 값(또는 수식 계산 결과)에 셀의 숫자 표시 형식(쉼표/통화/백분율)을 입힌다 — 저장된 값
   * 자체는 절대 안 바꾸고 "화면에 어떻게 보일지"만 바꾼다(수식 계산 결과 표시와 같은 원리). */
  function formatNumberForDisplay(text: string, style: CellStyleAttrs | undefined): string {
    if (!style?.nf || text.trim() === '') return text;
    const num = Number(text);
    if (!Number.isFinite(num)) return text;
    if (num === 0) return ''; // 값 없음(0)은 화면을 어지럽히지 않게 빈 칸으로 보여준다
    const decimals = style.nd ?? 0;
    const opts = { minimumFractionDigits: decimals, maximumFractionDigits: decimals };
    switch (style.nf) {
      case 'percent':
        return `${(num * 100).toLocaleString('ko-KR', opts)}%`;
      case 'currency':
        return `₩${num.toLocaleString('ko-KR', opts)}`;
      case 'number':
        return num.toLocaleString('ko-KR', opts);
      default:
        return text;
    }
  }

  /** 셀 하나를 "지금 이 셀이 포커스 중인지"에 맞춰 원본 수식 또는 계산 결과로 그린다. */
  function renderCellDisplay(td: HTMLTableCellElement, key: string): void {
    const raw = cells.get(key) ?? '';
    if (editingKey === key) {
      // 편집 중에는 계산 결과가 아니라 실제 입력한 수식이 보여야 한다 — contentEditable 커서/선택
      // 범위가 span으로 감싸면 꼬일 수 있어 이 모드에서는 스필 span을 쓰지 않고 그대로 넣는다.
      td.replaceChildren();
      td.textContent = raw;
      return;
    }
    const style = styles.get(key);
    const text = !raw.startsWith('=')
      ? formatNumberForDisplay(raw, style)
      : formatNumberForDisplay(
          formatComputedValue(
            (() => {
              const parsed = parseGridCellKey(key);
              return parsed ? computeCellDisplay(parsed.sheetIdx, parsed.row, parsed.col) : raw;
            })(),
          ),
          style,
        );
    applySpillDisplay(td, key, text);
  }

  /** 실제 엑셀처럼 — 오른쪽 이웃 칸이 비어 있으면 이 칸의 표시 내용이 그 위로 넘쳐 보이게 한다
   * (값 있는 칸이나 병합 칸을 만나면 거기서 멈춘다). td.textContent를 직접 쓰지 않고 안쪽
   * span(.g-cell-spill)의 너비만 늘려야 td 경계 밖으로 시각적으로 넘칠 수 있다(CSS 쪽 참고). */
  function applySpillDisplay(td: HTMLTableCellElement, key: string, text: string): void {
    let span = td.querySelector<HTMLSpanElement>(':scope > .g-cell-spill');
    if (!span) {
      span = document.createElement('span');
      span.className = 'g-cell-spill';
      td.replaceChildren(span);
    }
    span.textContent = text;
    span.style.width = '';
    if (!text.trim() || td.colSpan > 1) return; // 병합 칸은 이미 colspan으로 넓어져 있어 스필이 필요 없다
    const parsed = parseGridCellKey(key);
    if (!parsed) return;
    const { sheetIdx, row, col } = parsed;
    const sheet = structure.get(String(sheetIdx));
    if (!sheet) return;
    const colWidthOf = (c: number) => sheet.colWidths?.[c] ?? DEFAULT_COL_WIDTH_PX;
    // renderCountsFor()는 scrollHostEl.clientWidth/clientHeight를 읽는데, 이 값들은 레이아웃이
    // 최신 상태여야만 정확해서 브라우저가 그 자리에서 강제로 레이아웃을 다시 계산한다(forced
    // synchronous layout). 스크롤 패치(updateScrollWindow)처럼 한 번에 수십~수백 칸을 새로 그리는
    // 도중 renderCellDisplay→applySpillDisplay가 칸마다 이걸 부르면, DOM을 막 바꾼 직후 레이아웃을
    // 다시 읽는 패턴(레이아웃 스래싱)이 칸 수만큼 반복되어 렌더 비용이 실측상 눈에 띄게 늘었다.
    // 지금 칸이 속한 시트가 마지막 실제 렌더(vw* 캐시)와 같다면 그 열 수·병합 조회표를 그대로
    // 쓴다 — 스필 폭은 어차피 "지금 화면에 그려진 레이아웃" 기준이라 매번 새로 재는 것보다 이
    // 캐시가 더 일관되게 맞다. (시트 전환 직후처럼 캐시가 아직 안 맞을 때만 예전처럼 새로 계산한다.)
    const useCache = vwValid && sheetIdx === vwSheetIdx;
    const colCount = useCache ? vwRenderColCount : renderCountsFor(sheet).cols;
    const { anchors, covered } = useCache ? { anchors: vwAnchors, covered: vwCovered } : buildMergeLookup(sheetIdx);
    let width = colWidthOf(col);
    for (let cc = col + 1; cc < colCount; cc += 1) {
      if (covered.has(`${row}:${cc}`) || anchors.has(`${row}:${cc}`)) break;
      if ((cells.get(gridCellKey(sheetIdx, row, cc)) ?? '').trim() !== '') break;
      width += colWidthOf(cc);
    }
    if (width > colWidthOf(col)) span.style.width = `${width}px`;
  }

  /** key가 가리키는 칸이 바뀌면, 그 칸 자신뿐 아니라 왼쪽에서 그 칸으로 넘쳐 들어오고 있었을 수
   * 있는 가장 가까운 "값 있는" 칸도 스필 폭을 다시 계산해야 한다(그 칸이 비었다가 채워지면 왼쪽
   * 칸의 넘침이 줄어들어야 하고, 반대로 비면 다시 늘어나야 한다). */
  function refreshSpillNeighbor(key: string): void {
    const parsed = parseGridCellKey(key);
    if (!parsed) return;
    const { sheetIdx, row, col } = parsed;
    const { covered } = buildMergeLookup(sheetIdx);
    for (let cc = col - 1; cc >= 0; cc -= 1) {
      if (covered.has(`${row}:${cc}`)) continue;
      const leftKey = gridCellKey(sheetIdx, row, cc);
      if ((cells.get(leftKey) ?? '').trim() === '') continue;
      const leftTd = tdFor(row, cc);
      if (leftTd && editingKey !== leftKey) renderCellDisplay(leftTd, leftKey);
      break;
    }
  }

  /** 어떤 셀이든 값이 바뀌면, 그 값을 참조할 수 있는 다른 수식 셀들도 다시 계산해 보여준다
   * (의존성 그래프 없이 "지금 그려진 수식 셀 전부 다시 계산"으로 단순하게 처리 — 편집기 규모에서
   * 충분히 빠르다). */
  function recalcFormulaCells(): void {
    // 합계(N열)처럼 수식이 들어있는 칸은 대개 잠겨 있어(.g-editable이 아님) 여기서 빠지면 값이
    // 바뀌어도 화면이 갱신되지 않는다 — 잠금 여부와 무관하게 "실제 칸"이면(data-row가 있으면) 전부 훑는다.
    tableEl.querySelectorAll<HTMLTableCellElement>('td[data-row]').forEach((td) => {
      const row = Number(td.dataset.row);
      const col = Number(td.dataset.col);
      const key = gridCellKey(activeSheet, row, col);
      if (editingKey === key) return;
      if ((cells.get(key) ?? '').startsWith('=')) renderCellDisplay(td, key);
    });
  }

  /** 화면 순서는 order 필드 기준(없으면 시트 키 번호로 대체) — 키 번호는 cells/styles/merges가
   * 참조하는 안정적인 식별자라 순서가 바뀌어도 그대로 두고, order만 바꿔서 재배치한다. */
  function sortedSheetIndices(): number[] {
    return [...structure.keys()]
      .map(Number)
      .sort((a, b) => (structure.get(String(a))?.order ?? a) - (structure.get(String(b))?.order ?? b));
  }
  function tabElFor(idx: number): HTMLElement | null {
    return sheetTabsEl.querySelector(`[data-sheet-idx="${idx}"]`);
  }
  function closeSheetPopup(): void {
    document.querySelectorAll('.sheet-context-menu, .sheet-move-menu, .grid-context-menu').forEach((el) => el.remove());
  }

  function renderSheetTabs(): void {
    const indices = sortedSheetIndices();
    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'grid-sheet-tab-add';
    addBtn.title = '시트 추가';
    addBtn.setAttribute('aria-label', '시트 추가');
    addBtn.innerHTML =
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>';
    addBtn.addEventListener('click', () => addSheet());
    sheetTabsEl.replaceChildren(
      ...indices.map((idx) => {
        const sheet = structure.get(String(idx))!;
        const tab = document.createElement('button');
        tab.type = 'button';
        tab.className = 'grid-sheet-tab' + (idx === activeSheet ? ' active' : '');
        tab.textContent = sheet.name;
        tab.dataset.sheetIdx = String(idx);
        tab.addEventListener('click', () => {
          if (activeSheet === idx) return;
          activeSheet = idx;
          resetSelectionForSheetSwitch();
          renderSheetTabs();
          renderTableVirtualized();
        });
        tab.addEventListener('dblclick', () => startSheetRename(idx));
        tab.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          openSheetContextMenu(idx, event.clientX, event.clientY);
        });
        return tab;
      }),
      addBtn,
    );
  }

  /** order 필드만 다시 매긴다 — 셀 데이터는 전혀 건드리지 않으므로 시트가 몇 개든 가벼운 연산이다. */
  function applySheetOrder(orderedKeys: readonly number[]): void {
    ydoc.transact(() => {
      orderedKeys.forEach((key, i) => {
        const sheet = structure.get(String(key));
        if (sheet && sheet.order !== i) structure.set(String(key), { ...sheet, order: i });
      });
    });
  }

  /** 엑셀의 "새 시트"/시트 탭 우클릭 "삽입"과 동일 — 기존 이름과 안 겹치는 "SheetN" 이름으로 빈
   * 시트를 만들어 바로 전환한다. afterIdx를 주면 그 시트 바로 뒤에, 안 주면 맨 뒤에 넣는다. */
  function addSheet(afterIdx?: number): void {
    const indices = sortedSheetIndices();
    const newKey = indices.length ? Math.max(...indices) + 1 : 0;
    const existingNames = new Set(indices.map((idx) => structure.get(String(idx))?.name));
    let n = indices.length + 1;
    while (existingNames.has(`Sheet${n}`)) n += 1;
    const insertPos = afterIdx === undefined ? indices.length : indices.indexOf(afterIdx) + 1;
    const orderedKeys = [...indices];
    orderedKeys.splice(insertPos, 0, newKey);
    activeSheet = newKey;
    resetSelectionForSheetSwitch();
    ydoc.transact(() => {
      structure.set(String(newKey), { name: `Sheet${n}`, rowCount: 1, colCount: 1, order: insertPos });
      orderedKeys.forEach((key, i) => {
        if (key === newKey) return;
        const sheet = structure.get(String(key));
        if (sheet && sheet.order !== i) structure.set(String(key), { ...sheet, order: i });
      });
    });
  }

  /** 엑셀 시트 탭 우클릭 "삭제" — 마지막 남은 시트는 지울 수 없고, 지금 보고 있던 시트를 지우면
   * 그 자리에 있던(원래 바로 뒤였던) 시트로 자동 전환한다. */
  function deleteSheet(idx: number): void {
    const indices = sortedSheetIndices();
    if (indices.length <= 1) {
      alert('마지막 남은 시트는 삭제할 수 없습니다.');
      return;
    }
    const sheet = structure.get(String(idx));
    if (!confirm(`"${sheet?.name ?? ''}" 시트를 삭제하시겠습니까? 이 시트의 모든 내용이 사라집니다.`)) return;
    const pos = indices.indexOf(idx);
    const remaining = indices.filter((i) => i !== idx);
    if (activeSheet === idx) {
      activeSheet = remaining[Math.min(pos, remaining.length - 1)];
      resetSelectionForSheetSwitch();
    }
    ydoc.transact(() => {
      structure.delete(String(idx));
      remaining.forEach((key, i) => {
        const s = structure.get(String(key));
        if (s && s.order !== i) structure.set(String(key), { ...s, order: i });
      });
    });
  }

  /** 엑셀 시트 탭 우클릭/더블클릭 "이름 바꾸기" — 탭 자리에 바로 입력창을 띄운다. */
  function startSheetRename(idx: number): void {
    const tab = tabElFor(idx);
    const sheet = structure.get(String(idx));
    if (!tab || !sheet) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'grid-sheet-tab-rename';
    input.value = sheet.name;
    tab.replaceWith(input);
    input.focus();
    input.select();
    const commit = () => {
      const trimmed = input.value.trim();
      if (trimmed && trimmed !== sheet.name) {
        structure.set(String(idx), { ...sheet, name: trimmed });
      } else {
        renderSheetTabs(); // 값이 그대로면 observe가 안 불려 렌더가 안 되므로 직접 되돌린다
      }
    };
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') input.blur();
      if (event.key === 'Escape') {
        input.value = sheet.name;
        input.blur();
      }
    });
    input.addEventListener('blur', commit, { once: true });
  }

  /** 엑셀 시트 탭 우클릭 "이동/복사"의 단순화된 버전 — 목록에서 기준 시트를 고르면 그 앞으로
   * 옮긴다(드래그 대신 클릭 한 번으로 재배치). */
  function openMoveMenu(idx: number, anchorEl: HTMLElement): void {
    closeSheetPopup();
    const indices = sortedSheetIndices();
    const menu = document.createElement('div');
    menu.className = 'sheet-move-menu';
    indices.forEach((targetIdx) => {
      if (targetIdx === idx) return;
      const targetSheet = structure.get(String(targetIdx));
      const item = document.createElement('button');
      item.type = 'button';
      item.textContent = `"${targetSheet?.name ?? ''}" 앞으로 이동`;
      item.addEventListener('click', () => {
        const reordered = indices.filter((i) => i !== idx);
        reordered.splice(reordered.indexOf(targetIdx), 0, idx);
        applySheetOrder(reordered);
        closeSheetPopup();
      });
      menu.appendChild(item);
    });
    const toEnd = document.createElement('button');
    toEnd.type = 'button';
    toEnd.textContent = '맨 뒤로 이동';
    toEnd.addEventListener('click', () => {
      applySheetOrder([...indices.filter((i) => i !== idx), idx]);
      closeSheetPopup();
    });
    menu.appendChild(toEnd);
    document.body.appendChild(menu);
    const anchorRect = anchorEl.getBoundingClientRect();
    positionSheetPopup(menu, anchorRect.left, anchorRect.bottom + 4, anchorRect.top - 4);
    setTimeout(() => document.addEventListener('mousedown', onSheetPopupOutsideClick), 0);
  }

  /** 뷰포트 아래쪽에 붙어 있는 시트 탭(항상 그렇다)에서 메뉴를 열면 아래로 펼칠 공간이 부족한
   * 경우가 대부분이라, 실제로 렌더링된 높이를 재서 넘치면 위로 뒤집는다. */
  function positionSheetPopup(menu: HTMLElement, x: number, belowY: number, aboveY: number = belowY): void {
    const rect = menu.getBoundingClientRect();
    const fitsBelow = belowY + rect.height <= window.innerHeight;
    const top = fitsBelow ? belowY : Math.max(4, aboveY - rect.height);
    const left = Math.max(4, Math.min(x, window.innerWidth - rect.width - 4));
    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;
  }

  function onSheetPopupOutsideClick(event: MouseEvent): void {
    const popups = document.querySelectorAll('.sheet-context-menu, .sheet-move-menu, .grid-context-menu');
    const inside = [...popups].some((el) => el.contains(event.target as Node));
    if (!inside) {
      closeSheetPopup();
      document.removeEventListener('mousedown', onSheetPopupOutsideClick);
    }
  }

  /** 엑셀 시트 탭 우클릭 메뉴와 동일한 4가지: 삽입/삭제/이름 바꾸기/이동. */
  function openSheetContextMenu(idx: number, x: number, y: number): void {
    closeSheetPopup();
    const menu = document.createElement('div');
    menu.className = 'sheet-context-menu';
    const items: Array<[string, () => void]> = [
      ['삽입', () => addSheet(idx)],
      ['삭제', () => deleteSheet(idx)],
      ['이름 바꾸기', () => startSheetRename(idx)],
      ['이동', () => openMoveMenu(idx, tabElFor(idx) ?? menu)],
    ];
    items.forEach(([label, action]) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        if (label !== '이동') closeSheetPopup();
        action();
      });
      menu.appendChild(btn);
    });
    document.body.appendChild(menu);
    positionSheetPopup(menu, x, y);
    setTimeout(() => document.addEventListener('mousedown', onSheetPopupOutsideClick), 0);
  }

  /**
   * 행/열 삽입·삭제 — 클라이언트는 **의도만 서버에 알리고** 실제 이동은 하지 않는다.
   *
   * **왜 직접 옮기지 않는가**(예전에는 shiftRowsAndMerges/shiftColsAndMerges가 Yjs의 cells/styles
   * 키를 한 칸씩 직접 옮겼다):
   *
   * 1. 그 방식은 "이 시트의 모든 셀이 Yjs에 들어 있다"를 전제한다. 지금 Yjs는 **구독 중인 500행
   *    페이지만 담는 창**이라(sendPageSubscription 참고) 창 밖의 행은 애초에 옮길 수가 없다 —
   *    삽입 지점 아래의 나머지 시트 전체가 위쪽과 한 칸씩 어긋난 채로 남는다.
   * 2. 그 루프의 상한은 시트 실제 크기가 아니라 **하드코딩된 1000**이었다(1000행 넘는 시트를
   *    지원하기 전 시대의 코드). 1001행이 넘는 시트에서는 행 1000의 내용을 덮어쓰고 1001행
   *    이후는 손대지도 않았다.
   *
   * 서버는 구독 상태와 무관하게 시트 전체를 담은 canonical 그리드를 들고 있으므로(room.ts의
   * applyGridStructureOp) 거기서 옮기면 두 문제가 동시에 사라진다. 이동 결과는 평소의 Yjs
   * 업데이트로 되돌아오고, 이미 있는 관찰자들이 화면을 맞춘다 — cells.observe가 바뀐 칸을 다시
   * 그리고, structure.observe가 새 행/열 수로 표를 다시 그린다(서버가 구조 메타데이터를 일부러
   * **마지막에** 보내므로 그 재렌더 시점엔 셀 데이터가 이미 도착해 있다).
   *
   * 알려진 대가: 행/열 삽입·삭제는 더 이상 Ctrl+Z(undoManager)로 되돌릴 수 없다 — UndoManager는
   * 로컬 트랜잭션만 추적하고 이 변경은 서버에서 온다. 삭제된 행의 내용을 되살리려면 서버 쪽
   * 역연산이 필요해 이번 범위 밖으로 둔다(예전 방식의 undo도 1000행 넘는 시트에서는 이미
   * 깨져 있었다).
   */
  /** room.ts의 MESSAGE_GRID_STRUCTURE_OP와 **같은 값이어야 한다**(0~3은 y-websocket 표준
   * 프로토콜이 이미 쓰고 있다 — 그쪽 주석 참고). */
  const MESSAGE_GRID_STRUCTURE_OP = 4;

  function sendGridStructureOp(
    kind: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols',
    at: number,
    count = 1,
  ): void {
    if (!canEdit) return;
    // sendPageSubscription과 완전히 같은 전송 방식(같은 WebSocket에 varUint 타입 + varString JSON).
    if (!provider.ws || provider.ws.readyState !== WebSocket.OPEN) return;
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_GRID_STRUCTURE_OP);
    encoding.writeVarString(encoder, JSON.stringify({ sheet: activeSheet, kind, at, count }));
    provider.ws.send(encoding.toUint8Array(encoder));
  }

  function insertRowAt(row: number): void {
    if (!structure.get(String(activeSheet))) return;
    sendGridStructureOp('insertRows', row);
  }
  function deleteRowAt(row: number): void {
    const sheet = structure.get(String(activeSheet));
    // 마지막 남은 행은 지우지 않는다 — 서버도 같은 검사를 하지만(applyGridStructureOp) 여기서
    // 먼저 막아 왕복 없이 곧바로 아무 일도 일어나지 않게 한다.
    if (!sheet || sheet.rowCount <= 1) return;
    sendGridStructureOp('deleteRows', row);
  }
  function insertColumnAt(col: number): void {
    if (!structure.get(String(activeSheet))) return;
    sendGridStructureOp('insertCols', col);
  }
  function deleteColumnAt(col: number): void {
    const sheet = structure.get(String(activeSheet));
    if (!sheet || sheet.colCount <= 1) return;
    sendGridStructureOp('deleteCols', col);
  }

  /** 선택한 열 기준 오름차순/내림차순 정렬 — 데이터 행(구조상 실제 rowCount) 전체를 그 열의
   * 값으로 재배열한다. 숫자로 보이면 숫자 비교, 아니면 문자열 비교(로캘 기준)로 — 엑셀의
   * "간단한 오름차순/내림차순 정렬"과 같은 수준의 단순한 구현이다. */
  function sortByColumn(col: number, direction: 'asc' | 'desc'): void {
    const sheet = structure.get(String(activeSheet));
    if (!sheet) return;
    const rows = Array.from({ length: sheet.rowCount }, (_, r) => r);
    const valueOf = (r: number) => cells.get(gridCellKey(activeSheet, r, col)) ?? '';
    rows.sort((a, b) => {
      const va = valueOf(a);
      const vb = valueOf(b);
      const na = NUMERIC_VALUE.test(va.trim()) ? Number(va) : null;
      const nb = NUMERIC_VALUE.test(vb.trim()) ? Number(vb) : null;
      const cmp = na !== null && nb !== null ? na - nb : va.localeCompare(vb, 'ko');
      return direction === 'asc' ? cmp : -cmp;
    });
    ydoc.transact(() => {
      // 원본을 전부 스냅샷으로 떠 둔 뒤(재배치 중 원본을 덮어써도 안전하게) 새 순서로 다시 쓴다.
      const snapshot = rows.map((r) =>
        Array.from({ length: sheet.colCount }, (_, c) => ({
          value: cells.get(gridCellKey(activeSheet, r, c)),
          style: styles.get(gridCellKey(activeSheet, r, c)),
        })),
      );
      snapshot.forEach((rowData, newR) => {
        rowData.forEach((cell, c) => {
          const key = gridCellKey(activeSheet, newR, c);
          if (cell.value !== undefined) cells.set(key, cell.value);
          else cells.delete(key);
          if (cell.style) styles.set(key, cell.style);
          else styles.delete(key);
        });
      });
    });
  }

  function openRowContextMenu(row: number, x: number, y: number): void {
    closeSheetPopup();
    const menu = document.createElement('div');
    menu.className = 'grid-context-menu sheet-context-menu';
    const items: Array<[string, () => void]> = [
      ['위에 행 삽입', () => insertRowAt(row)],
      ['행 삭제', () => deleteRowAt(row)],
    ];
    items.forEach(([label, action]) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        closeSheetPopup();
        action();
      });
      menu.appendChild(btn);
    });
    document.body.appendChild(menu);
    positionSheetPopup(menu, x, y);
    setTimeout(() => document.addEventListener('mousedown', onSheetPopupOutsideClick), 0);
  }

  function openColumnContextMenu(col: number, x: number, y: number): void {
    closeSheetPopup();
    const menu = document.createElement('div');
    menu.className = 'grid-context-menu sheet-context-menu';
    const items: Array<[string, () => void]> = [
      ['왼쪽에 열 삽입', () => insertColumnAt(col)],
      ['열 삭제', () => deleteColumnAt(col)],
      ['오름차순 정렬', () => sortByColumn(col, 'asc')],
      ['내림차순 정렬', () => sortByColumn(col, 'desc')],
    ];
    items.forEach(([label, action]) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        closeSheetPopup();
        action();
      });
      menu.appendChild(btn);
    });
    document.body.appendChild(menu);
    positionSheetPopup(menu, x, y);
    setTimeout(() => document.addEventListener('mousedown', onSheetPopupOutsideClick), 0);
  }

/** 셀 근처에 잠깐 떴다 사라지는 안내 말풍선 — 유효성 검사에 걸린 입력을 왜 안 받았는지 알려준다. */
  function showCellTooltip(td: HTMLTableCellElement, message: string): void {
    const rect = td.getBoundingClientRect();
    const tip = document.createElement('div');
    tip.className = 'g-cell-tooltip';
    tip.textContent = message;
    tip.style.left = `${rect.left}px`;
    tip.style.top = `${rect.bottom + 4}px`;
    document.body.appendChild(tip);
    setTimeout(() => tip.remove(), 2200);
  }

  /** 데이터가 없는 빈 셀에 값을 입력하면 시트 범위(structure)를 자동으로 넓힌다 — 엑셀처럼
   * "칸이 있으니까 그냥 입력하면 늘어난다"는 동작을 흉내낸다. */
  function commitCell(td: HTMLTableCellElement): void {
    const row = Number(td.dataset.row);
    const col = Number(td.dataset.col);
    const key = gridCellKey(activeSheet, row, col);
    const value = td.textContent ?? '';
    const current = cells.get(key) ?? '';
    if (value === current) return;

    // 이 칸이 숫자 전용(엑셀 데이터 유효성 검사)이면 숫자가 아닌 입력은 저장하지 않고 안내만 띄운다
    // — blur 이후 renderCellDisplay가 다시 그리면서 원래 값으로 자연히 되돌아간다.
    if (styles.get(key)?.nu === 1 && value.trim() !== '' && !NUMERIC_VALUE.test(value.trim())) {
      showCellTooltip(td, '숫자만 입력할 수 있습니다.');
      return;
    }

    ydoc.transact(() => {
      if (value === '') {
        cells.delete(key);
        return;
      }
      cells.set(key, value);
      const s = currentStructure();
      if (row >= s.rowCount || col >= s.colCount) {
        structure.set(String(activeSheet), {
          ...s,
          rowCount: Math.max(s.rowCount, row + 1),
          colCount: Math.max(s.colCount, col + 1),
        });
      }
    });
  }

  function isPasteTargetEditable(row: number, col: number): boolean {
    if (!canEdit) return false;
    return styles.get(gridCellKey(activeSheet, row, col))?.lk !== 1;
  }

  /** 실제 엑셀에서 여러 칸을 복사해 붙여넣을 때를 처리한다. 이 <td>가 contentEditable이라 아무
   * 처리도 안 하면 브라우저가 클립보드의 HTML(엑셀이 함께 넣어주는 <table> 조각)을 이 한 칸 안에
   * 그대로 밀어 넣어버려 값이 한 칸에 뭉쳐 들어가거나 깨진다 — 붙여넣기를 직접 가로채
   * 클립보드의 plain text(엑셀은 항상 탭=열, 줄바꿈=행 구분의 TSV도 같이 넣어준다)를 파싱해
   * 시작 칸부터 각 칸에 정확히 나눠 넣는다. */
  function handlePaste(event: ClipboardEvent, anchorTd: HTMLTableCellElement): void {
    const text = event.clipboardData?.getData('text/plain');
    if (!text) return;
    event.preventDefault();

    const startRow = Number(anchorTd.dataset.row);
    const startCol = Number(anchorTd.dataset.col);
    const rows = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    while (rows.length > 1 && rows[rows.length - 1] === '') rows.pop();
    const grid = rows.map((line) => line.split('\t'));
    const maxCols = Math.max(...grid.map((r) => r.length));

    ydoc.transact(() => {
      grid.forEach((values, rOffset) => {
        values.forEach((rawValue, cOffset) => {
          const row = startRow + rOffset;
          const col = startCol + cOffset;
          if (!isPasteTargetEditable(row, col)) return; // 잠긴 칸·읽기 전용 문서는 조용히 건너뛴다
          const key = gridCellKey(activeSheet, row, col);
          const value = rawValue.trim();
          if (styles.get(key)?.nu === 1 && value !== '' && !NUMERIC_VALUE.test(value)) return; // 숫자 전용 칸에 숫자가 아니면 건너뜀
          if (value === '') cells.delete(key);
          else cells.set(key, value);
        });
      });
      const s = currentStructure();
      const maxRow = startRow + grid.length - 1;
      const maxCol = startCol + maxCols - 1;
      if (maxRow >= s.rowCount || maxCol >= s.colCount) {
        structure.set(String(activeSheet), {
          ...s,
          rowCount: Math.max(s.rowCount, maxRow + 1),
          colCount: Math.max(s.colCount, maxCol + 1),
        });
      }
    });

    grid.forEach((values, rOffset) => {
      values.forEach((_value, cOffset) => {
        const row = startRow + rOffset;
        const col = startCol + cOffset;
        const targetTd = tdFor(row, col);
        if (targetTd) renderCellDisplay(targetTd, gridCellKey(activeSheet, row, col));
      });
    });
    recalcFormulaCells();
  }

  /** 데이터가 있는 범위 너머로도 빈 칸을 넉넉히 그려서, 실제 엑셀처럼 "빈 시트도 격자가 있는"
   * 느낌을 준다 — 그 빈 칸에 입력하면 commitCell이 자동으로 범위를 넓혀준다. */
  const GRID_ROW_PADDING = 15;
  const GRID_COL_PADDING = 6;
  const GRID_MIN_RENDER_ROWS = 25;
  const GRID_MIN_RENDER_COLS = 12;
  /** spreadsheet.ts의 DEFAULT_COL_WIDTH_PX와 같은 값 — 클라이언트 번들에 ExcelJS를 끌어오는
   * spreadsheet.ts를 직접 import할 수 없어(gridShared.ts와 동일한 이유) 값만 맞춰 둔다. */
  const DEFAULT_COL_WIDTH_PX = 100;
  /** .g-edit-rownum의 min-width와 맞춘 대략치 — 아래 tableEl.style.width 계산에만 쓰이고, 실제
   * 행번호 열 폭은 여전히 CSS(min-width)가 정한다. 정확히 안 맞아도 상관없다: 테이블에 "auto가
   * 아닌" width를 준다는 사실 자체가 중요하지, 합계가 1px이라도 틀리면 안 되는 게 아니다. */
  const ROWNUM_COL_WIDTH_PX = 28;
  /** 실제 렌더된 <td>의 border-box 높이(1.7rem + 위아래 패딩 0.18rem*2, 대략 27~28px)와 맞춘
   * 근사치 — 화면을 채우는 데 필요한 행 수를 계산하는 용도라 1px 오차는 문제 되지 않는다. */
  const ROW_HEIGHT_PX = 27;
  /** 가상 스크롤(윈도잉 렌더링)이 화면 위아래로 실제 뷰포트보다 더 그려두는 여유 행 수 — 스크롤이
   * 살짝만 움직여도 매 프레임 다시 그릴 필요 없게 완충 역할을 한다(너무 작으면 빠르게 스크롤할 때
   * 빈 줄이 잠깐 보이고, 너무 크면 DOM 노드가 늘어나 다시 느려진다). */
  const OVERSCAN_ROWS = 50;

  /** 시트 크기만으로는 "확대율을 낮추면 화면에 빈 칸이 더 많이 보여야 한다"(실제 엑셀은 격자가
   * 사실상 무한이라 화면에 보이는 만큼 항상 그려진다)는 걸 반영 못 한다 — 지금 뷰포트/줌 배율로
   * 실제 보이는 픽셀 영역을 재서, 그걸 채우는 데 필요한 행/열 수까지 함께 반영한다. CSS zoom은
   * 확대율이 낮을수록 이 컨테이너의 clientWidth/clientHeight를(줌이 적용되기 전 좌표계 기준으로)
   * 더 크게 보고하므로, 별도로 zoom 값을 나눌 필요 없이 그대로 쓰면 된다. */
  function renderCountsFor(sheet: SheetStructure): { rows: number; cols: number } {
    const availWidth = scrollHostEl.clientWidth;
    const availHeight = scrollHostEl.clientHeight;
    let usedWidth = ROWNUM_COL_WIDTH_PX;
    let fillCols = 0;
    while (usedWidth < availWidth && fillCols < 1000) {
      usedWidth += sheet.colWidths?.[fillCols] ?? DEFAULT_COL_WIDTH_PX;
      fillCols += 1;
    }
    const fillRows = Math.max(0, Math.ceil(availHeight / ROW_HEIGHT_PX) - 1);
    return {
      rows: Math.max(sheet.rowCount + GRID_ROW_PADDING, GRID_MIN_RENDER_ROWS, fillRows),
      cols: Math.max(sheet.colCount + GRID_COL_PADDING, GRID_MIN_RENDER_COLS, fillCols),
    };
  }

  /** 병합 정보(1-based)를 0-based (row:col) 조회표로 바꾼다 — viewer.js의 buildMergeIndex와 동일한 원리.
   * anchors: 병합의 왼쪽위 칸 → {rs,cs}(rowspan/colspan). covered: 그 외 병합에 덮인 칸(<td>를 아예 안 그림). */
  function buildMergeLookup(sheetIdx: number): {
    anchors: Map<string, { rs: number; cs: number }>;
    covered: Set<string>;
  } {
    const anchors = new Map<string, { rs: number; cs: number }>();
    const covered = new Set<string>();
    const sheetMerges = merges.get(String(sheetIdx)) ?? [];
    for (const m of sheetMerges) {
      const r0 = m.r - 1;
      const c0 = m.c - 1;
      anchors.set(`${r0}:${c0}`, { rs: m.rs, cs: m.cs });
      for (let r = r0; r < r0 + m.rs; r += 1) {
        for (let c = c0; c < c0 + m.cs; c += 1) {
          if (r === r0 && c === c0) continue;
          covered.add(`${r}:${c}`);
        }
      }
    }
    return { anchors, covered };
  }

  /** ---------- 스크롤 전용 증분 패치 상태 ----------
   * renderTableVirtualized()는 스크롤이 한 번 일어날 때마다 <colgroup>/<thead>/보이는 모든
   * <tr>·<td>를 통째로 버리고 다시 만들었다 — 각 <td>에 붙는 5~8개의 이벤트 리스너(focus/
   * mousedown/blur/dblclick/input/paste/compositionstart/keydown)까지 매 프레임 떼고 다시
   * 붙이는 셈이라, 실측(Chrome longtask) 기준 한 번에 54~108ms가 걸렸다(16ms 프레임 예산의
   * 3~6배). 스크롤만으로 인한 재호출은 "화면에 보이는 행 창(window)"만 바뀔 뿐 시트/구조/
   * 틀고정/병합/열 폭은 그대로이므로, updateScrollWindow()가 직전에 그렸던 행 창과 새 행 창을
   * 비교해 새로 들어온 행만 만들고 나간 행만 지우고 나머지는 기존 DOM 노드(와 거기 걸린 리스너)를
   * 그대로 재사용한다. 시트 전환·구조 변경·확대축소처럼 실제로 다시 그려야 하는 경우는 여전히
   * renderTableVirtualized()의 전체 재빌드를 그대로 타므로(sync/시트탭 클릭/structure.observe/
   * scheduleViewportRefit 등은 지금처럼 renderTableVirtualized를 직접 부른다) 정확성 리스크가
   * 없다 — 이 캐시는 "지금 화면에 그려진 것과 정확히 같은 문맥(시트/열 폭/틀고정/병합)에서
   * 스크롤만 됐다"는 게 보장된 경우에만 쓰이고, 그 전제가 조금이라도 의심되면(vwValid===false거나
   * 시트/열·행 수가 다르면) 곧바로 전체 재빌드로 물러난다. */
  let vwValid = false;
  let vwSheetIdx = -1;
  let vwRenderColCount = 0;
  let vwRenderRowCount = 0;
  let vwFreezeRows = 0;
  let vwFreezeCols = 0;
  let vwAnchors: Map<string, { rs: number; cs: number }> = new Map();
  let vwCovered: Set<string> = new Set();
  let vwFirstVisible = -1;
  let vwLastVisible = -1;
  let vwTbody: HTMLTableSectionElement | null = null;
  let vwTopSpacer: HTMLTableRowElement | null = null;
  let vwBottomSpacer: HTMLTableRowElement | null = null;
  /** r(0-based 행 번호) → 지금 DOM에 있는 그 행의 <tr> — 틀 고정 행(0..freezeRows-1)과 지금
   * 스크롤 창 안의 행이 모두 들어있다(스페이서는 여기 포함되지 않는다 — 리스너가 없어 재사용할
   * 이유가 없으므로 updateScrollWindow는 스페이서를 매번 새로 만든다). */
  let vwRowEls: Map<number, HTMLTableRowElement> = new Map();

  /** 화면 밖(위/아래) 구간 전체를 대신 차지하는 "안 보이는" 스페이서 한 행 — buildRow 설명 참고.
   * rowCount<=0(그릴 게 없는 구간)이면 아예 만들지 않는다. 리스너가 없는 단순한 <tr>/<td> 한 쌍이라
   * updateScrollWindow()도 재사용하지 않고 매번 새로 만든다(재사용을 위한 위치 계산 비용이 굳이
   * 만드는 비용보다 크다). */
  function buildSpacerRow(rowCount: number): HTMLTableRowElement | null {
    if (rowCount <= 0) return null;
    const heightPx = `${rowCount * ROW_HEIGHT_PX}px`;
    const tr = document.createElement('tr');
    tr.setAttribute('aria-hidden', 'true');
    tr.style.height = heightPx;
    const td = document.createElement('td');
    td.colSpan = vwRenderColCount + 1; // 행번호 열까지 포함해 표 전체 폭을 덮는다
    // .grid-edit-table th/td에 공통으로 걸린 height:1.7rem(고정 행 높이용)을 인라인 스타일로
    // 덮어써야 스페이서가 의도한 실제 높이(수천 px일 수 있음)로 그려진다 — 인라인 스타일이
    // 클래스 규칙보다 항상 이긴다.
    td.style.height = heightPx;
    td.style.padding = '0';
    td.style.border = 'none';
    tr.appendChild(td);
    return tr;
  }

  /** 행 하나(<tr>)를 만든다 — 셀 생성·서식·이벤트 바인딩은 예전 renderTable과 동일하다. 이제는
   * 틀 고정 행이든 스크롤 윈도우 안의 행이든 항상 정상 흐름(<colgroup>이 적용되는 진짜 표 행)으로
   * 그린다 — 화면에 없는 구간은 이 함수가 아니라 buildSpacerRow가 대신 높이를 차지한다.
   * renderTableVirtualized()의 전체 재빌드와 updateScrollWindow()의 증분 패치 양쪽에서 함께
   * 쓰므로, 매개변수 대신 vw*(직전 렌더가 채워 둔) 캐시를 읽는다 — 두 호출부 모두 이 함수를 부르기
   * 전에 그 캐시가 지금 그리려는 행과 같은 문맥으로 맞춰져 있음을 보장한다. */
  function buildRow(r: number): HTMLTableRowElement {
    const tr = document.createElement('tr');
    // Chromium은 서로 다른 <tr> 소속인 position:sticky 셀들 사이의 z-index 비교를 신뢰할 수 없게
    // 처리한다(고정된 행의 배경보다 스크롤된 다른 행의 셀이 위로 그려지는 버그) — 각 행을
    // position+z-index로 그 자체의 스태킹 컨텍스트로 격리하면, 셀별 z-index는 그 행 안에서만
    // 비교되고 행 간 우열은 이 tr의 z-index로만 결정되어 늘 정확하다.
    tr.style.position = 'relative';
    tr.style.zIndex = r < vwFreezeRows ? '2' : '1';
    const rowHead = document.createElement('th');
    rowHead.className = 'g-edit-rownum';
    rowHead.textContent = String(r + 1);
    rowHead.dataset.row = String(r);
    rowHead.addEventListener('click', () => selectRow(r));
    if (canEdit) {
      rowHead.addEventListener('contextmenu', (event) => {
        event.preventDefault();
        openRowContextMenu(r, event.clientX, event.clientY);
      });
    }
    tr.appendChild(rowHead);
    for (let c = 0; c < vwRenderColCount; c += 1) {
      if (vwCovered.has(`${r}:${c}`)) continue; // 병합에 덮인 칸 — 앵커 쪽에서 colspan/rowspan으로 대신 차지한다
      const key = gridCellKey(activeSheet, r, c);
      const td = document.createElement('td');
      const span = vwAnchors.get(`${r}:${c}`);
      if (span) {
        if (span.rs > 1) td.rowSpan = span.rs;
        if (span.cs > 1) td.colSpan = span.cs;
      }
      const isLocked = styles.get(key)?.lk === 1;
      const canEditThisCell = !isLocked && canEdit;
      // 클릭만으로는 "선택"만 되고 편집모드로 들어가지 않는다(엑셀과 동일) — 실제 편집모드는
      // enterEditMode()가 필요할 때만 이 값을 'true'로 바꾼다. 잠긴 칸/읽기전용 문서는 애초에
      // canEditThisCell이 false라 keydown 핸들러가 편집 진입 자체를 막는다.
      td.contentEditable = 'false';
      // 잠긴 칸도 화살표 키로 지나가거나 선택해서 볼 수 있어야 하므로(엑셀도 보호된 칸을 선택은
      // 할 수 있음) 편집 가능 여부와 무관하게 모든 칸을 포커스 가능하게 만든다. tabIndex=-1은 Tab
      // 키 순서에는 안 끼지만 클릭·프로그램적 focus()는 그대로 받는다.
      td.tabIndex = -1;
      if (canEditThisCell) td.classList.add('g-editable');
      td.dataset.row = String(r);
      td.dataset.col = String(c);
      // 엑셀 "틀 고정" — 위 N행/왼쪽 N열은 스크롤해도 계속 보여야 한다. 실제 sticky 좌표(top/left)는
      // 렌더링이 끝나 실제 셀 크기를 잴 수 있게 된 뒤 applyFreezePanes()에서 한 번에 계산한다.
      if (r < vwFreezeRows) td.classList.add('g-frozen-row');
      if (c < vwFreezeCols) td.classList.add('g-frozen-col');
      renderCellDisplay(td, key);
      applyStyleToTd(td, styles.get(key));
      // 잠겨 있으면서 맨 위 행에 있고 내용이 실제로 있는 칸(제목 배너, "조회쿼리" 버튼 등)은 클릭
      // 시 부모 창(이 문서를 iframe으로 임베드한 화면)에 알림만 보낸다 — 무엇을 보여줄지(예: 연도
      // 목록 드롭다운, 쿼리 팝업)는 전적으로 부모 쪽 책임이다. webviewer 자신은 "연도"/"쿼리" 같은
      // 특정 문서의 의미를 전혀 알 필요가 없고, 클릭된 칸의 위치·텍스트만 그대로 전달한다.
      const hasContent = (cells.get(key) ?? '').trim() !== '';
      if (isLocked && r === 0 && hasContent) {
        td.classList.add('g-clickable-header');
        // 맨 왼쪽(제목 배너)만 "목록에서 고른다"는 의미의 드롭다운 화살표를 보여준다 — 그 옆의
        // 다른 클릭 가능한 칸(조회쿼리 버튼 등)은 텍스트 자체에 "(클릭)"을 넣어 알려주므로
        // 화살표가 필요 없다.
        if (c === 0) td.classList.add('g-header-dropdown');
        td.addEventListener('click', () => {
          if (window.parent !== window) {
            // 셀의 화면상 위치(이 iframe 안에서의 좌표)를 같이 보내, 부모 쪽이 임베드한 iframe의
            // 위치와 합쳐 팝오버를 셀 바로 밑에 자연스럽게 놓을 수 있게 한다 — webviewer는 "연도
            // 팝오버를 어디에 그릴지" 자체는 몰라도 되고, 그저 클릭된 칸의 위치만 알려준다.
            const rect = td.getBoundingClientRect();
            window.parent.postMessage(
              {
                type: 'docview:title-click',
                row: r,
                col: c,
                text: cells.get(key) ?? '',
                rect: { top: rect.top, left: rect.left, bottom: rect.bottom, right: rect.right, width: rect.width, height: rect.height },
              },
              '*',
            );
          }
        });
      }
      // 이 칸이 어느 "세대"(renderGeneration)에 만들어졌는지 기록해 둔다 — 아래 delegated
      // focusout의 stale 판정에 쓰인다: 스크롤 패치(updateScrollWindow)는 살아남는 칸을 다시
      // 만들지 않으므로 renderGeneration을 새로 올리지 않고, 그 사이 전체 재빌드가 한 번도
      // 없었다면 이 값은 여전히 "지금" 세대와 같다 — 즉 이 판정은 "이 특정 DOM 노드가 방금
      // 만들어졌는가"가 아니라 "마지막 전체 재빌드 이후로 편집 세션이 안전한가"를 묻는다.
      td.dataset.gen = String(renderGeneration);
      tr.appendChild(td);
    }
    return tr;
  }

  /** 이벤트 위임(delegation)에서 실제 <td>를 찾는다 — 화면에 보일 때는 스필 span(g-cell-spill)이
   * <td> 안에 들어있어 event.target이 그 span일 수 있으므로 closest로 조상을 거슬러 올라간다. */
  function cellFromEvent(event: Event): { td: HTMLTableCellElement; row: number; col: number; key: string } | null {
    const target = event.target as HTMLElement | null;
    const td = target?.closest<HTMLTableCellElement>('td[data-row][data-col]') ?? null;
    if (!td) return null;
    const row = Number(td.dataset.row);
    const col = Number(td.dataset.col);
    return { td, row, col, key: gridCellKey(activeSheet, row, col) };
  }

  /**
   * <td>마다 5~8개씩(focus/mousedown/blur/dblclick/input/paste/compositionstart/keydown) 걸던
   * 리스너를 표 전체에 단 한 번만 거는 위임 리스너로 대체한다 — buildRow()가 스크롤 패치
   * (updateScrollWindow)로 자주 다시 불리므로, 리스너를 셀마다 붙였다 뗐다 하는 비용 자체를 렌더
   * 경로에서 없앤다. 각 핸들러는 event.target에서 실제 칸을 역추적해(cellFromEvent) 예전에
   * r/c/key/canEditThisCell 클로저로 갖고 있던 정보를 그때그때 다시 구한다(canEditThisCell은
   * buildRow가 같은 시점에 붙여 둔 g-editable 클래스로 복원한다).
   */
  tableEl.addEventListener('focusin', (event) => {
    const hit = cellFromEvent(event);
    if (!hit) return;
    const { td, row: r, col: c, key } = hit;
    const canEditThisCell = td.classList.contains('g-editable');
    activeCellKey = key;
    selAnchorRow = r;
    selAnchorCol = c;
    selFocusRow = r;
    selFocusCol = c;
    renderRangeHighlight();
    syncCellToolbarState(key);
    setActiveHeaders(r, c);
    cellRefBoxEl.textContent = cellRefLabel(r, c);
    formulaInputEl.value = cells.get(key) ?? '';
    positionFillHandle(canEditThisCell ? td : null);
  });
  // 다중 셀 드래그 선택 시작 — 편집 중인 칸 안에서는 텍스트 드래그 선택이 우선이라 건드리지 않는다.
  tableEl.addEventListener('mousedown', (event) => {
    if (event.button !== 0) return;
    const hit = cellFromEvent(event);
    if (!hit) return;
    const { td, row: r, col: c, key } = hit;
    if (editingKey === key) return;
    // 기본 동작을 그대로 두면 브라우저가 "텍스트 드래그 선택"도 같이 시작해버린다 — 그건 사각형이
    // 아니라 문서(행) 순서로 훑어서, 계산한 범위보다 훨씬 넓게(중간 행의 다른 칸까지) 파랗게/회색으로
    // 잡혀 보인다. preventDefault로 그 기본 동작을 끄고, 그때 같이 꺼지는 포커스만 아래에서 직접
    // 다시 걸어준다.
    event.preventDefault();
    clearHeaderSelection();
    selAnchorRow = r;
    selAnchorCol = c;
    selFocusRow = r;
    selFocusCol = c;
    isRangeDragging = true;
    renderRangeHighlight();
    td.focus();
    document.addEventListener('mousemove', onRangeDragMove);
    document.addEventListener('mouseup', onRangeDragEnd);
  });
  // focus의 위임 버전인 focusin과 대칭으로, blur 대신 focusout(버블링됨)에 건다 — 이미 아래에
  // "그리드 영역을 완전히 벗어났는지"만 보는 별개의 focusout 리스너가 있고, 이건 그것과 별도로
  // "이 칸이 편집 중이었다면 커밋한다"만 담당한다.
  tableEl.addEventListener('focusout', (event) => {
    const hit = cellFromEvent(event);
    if (!hit) return;
    const { td, key } = hit;
    if (editingKey !== key) return; // 편집 중이 아니었으면(그냥 선택만) 커밋할 것이 없다
    // 붙여넣기/자동채우기로 시트 범위(rowCount/colCount)가 늘어나거나 스크롤로 가상 렌더가 다시
    // 그리면 renderTableVirtualized()가 트랜잭션 도중 동기적으로 표 전체를 다시 그리며 이 <td>를
    // DOM에서 떼어내는데, 포커스가 있던 칸이 떨어져 나가면 브라우저가 focusout을 보낸다 — 그
    // 시점의 td.textContent는 그 작업 "전" 값이라, 그대로 commitCell을 부르면 방금 쓴 값을
    // 덮어써 버린다. 제거 시점과 focusout 디스패치 시점이 완전히 동기적이진 않아 isConnected만으로는
    // 못 거르므로, "이 칸이 만들어진 세대"가 지금도 최신인지로 판정한다(buildRow의 data-gen 참고).
    const cellGeneration = Number(td.dataset.gen);
    if (cellGeneration !== renderGeneration) {
      editingKey = null;
      provider.awareness.setLocalStateField('cell', null);
      return;
    }
    commitCell(td);
    editingKey = null;
    provider.awareness.setLocalStateField('cell', null);
    td.contentEditable = 'false';
    renderCellDisplay(td, key); // 편집 종료 — 수식이면 계산 결과로 되돌린다
    recalcFormulaCells(); // 이 셀 값이 바뀌었으니 이를 참조하는 다른 수식 셀도 갱신
    setGridStatusText('준비');
    flushPendingStructureRender(); // 편집 중이라 미뤄 뒀던 구조 변경이 있으면 지금 반영한다
  });
  tableEl.addEventListener('dblclick', (event) => {
    const hit = cellFromEvent(event);
    if (!hit || !hit.td.classList.contains('g-editable')) return;
    enterEditMode(hit.td, hit.key);
  });
  tableEl.addEventListener('input', (event) => {
    const hit = cellFromEvent(event);
    if (!hit || !hit.td.classList.contains('g-editable')) return;
    if (activeCellKey === hit.key) formulaInputEl.value = hit.td.textContent ?? '';
  });
  tableEl.addEventListener('paste', (event) => {
    const hit = cellFromEvent(event);
    if (!hit || !hit.td.classList.contains('g-editable')) return;
    handlePaste(event, hit.td);
  });
  // 셀을 선택만 한 채(더블클릭 없이) 바로 한글 등 IME 문자를 치는 경우 — 조합 중인 글자는 keydown의
  // event.key가 길이 1이 아니라("Process" 등) handleCellKeydown의 "글자 하나 치면 편집모드 진입"
  // 분기를 못 타서, 셀이 그대로 선택 상태에 머물며 입력이 통째로 사라졌었다. compositionstart
  // 시점에 직접 편집모드로 들어가면(내용은 비워서 시작 — 엑셀처럼 선택된 채 타이핑하면 기존 값을
  // 덮어쓴다) 그 뒤 IME 조합이 이 칸에서 이어진다.
  tableEl.addEventListener('compositionstart', (event) => {
    const hit = cellFromEvent(event);
    if (!hit || !hit.td.classList.contains('g-editable')) return;
    if (editingKey !== hit.key) enterEditMode(hit.td, hit.key, '');
  });
  tableEl.addEventListener('keydown', (event) => {
    const hit = cellFromEvent(event);
    if (!hit) return;
    const canEditThisCell = hit.td.classList.contains('g-editable');
    handleCellKeydown(event, hit.td, hit.row, hit.col, hit.key, canEditThisCell);
  });

  /** 전체 행만큼 <tr>을 만들던 예전 renderTable()을 대체한다 — 시트가 아무리 커도 실제 DOM에는
   * "지금 보이는 범위 + 여유분(OVERSCAN_ROWS)"만큼의 행만 만든다.
   *
   * (이전 시도였던 position:absolute + top 방식은 실제 브라우저에서 깨진다는 게 리뷰에서 밝혀졌다
   * — position:absolute인 <tr>은 계산된 display가 표준상 block으로 바뀌어 그 자손 <td>가 표의 열
   * 그리드(<colgroup>/table-layout:fixed)에서 완전히 떨어져 나가고, 브라우저가 그 <td>들을 별도의
   * 익명 미니 테이블로 감싸 table-layout:auto로 재배치한다 — 그 안의 인라인 width는 "힌트"일 뿐
   * 강제되지 않아, 실제 데이터가 있는 마지막 열이 내용 크기 기준으로 크게 부풀고 그 뒤 열들이
   * 화면 밖으로 밀려났다. 셀마다 인라인 width를 주는 우회책은 근본 원인을 가리지 못했다.)
   *
   * 지금 방식(스페이서 행): 보이는 행(<tr>)은 전부 정상 흐름(position 지정 없음)으로 그려서
   * <colgroup>/table-layout:fixed 열 정렬을 그대로 받는다(틀 고정 행이 이미 하던 방식과 동일).
   * 화면에 없는 위/아래 구간은 실제 <tr>을 만드는 대신, 그 구간 전체 높이(안 그린 행 수 ×
   * ROW_HEIGHT_PX)를 가진 "스페이서" <tr> 하나씩으로 대신한다(buildSpacerRow). 그러면 표의 자연스러운
   * 문서 흐름만으로 전체 스크롤 높이(틀 고정 행 + 위 스페이서 + 보이는 행 + 아래 스페이서 =
   * 전체 시트 높이)가 정확히 나오고, 보이는 행들은 정상 흐름 안에서 실제로 그렸을 때와 똑같은
   * 자리에 놓인다 — 수동으로 top을 계산해 줄 필요가 없다. */
  function renderTableVirtualized(): void {
    renderGeneration += 1;
    // 다시 그리면 이전 <td>들은 전부 버려지므로, 그 순간 실제로 "편집 중"이던 칸도 이제 없다 —
    // blur가 안전하게 안 오는 경우에도(과거 stale-blur 버그) 여기서 확실히 리셋해 둔다.
    editingKey = null;
    // tableEl.replaceChildren()으로 표 내용을 통째로 비우면 그 순간 스크롤 가능한 실제 높이가
    // (다시 채우기 전까지) 뷰포트 높이 수준으로 줄어들어, 브라우저가 scrollHostEl.scrollTop을
    // 그 시점의 최댓값(사실상 0)으로 즉시 clamp해 버린다 — 실측 확인됨(Playwright로
    // replaceChildren() 전/후 scrollTop을 직접 비교: 67500 → 0, 표를 되돌려도 0에 머묾). 표를
    // 비우기 "전에" 지금 스크롤 위치를 미리 읽어 둬야, 아래에서 이 값으로 가상 스크롤 범위를
    // 계산할 때 방금 clamp된 값이 아니라 실제 스크롤 위치를 쓰게 된다 — 그렇지 않으면 스크롤이
    // 일어날 때마다(이 함수가 스크롤 리스너에서 다시 불릴 때마다) 화면이 맨 위로 튕겨 돌아가
    // 사실상 스크롤이 전혀 안 되는 것처럼 보인다.
    const scrollTopSnapshot = scrollHostEl.scrollTop;
    const viewportHeightSnapshot = scrollHostEl.clientHeight;
    const sheet = structure.get(String(activeSheet));
    // 전체 재빌드가 시작되면 지금까지의 증분 패치 캐시(vw*)는 곧 통째로 버려질 DOM을 가리키게
    // 되므로 즉시 무효화한다 — 이 함수가 아래서 일찍 return하더라도(sheet 없음) updateScrollWindow가
    // 낡은 참조로 패치를 시도하지 않고 안전하게 다시 전체 재빌드로 물러난다.
    vwValid = false;
    tableEl.replaceChildren();
    if (!sheet) return;

    warningEl.hidden = !(sheet.rowCount > EDIT_SOFT_ROW_LIMIT || sheet.colCount > EDIT_SOFT_COL_LIMIT);
    const { rows: renderRowCount, cols: renderColCount } = renderCountsFor(sheet);
    const { anchors, covered } = buildMergeLookup(activeSheet);
    // freezeRows는 항상 정상 흐름으로 그려지므로, 실제 렌더 행 수를 넘지 않게 안전하게 자른다.
    const freezeRows = Math.min(sheet.freezeRows ?? 0, renderRowCount);
    const freezeCols = sheet.freezeCols ?? 0;
    // 아래 buildRow/buildSpacerRow(hoisted)는 이 문맥을 매개변수가 아니라 vw* 캐시로 읽는다 — 이
    // 재빌드가 끝난 뒤 updateScrollWindow()가 스크롤만으로 인한 재호출에서 같은 캐시를 그대로
    // 재사용해 전체 재빌드를 건너뛸 수 있게 하기 위해서다.
    vwSheetIdx = activeSheet;
    vwRenderColCount = renderColCount;
    vwRenderRowCount = renderRowCount;
    vwFreezeRows = freezeRows;
    vwFreezeCols = freezeCols;
    vwAnchors = anchors;
    vwCovered = covered;

    // table-layout:fixed에서 열 너비를 정확히 고정하려면 <colgroup>이 가장 확실하다(그 전까진
    // 첫 번째 행 셀 너비에 의존해 왔음 — 실제 엑셀 열 너비 정보가 없던 시절의 임시 방편이었다).
    // 행번호 열은 폭을 지정하지 않아 기존 CSS(.g-edit-rownum min-width)가 그대로 정한다 — 실제
    // 엑셀 열이 아니라 에디터가 그리는 것이라 colWidths에 대응하는 값이 없다.
    const colgroup = document.createElement('colgroup');
    colgroup.appendChild(document.createElement('col'));
    let totalColWidthPx = ROWNUM_COL_WIDTH_PX;
    for (let c = 0; c < renderColCount; c += 1) {
      const width = sheet.colWidths?.[c] ?? DEFAULT_COL_WIDTH_PX;
      const col = document.createElement('col');
      col.style.width = `${width}px`;
      colgroup.appendChild(col);
      totalColWidthPx += width;
    }
    tableEl.appendChild(colgroup);
    // Chromium은 table-layout:fixed + <colgroup>이어도, 테이블 자체 width가 auto인 채로 열이
    // 많으면(대략 12~15개 이상) 개별 <col> 폭을 조용히 무시하고 자체적으로 재분배해버린다(실측
    // 확인됨 — 열이 적을 때는 정상 동작). 테이블에 명시적 width를 주면(정확한 합계가 아니어도,
    // "auto가 아니다"라는 사실만으로) 이 경로를 피하고 각 열이 지정한 폭 그대로 그려진다.
    tableEl.style.width = `${totalColWidthPx}px`;

    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    headRow.style.position = 'relative';
    headRow.style.zIndex = '3';
    const cornerTh = document.createElement('th');
    // table-layout:fixed는 "첫 번째 행"의 셀 너비로 각 열 너비를 확정한다 — 행번호 열의 실제 좁은
    // 너비(.g-edit-rownum)를 적용하려면 헤더 행의 이 모서리 칸도 같은 클래스를 받아야 한다. 이걸
    // 빼먹으면 이 칸이 일반 칸 min-width(100px)를 그대로 물려받아 행번호 열 전체가 넓어져 버린다.
    cornerTh.className = 'g-edit-rownum';
    headRow.appendChild(cornerTh);
    for (let c = 0; c < renderColCount; c += 1) {
      const th = document.createElement('th');
      th.textContent = colLetter(c);
      th.dataset.col = String(c);
      th.addEventListener('click', () => selectColumn(c));
      if (canEdit) {
        th.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          openColumnContextMenu(c, event.clientX, event.clientY);
        });
      }
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    tableEl.appendChild(thead);

    // ---------- 가상 스크롤 범위 계산 ----------
    // (함수 맨 위에서 표를 비우기 전에 미리 읽어 둔 스냅샷을 쓴다 — 지금 다시 읽으면 방금
    // replaceChildren()이 clamp한 값(0)이 나온다.)
    const scrollTop = scrollTopSnapshot;
    const viewportHeight = viewportHeightSnapshot;
    const firstVisible = Math.max(freezeRows, Math.floor(scrollTop / ROW_HEIGHT_PX) - OVERSCAN_ROWS);
    const lastVisible = Math.min(
      renderRowCount - 1,
      Math.ceil((scrollTop + viewportHeight) / ROW_HEIGHT_PX) + OVERSCAN_ROWS,
    );

    const tbody = document.createElement('tbody');

    // 틀 고정된 위쪽 행은 항상 정상 흐름으로(스크롤 위치와 무관하게 계속 보이도록) 그린다. 그
    // 아래, 화면 밖(위쪽)에 있는 안 그리는 구간은 스페이서 행 하나로 대신하고, 실제로 보이는
    // 범위(+여유분)만 진짜 행으로 그린 뒤, 화면 밖(아래쪽) 나머지 구간도 스페이서 행 하나로
    // 마무리한다. firstVisible은 이미 freezeRows 이상으로 보정돼 있다. buildRow가 만든 각 <tr>은
    // vwRowEls에도 기록해 둔다 — 이후 스크롤만으로 인한 재호출(updateScrollWindow)이 이 행들을
    // 재사용할 수 있게 하기 위해서다.
    const newRowEls = new Map<number, HTMLTableRowElement>();
    for (let r = 0; r < freezeRows; r += 1) {
      const tr = buildRow(r);
      newRowEls.set(r, tr);
      tbody.appendChild(tr);
    }
    const topSpacer = buildSpacerRow(firstVisible - freezeRows);
    if (topSpacer) tbody.appendChild(topSpacer);
    for (let r = firstVisible; r <= lastVisible; r += 1) {
      const tr = buildRow(r);
      newRowEls.set(r, tr);
      tbody.appendChild(tr);
    }
    const bottomSpacer = buildSpacerRow(renderRowCount - 1 - lastVisible);
    if (bottomSpacer) tbody.appendChild(bottomSpacer);
    tableEl.appendChild(tbody);
    // 표 내용이 다시 채워져 스크롤 가능한 실제 높이가 복구된 지금(스페이서 행 포함), 함수 맨 위의
    // clamp(스크롤 위치가 0으로 튕겨나간 것)를 되돌린다 — 이 시점 이후로는 (스페이서 덕분에) 그
    // 위치가 실제로 유효한 범위 안이라 clamp되지 않고 그대로 반영된다. 이미 같은 값이면 실제
    // 스크롤 이벤트가 다시 발생하지 않도록(불필요한 재렌더 방지) 값이 다를 때만 대입한다.
    if (scrollHostEl.scrollTop !== scrollTopSnapshot) scrollHostEl.scrollTop = scrollTopSnapshot;
    // 이 전체 재빌드가 방금 그린 것을 스크롤 전용 증분 패치(updateScrollWindow)가 이어받을 수
    // 있도록 캐시를 채운다 — vwValid는 여기서 성공적으로 다 그린 뒤에만 true가 된다.
    vwTbody = tbody;
    vwTopSpacer = topSpacer;
    vwBottomSpacer = bottomSpacer;
    vwRowEls = newRowEls;
    vwFirstVisible = firstVisible;
    vwLastVisible = lastVisible;
    vwValid = true;
    applyFreezePanes(freezeRows, freezeCols);
    renderRemoteCellHighlights();
    // document.hasFocus()가 false면 실제 포커스는 이 문서 밖에 있다 — 이 뷰어는 호스트 페이지에
    // iframe으로 박혀 돌아가는 것을 정식으로 지원하므로(DOCVIEW_ALLOWED_ORIGINS), 사용자가 호스트
    // 쪽 입력칸에 타이핑하는 동안에도 원격 변경·창 크기 변경·확대율 변경 같은 이유로 이 전체
    // 재빌드가 돌 수 있다. 그때 focus()를 부르면 그리드가 호스트의 포커스를 소리 없이 빼앗아,
    // 이어지는 타자가 엉뚱한 셀에 들어가 그대로 커밋된다(restoreScrollWindowFocus의 같은 가드와
    // 짝이다 — 거기 주석에 이유를 더 자세히 적어 뒀다).
    if (activeCellKey && document.hasFocus()) {
      const parsed = parseGridCellKey(activeCellKey);
      // 실제 DOM 포커스를 같은 칸의 새 <td>로 되돌린다 — focus 이벤트가 setActiveHeaders/채우기
      // 핸들 위치 재계산까지 같이 해 준다. 재렌더로 그 칸이 더 이상 없으면(시트 전환, 또는 가상
      // 스크롤 범위 밖으로 밀려남 등) 그냥 둔다.
      if (parsed && parsed.sheetIdx === activeSheet) tdFor(parsed.row, parsed.col)?.focus({ preventScroll: true });
    }
    updateHScrollbar();
    document
      .getElementById('btn-freeze-panes')
      ?.setAttribute('aria-pressed', String(freezeRows > 0 || freezeCols > 0));
    // 지금 화면에 필요한 페이지(서버가 Yjs에 채워 둬야 하는 500행 단위 구간)를 서버에 알린다 —
    // 시트 전환·구조 변경·스크롤로 인한 재렌더 등 renderTableVirtualized가 불리는 모든 경로를
    // 여기 한 곳에서 커버한다(각 호출부에 일일이 추가하는 것보다 누락 위험이 적다).
    scheduleSubscriptionUpdate();
  }

  /**
   * 스크롤"만"으로 인한 재렌더 전용 — renderTableVirtualized()의 가벼운 대체판. <colgroup>/
   * <thead>는 스크롤로 바뀌지 않으므로 아예 건드리지 않고, tbody의 "행 창"도 직전에 그렸던 범위와
   * 겹치는 부분은 기존 <tr>(과 거기 달린 이벤트 리스너)을 그대로 두고, 창 밖으로 나간 행만 지우고
   * 새로 들어온 행만 만들어 끼워 넣는다. 이번 문맥이 직전 전체 재빌드 때와 정확히 같다는 확신이
   * 없으면(시트가 바뀌었다거나, 열/행 수가 달라졌다거나, tbody가 더 이상 표 안에 없다거나) 안전하게
   * renderTableVirtualized()의 전체 재빌드로 물러난다 — 그 판정 자체는 값 비교뿐이라 비용이 거의
   * 없으므로 "혹시 몰라서" 지나치게 자주 물러나도 손해는 아니다.
   */
  function updateScrollWindow(): void {
    if (!vwValid || !vwTbody || vwTbody.parentNode !== tableEl || vwSheetIdx !== activeSheet) {
      renderTableVirtualized();
      return;
    }
    const sheet = structure.get(String(activeSheet));
    if (!sheet) {
      renderTableVirtualized();
      return;
    }
    const { rows: renderRowCount, cols: renderColCount } = renderCountsFor(sheet);
    // 열/행 렌더 수가 달라졌다는 건 구조·확대율·창 크기 중 뭔가 바뀌었다는 뜻 — 그런 변경은 전부
    // 자기 경로에서 이미 renderTableVirtualized()를 직접 부르므로(structure.observe,
    // scheduleViewportRefit 등) 정상적으로는 여기서 다를 일이 없지만, 방어적으로 확인한다.
    if (renderRowCount !== vwRenderRowCount || renderColCount !== vwRenderColCount) {
      renderTableVirtualized();
      return;
    }

    const freezeRows = vwFreezeRows;
    const scrollTop = scrollHostEl.scrollTop;
    const viewportHeight = scrollHostEl.clientHeight;
    const firstVisible = Math.max(freezeRows, Math.floor(scrollTop / ROW_HEIGHT_PX) - OVERSCAN_ROWS);
    const lastVisible = Math.min(
      renderRowCount - 1,
      Math.ceil((scrollTop + viewportHeight) / ROW_HEIGHT_PX) + OVERSCAN_ROWS,
    );
    if (firstVisible === vwFirstVisible && lastVisible === vwLastVisible) return; // 창이 그대로면 할 일 없음

    const tbody = vwTbody;
    // 스페이서는 리스너 없는 <tr>/<td> 한 쌍이라 재사용을 위한 위치 계산보다 새로 만드는 쪽이 싸다
    // — 매번 지우고 끝에서 다시 만든다(실제로 비싼 건 아래에서 그대로 두는 데이터 행이다).
    vwTopSpacer?.remove();
    vwBottomSpacer?.remove();

    // 창 밖으로 나간 행 제거(틀 고정 행은 스크롤과 무관하게 항상 남는다) — Map 이터레이션 도중
    // 지금 방문 중이거나 이미 지나간 항목을 delete하는 것은 안전하다(사양대로 동작).
    for (const [r, tr] of vwRowEls) {
      if (r < freezeRows) continue;
      if (r < firstVisible || r > lastVisible) {
        tr.remove();
        vwRowEls.delete(r);
      }
    }

    // 남아있는 행은 건드리지 않고, 새로 창에 들어온 행만 만들어 순서를 지켜 끼워 넣는다 — 앵커를
    // 오름차순으로 이어가며 "직전에 처리한(기존 또는 방금 만든) 행 바로 뒤"에 놓는다.
    let anchor: ChildNode | null = freezeRows > 0 ? vwRowEls.get(freezeRows - 1) ?? null : null;
    for (let r = firstVisible; r <= lastVisible; r += 1) {
      const existing = vwRowEls.get(r);
      if (existing) {
        anchor = existing;
        continue;
      }
      const tr = buildRow(r);
      if (anchor) anchor.after(tr);
      else tbody.insertBefore(tr, tbody.firstChild);
      vwRowEls.set(r, tr);
      anchor = tr;
    }

    const topSpacer = buildSpacerRow(firstVisible - freezeRows);
    if (topSpacer) {
      const firstWindowRow = vwRowEls.get(firstVisible) ?? null;
      if (firstWindowRow) tbody.insertBefore(topSpacer, firstWindowRow);
      else tbody.appendChild(topSpacer); // 방어적 처리 — 이론상 창이 비어 있을 때만 여기로 온다
    }
    const bottomSpacer = buildSpacerRow(renderRowCount - 1 - lastVisible);
    if (bottomSpacer) tbody.appendChild(bottomSpacer); // 항상 마지막 자식이 맞다(그 뒤엔 아무것도 없다)
    vwTopSpacer = topSpacer;
    vwBottomSpacer = bottomSpacer;
    vwFirstVisible = firstVisible;
    vwLastVisible = lastVisible;

    // 틀 고정 "열"은 스크롤 창 밖에 있던 새 행의 셀에도 sticky left를 걸어야 한다(틀 고정 행은
    // 창 밖으로 나가지 않으니 다시 계산할 필요가 없다 — applyFreezePanes는 두 축을 한 번에
    // 처리하지만 freezeRows>0 branch는 이미 있는 고정 행만 다시 훑어 idempotent하다). 틀 고정을
    // 아예 안 쓰는 보통 시트에서는 이 함수가 즉시 리턴해 비용이 없다.
    if (freezeRows > 0 || vwFreezeCols > 0) applyFreezePanes(freezeRows, vwFreezeCols);
    renderRemoteCellHighlights();
    // 전체 재빌드(renderTableVirtualized)가 맨 끝에서 하던 "활성 칸 포커스 되돌리기"와, 그 포커스
    // 복원이 focusin을 통해 곁들여 해 주던 "선택 표시 다시 칠하기"를 이 경로에도 똑같이 붙인다 —
    // 이게 없어서 (1) 스크롤 왕복 뒤 키 입력이 <body>로 새어 조용히 사라지고, (2) 화면에는 아무
    // 선택도 안 보이는데 툴바 서식은 예전 범위 전체에 먹는 문제가 있었다. 순서가 중요하다:
    // 포커스 복원이 먼저고(그 안에서 focusin이 하이라이트를 한 번 지운다), 칠하기가 나중이다.
    restoreScrollWindowFocus(firstVisible, lastVisible, freezeRows);
    repaintScrollWindowSelection();
  }

  // ---------- 셀 선택 vs 편집모드, 방향키 이동 (엑셀과 동일하게) ----------
  // 칸을 클릭하면(focus) "선택"만 된다 — contenteditable은 여전히 false다. 실제로 값을 고칠 수
  // 있으려면(1) F2, (2) 더블클릭, (3) 선택된 채로 문자를 바로 타이핑, 셋 중 하나로 편집모드에
  // 들어가야 한다. 편집모드에서 Enter/Tab은 커밋 후 아래/오른쪽 칸으로 넘어가고, Escape는 원래
  // 값으로 되돌리고 선택 상태로만 남는다.

  function placeCaretAtEnd(el: HTMLElement): void {
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
  }

  /** 지금 이 칸을 다른 참여자가 편집 중이면 그 사람 정보를 돌려준다(없으면 null) — 있으면 이 칸은
   * 편집모드로 못 들어간다(같은 칸을 동시에 고치는 충돌을 막는다. 서버가 강제하는 진짜 락은 아니고
   * awareness 브로드캐스트 기반의 느슨한 락이라, 정확히 동시에 누르는 극단적인 경우까지 완벽히
   * 막지는 못하지만 — 실사용에서 충분하다). */
  function remoteEditorOf(row: number, col: number): { name: string; color: string } | null {
    const states = provider.awareness.getStates() as Map<number, RemoteCellState>;
    for (const [clientId, state] of states) {
      if (clientId === ydoc.clientID) continue;
      const cell = state.cell;
      if (!cell || cell.sheetIdx !== activeSheet || cell.row !== row || cell.col !== col) continue;
      const name = state.user?.name;
      if (!name) continue;
      return { name, color: state.user?.color ?? pickPresenceColorForName(name) };
    }
    return null;
  }

  function enterEditMode(td: HTMLTableCellElement, key: string, replaceWith?: string): void {
    if (!td.classList.contains('g-editable')) return; // 잠긴 칸/읽기전용 문서는 편집모드 자체가 없다
    const parsed = parseGridCellKey(key);
    if (parsed) {
      const other = remoteEditorOf(parsed.row, parsed.col);
      if (other) {
        showCellTooltip(td, `${other.name}님이 편집 중입니다.`);
        return;
      }
    }
    editingKey = key;
    if (parsed) provider.awareness.setLocalStateField('cell', parsed); // "지금 편집 중" — 다른 참여자에게 표시
    td.contentEditable = 'true';
    td.textContent = replaceWith !== undefined ? replaceWith : cells.get(key) ?? '';
    td.focus();
    placeCaretAtEnd(td);
    formulaInputEl.value = td.textContent ?? '';
    // replaceWith가 주어졌다는 건 기존 값을 지우고 새로 치는 중(엑셀의 "입력")이라는 뜻 —
    // undefined면 더블클릭 등으로 기존 내용을 그대로 두고 이어 쓰는 "편집"이다.
    setGridStatusText(replaceWith !== undefined ? '입력' : '편집');
  }

  function commitAndMove(td: HTMLTableCellElement, key: string, row: number, col: number, dRow: number, dCol: number): void {
    commitCell(td);
    editingKey = null;
    provider.awareness.setLocalStateField('cell', null);
    td.contentEditable = 'false';
    renderCellDisplay(td, key);
    recalcFormulaCells();
    moveSelection(row, col, dRow, dCol);
    setGridStatusText('준비');
    flushPendingStructureRender();
  }

  function cancelEdit(td: HTMLTableCellElement, key: string): void {
    editingKey = null;
    provider.awareness.setLocalStateField('cell', null);
    td.contentEditable = 'false';
    renderCellDisplay(td, key); // 원래 값(수정 전)으로 되돌린다
    td.focus();
    setGridStatusText('준비');
    flushPendingStructureRender();
  }

  /** td.scrollIntoView는 sticky thead/행번호 열이 실제로 화면을 가리고 있다는 걸 모른다 — 그
   * 칸의 bounding box가 스크롤 컨테이너 사각형과 조금이라도 겹치면 "이미 보인다"고 판단해 스크롤을
   * 안 해 버려서, 위로 점프한 칸이 고정 헤더 뒤에 숨어버리는 경우가 있었다(아래로 점프할 땐 그
   * 자리를 가리는 고정 요소가 없어서 문제가 안 드러났다). 고정 헤더/행번호 열의 실제 두께만큼을
   * "보이는 영역"에서 미리 빼고 직접 스크롤한다. */
  function scrollCellIntoView(td: HTMLElement): void {
    const hostRect = scrollHostEl.getBoundingClientRect();
    const cellRect = td.getBoundingClientRect();
    const theadHeight = tableEl.querySelector('thead')?.getBoundingClientRect().height ?? 0;
    const rownumWidth = tableEl.querySelector<HTMLElement>('.g-edit-rownum')?.getBoundingClientRect().width ?? 0;
    const visibleTop = hostRect.top + theadHeight;
    const visibleLeft = hostRect.left + rownumWidth;
    if (cellRect.top < visibleTop) scrollHostEl.scrollTop -= visibleTop - cellRect.top;
    else if (cellRect.bottom > hostRect.bottom) scrollHostEl.scrollTop += cellRect.bottom - hostRect.bottom;
    if (cellRect.left < visibleLeft) scrollHostEl.scrollLeft -= visibleLeft - cellRect.left;
    else if (cellRect.right > hostRect.right) scrollHostEl.scrollLeft += cellRect.right - hostRect.right;
  }

  /** 지금 칸에서 (dRow,dCol)만큼 떨어진 칸으로 선택을 옮긴다. 병합에 덮여 실제 <td>가 없는
   * 칸이면(예: 1행 제목 병합) 같은 방향으로 한 칸씩 더 건너뛴다. 시트 렌더 범위 끝에 닿으면 멈춘다. */
  function moveSelection(fromRow: number, fromCol: number, dRow: number, dCol: number): void {
    const sheet = currentStructure();
    const { rows: rowCount, cols: colCount } = renderCountsFor(sheet);
    let row = fromRow;
    let col = fromCol;
    for (let attempt = 0; attempt < 30; attempt += 1) {
      const nextRow = Math.min(Math.max(row + dRow, 0), rowCount - 1);
      const nextCol = Math.min(Math.max(col + dCol, 0), colCount - 1);
      if (nextRow === row && nextCol === col) return; // 끝에 닿아 더 못 감
      row = nextRow;
      col = nextCol;
      const targetTd = tdForEnsureRendered(row, col);
      if (targetTd) {
        targetTd.focus();
        scrollCellIntoView(targetTd);
        return;
      }
    }
  }

  /** 엑셀의 Ctrl+방향키(End 모드) — 빈 칸이거나 바로 다음 칸이 비어 있으면 다음 데이터(또는
   * 시트 끝)까지 건너뛰고, 데이터가 연속돼 있으면 그 블록이 끝나는 칸까지 단번에 이동한다. */
  function moveSelectionToDataEdge(fromRow: number, fromCol: number, dRow: number, dCol: number): void {
    const sheet = currentStructure();
    const { rows: rowCount, cols: colCount } = renderCountsFor(sheet);
    const inBounds = (r: number, c: number) => r >= 0 && c >= 0 && r < rowCount && c < colCount;
    const isEmpty = (r: number, c: number): boolean => {
      if (!inBounds(r, c)) return true;
      const v = cells.get(gridCellKey(activeSheet, r, c));
      return v === undefined || v === '';
    };
    let row = fromRow;
    let col = fromCol;
    if (isEmpty(row, col) || isEmpty(row + dRow, col + dCol)) {
      let r = row + dRow;
      let c = col + dCol;
      while (inBounds(r, c) && isEmpty(r, c)) {
        r += dRow;
        c += dCol;
      }
      row = Math.min(Math.max(r, 0), rowCount - 1);
      col = Math.min(Math.max(c, 0), colCount - 1);
    } else {
      let r = row;
      let c = col;
      while (inBounds(r + dRow, c + dCol) && !isEmpty(r + dRow, c + dCol)) {
        r += dRow;
        c += dCol;
      }
      row = r;
      col = c;
    }
    // 병합에 덮인 칸이라 실제 <td>가 없으면(moveSelection과 동일한 보정) 같은 방향으로 한 칸 더 간다.
    for (let attempt = 0; attempt < 30; attempt += 1) {
      const targetTd = tdForEnsureRendered(row, col);
      if (targetTd) {
        targetTd.focus();
        scrollCellIntoView(targetTd);
        return;
      }
      const nextRow = Math.min(Math.max(row + dRow, 0), rowCount - 1);
      const nextCol = Math.min(Math.max(col + dCol, 0), colCount - 1);
      if (nextRow === row && nextCol === col) return;
      row = nextRow;
      col = nextCol;
    }
  }

  function handleCellKeydown(
    event: KeyboardEvent,
    td: HTMLTableCellElement,
    row: number,
    col: number,
    key: string,
    canEditThisCell: boolean,
  ): void {
    if (editingKey === key) {
      if (event.key === 'Enter') {
        event.preventDefault();
        commitAndMove(td, key, row, col, 1, 0);
      } else if (event.key === 'Tab') {
        event.preventDefault();
        commitAndMove(td, key, row, col, 0, event.shiftKey ? -1 : 1);
      } else if (event.key === 'Escape') {
        event.preventDefault();
        cancelEdit(td, key);
      }
      return; // 다른 키(화살표 등)는 텍스트 커서 이동 등 브라우저 기본 동작 그대로 둔다
    }

    switch (event.key) {
      case 'F2':
        if (canEditThisCell) {
          event.preventDefault();
          enterEditMode(td, key);
        }
        break;
      case 'Enter':
        event.preventDefault();
        moveSelection(row, col, 1, 0);
        break;
      case 'Tab':
        event.preventDefault();
        moveSelection(row, col, 0, event.shiftKey ? -1 : 1);
        break;
      case 'ArrowUp':
        event.preventDefault();
        if (event.ctrlKey || event.metaKey) moveSelectionToDataEdge(row, col, -1, 0);
        else moveSelection(row, col, -1, 0);
        break;
      case 'ArrowDown':
        event.preventDefault();
        if (event.ctrlKey || event.metaKey) moveSelectionToDataEdge(row, col, 1, 0);
        else moveSelection(row, col, 1, 0);
        break;
      case 'ArrowLeft':
        event.preventDefault();
        if (event.ctrlKey || event.metaKey) moveSelectionToDataEdge(row, col, 0, -1);
        else moveSelection(row, col, 0, -1);
        break;
      case 'ArrowRight':
        event.preventDefault();
        if (event.ctrlKey || event.metaKey) moveSelectionToDataEdge(row, col, 0, 1);
        else moveSelection(row, col, 0, 1);
        break;
      case 'Delete':
      case 'Backspace': {
        if (!canEdit) break;
        event.preventDefault();
        const { r0, r1, c0, c1 } = rangeBounds();
        ydoc.transact(() => {
          for (let r = r0; r <= r1; r += 1) {
            for (let c = c0; c <= c1; c += 1) {
              const k = gridCellKey(activeSheet, r, c);
              if (styles.get(k)?.lk === 1) continue; // 잠긴 칸은 건너뜀(선택 범위에 섞여 있어도 됨)
              if ((cells.get(k) ?? '') !== '') cells.delete(k);
            }
          }
        });
        break;
      }
      default:
        // Ctrl/Cmd+C — 지금 선택된 범위를 TSV로 클립보드에 복사한다.
        if ((event.ctrlKey || event.metaKey) && (event.key === 'c' || event.key === 'C')) {
          event.preventDefault();
          copySelectionToClipboard();
          break;
        }
        // 선택된 채로 문자를 바로 치면(엑셀처럼) 기존 값을 지우고 그 문자로 편집모드에 들어간다.
        if (canEditThisCell && event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
          event.preventDefault();
          enterEditMode(td, key, event.key);
        }
    }
  }

  // ---------- 자동 채우기 핸들 (드래그로 값 복사, 엑셀과 동일) ----------
  const FILL_HANDLE_SIZE = 7;
  const fillHandleEl = document.createElement('div');
  fillHandleEl.className = 'g-fill-handle';
  fillHandleEl.hidden = true;
  scrollHostEl.appendChild(fillHandleEl);

  /** 지금 선택된 칸(td)의 오른쪽 아래 모서리에 채우기 핸들을 옮겨 놓는다. 편집 불가한 칸이 선택되면
   * 숨긴다. scrollHostEl의 스크롤 콘텐츠 안에 핸들이 같이 들어있으므로 스크롤해도 같이 움직인다. */
  function positionFillHandle(td: HTMLTableCellElement | null): void {
    if (!td) {
      fillHandleEl.hidden = true;
      return;
    }
    const tdRect = td.getBoundingClientRect();
    const hostRect = scrollHostEl.getBoundingClientRect();
    fillHandleEl.style.left = `${tdRect.right - hostRect.left + scrollHostEl.scrollLeft - FILL_HANDLE_SIZE / 2}px`;
    fillHandleEl.style.top = `${tdRect.bottom - hostRect.top + scrollHostEl.scrollTop - FILL_HANDLE_SIZE / 2}px`;
    fillHandleEl.hidden = false;
  }

  let fillSource: { row: number; col: number; key: string } | null = null;
  let fillTargetRow = 0;
  let fillTargetCol = 0;

  /** 드래그로 지나간 칸들(시작 칸 제외)에 한 번씩 콜백을 호출한다 — 세로/가로 중 더 크게 움직인
   * 한 방향으로만 채워진다(엑셀 자동 채우기와 동일하게 대각선으로는 안 채워짐). */
  function forEachFillCell(fn: (row: number, col: number) => void): void {
    if (!fillSource) return;
    const r0 = Math.min(fillSource.row, fillTargetRow);
    const r1 = Math.max(fillSource.row, fillTargetRow);
    const c0 = Math.min(fillSource.col, fillTargetCol);
    const c1 = Math.max(fillSource.col, fillTargetCol);
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        if (r === fillSource.row && c === fillSource.col) continue;
        fn(r, c);
      }
    }
  }

  function clearFillPreview(): void {
    tableEl.querySelectorAll('.g-fill-preview').forEach((el) => el.classList.remove('g-fill-preview'));
  }

  function onFillDragMove(event: MouseEvent): void {
    if (!fillSource) return;
    const el = document.elementFromPoint(event.clientX, event.clientY) as HTMLElement | null;
    const td = el?.closest('td[data-row][data-col]') as HTMLTableCellElement | null;
    if (!td) return;
    const row = Number(td.dataset.row);
    const col = Number(td.dataset.col);
    const dRow = row - fillSource.row;
    const dCol = col - fillSource.col;
    if (Math.abs(dRow) >= Math.abs(dCol)) {
      fillTargetRow = row;
      fillTargetCol = fillSource.col;
    } else {
      fillTargetRow = fillSource.row;
      fillTargetCol = col;
    }
    clearFillPreview();
    forEachFillCell((r, c) => tdFor(r, c)?.classList.add('g-fill-preview'));
  }

  function onFillDragEnd(): void {
    document.removeEventListener('mousemove', onFillDragMove);
    document.removeEventListener('mouseup', onFillDragEnd);
    clearFillPreview();
    // forEachFillCell이 module 범위의 fillSource를 그대로 읽으므로, 실제 작업이 끝나기 전까지는
    // null로 되돌리면 안 된다(먼저 비우면 forEachFillCell이 바로 빈 채로 리턴해버려 아무 칸도
    // 못 채운다) — 다 쓰고 나서 맨 마지막에만 정리한다.
    if (!fillSource || !canEdit) {
      fillSource = null;
      return;
    }
    const sourceRow = fillSource.row;
    const sourceCol = fillSource.col;
    const sourceValue = cells.get(fillSource.key) ?? '';
    const touched: Array<{ r: number; c: number }> = [];
    ydoc.transact(() => {
      forEachFillCell((r, c) => {
        const key = gridCellKey(activeSheet, r, c);
        const targetStyle = styles.get(key);
        if (targetStyle?.lk === 1) return; // 잠긴 칸은 건너뜀
        if (targetStyle?.nu === 1 && sourceValue !== '' && !NUMERIC_VALUE.test(sourceValue)) return; // 숫자 전용 칸에 숫자가 아니면 건너뜀
        if (sourceValue === '') cells.delete(key);
        else cells.set(key, sourceValue);
        touched.push({ r, c });
      });
      const s = currentStructure();
      const maxRow = Math.max(sourceRow, fillTargetRow);
      const maxCol = Math.max(sourceCol, fillTargetCol);
      if (maxRow >= s.rowCount || maxCol >= s.colCount) {
        structure.set(String(activeSheet), {
          ...s,
          rowCount: Math.max(s.rowCount, maxRow + 1),
          colCount: Math.max(s.colCount, maxCol + 1),
        });
      }
    });
    touched.forEach(({ r, c }) => {
      const td = tdFor(r, c);
      if (td) renderCellDisplay(td, gridCellKey(activeSheet, r, c));
    });
    recalcFormulaCells();
    fillSource = null;
  }

  fillHandleEl.addEventListener('mousedown', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!activeCellKey) return;
    const parsed = parseGridCellKey(activeCellKey);
    if (!parsed) return;
    fillSource = { row: parsed.row, col: parsed.col, key: activeCellKey };
    fillTargetRow = parsed.row;
    fillTargetCol = parsed.col;
    document.addEventListener('mousemove', onFillDragMove);
    document.addEventListener('mouseup', onFillDragEnd);
  });

  /**
   * 위 freezeRows행 / 왼쪽 freezeCols열을 실제로 sticky 처리한다. 렌더링 직후 DOM에 실제로 붙은
   * 뒤에 호출해야 offsetHeight/offsetWidth로 진짜 셀 크기를 잴 수 있다 — 그래야 컬럼 letter 헤더
   * (이미 sticky top:0)와 행번호 열(이미 sticky left:0) 바로 다음에 정확히 이어붙일 수 있다.
   * 배경을 불투명하게 칠해야 스크롤된 내용이 그 밑으로 비쳐 보이지 않는다.
   */
  function applyFreezePanes(freezeRows: number, freezeCols: number): void {
    if (freezeRows === 0 && freezeCols === 0) return;

    const theadHeight = tableEl.querySelector('thead')?.getBoundingClientRect().height ?? 0;
    const rownumWidth = tableEl.querySelector<HTMLElement>('.g-edit-rownum')?.getBoundingClientRect().width ?? 0;

    if (freezeRows > 0) {
      // 행 높이는 전부 동일(CSS 고정값)하므로, 얼릴 행들의 실제 렌더 높이를 각 행 대표 셀에서 재
      // 누적해 top을 계산한다.
      let top = theadHeight;
      for (let r = 0; r < freezeRows; r += 1) {
        const rowCells = tableEl.querySelectorAll<HTMLElement>(`tbody td[data-row="${r}"], tbody th[data-row="${r}"]`);
        let rowHeight = 0;
        rowCells.forEach((cell) => {
          cell.style.position = 'sticky';
          cell.style.top = `${top}px`;
          cell.style.zIndex = cell.classList.contains('g-frozen-col') ? '4' : '3';
          cell.classList.add('g-frozen-pane');
          rowHeight = Math.max(rowHeight, cell.getBoundingClientRect().height);
        });
        top += rowHeight;
      }
    }

    if (freezeCols > 0) {
      let left = rownumWidth;
      for (let c = 0; c < freezeCols; c += 1) {
        const colCells = tableEl.querySelectorAll<HTMLElement>(`tbody td[data-col="${c}"]`);
        let colWidth = 0;
        colCells.forEach((cell) => {
          cell.style.position = 'sticky';
          cell.style.left = `${left}px`;
          if (!cell.style.zIndex) cell.style.zIndex = '3';
          cell.classList.add('g-frozen-pane');
          colWidth = Math.max(colWidth, cell.getBoundingClientRect().width);
        });
        left += colWidth;
      }
    }
  }

  function clearActiveHeaders(): void {
    tableEl.querySelectorAll('.g-col-active').forEach((el) => el.classList.remove('g-col-active'));
    tableEl.querySelectorAll('.g-row-active').forEach((el) => el.classList.remove('g-row-active'));
  }
  function setActiveHeaders(row: number, col: number): void {
    clearActiveHeaders();
    clearHeaderSelection();
    tableEl.querySelector(`thead th[data-col="${col}"]`)?.classList.add('g-col-active');
    tableEl.querySelector(`tbody th[data-row="${row}"]`)?.classList.add('g-row-active');
  }

  // 열/행 머리글(A, 1 등)을 직접 클릭했을 때 — 엑셀처럼 그 열/행 전체를 선택 표시한다.
  /** 지금 머리글 클릭으로 통째로 선택돼 있는 열/행 — 클래스(.g-col-selected/.g-row-selected)는
   * 해당 <tr>이 가상 스크롤 창 밖으로 나가면 그 행과 함께 사라져 버려서, 화면만으로는 "원래
   * 무엇이 선택돼 있었는지"를 복구할 수 없다. 스크롤로 행이 다시 들어왔을 때 표시를 되살리려면
   * 클래스와 별개로 이 상태를 들고 있어야 한다. */
  let headerSelection: { kind: 'col' | 'row'; index: number } | null = null;
  function clearHeaderSelection(): void {
    headerSelection = null;
    tableEl.querySelectorAll('.g-col-selected, .g-row-selected').forEach((el) => {
      el.classList.remove('g-col-selected', 'g-row-selected');
    });
  }
  /** headerSelection이 가리키는 열/행에 선택 표시 클래스를 입힌다 — 머리글 클릭 직후와, 스크롤로
   * 그 행이 다시 그려진 뒤 양쪽에서 같은 코드를 쓴다. */
  function paintHeaderSelection(): void {
    if (!headerSelection) return;
    const { kind, index } = headerSelection;
    const attr = kind === 'col' ? 'data-col' : 'data-row';
    const cls = kind === 'col' ? 'g-col-selected' : 'g-row-selected';
    tableEl.querySelectorAll(`[${attr}="${index}"]`).forEach((el) => el.classList.add(cls));
  }
  // 열/행 머리글을 클릭하면 시각적 표시만 하는 게 아니라, 서식 툴바가 실제로 적용할 범위
  // (selAnchor~selFocus)도 그 열/행 전체로 잡아야 한다 — 안 그러면 rangeBounds()가 여전히
  // 마지막으로 포커스했던 칸 하나만 가리켜, 열/행을 선택한 채 서식을 바꿔도 그 칸에만 적용되고
  // 나머지는 그대로인 것처럼 보였다(실제로는 "적용이 안 되는" 게 아니라 "한 칸에만 적용되는" 버그).
  function selectColumn(col: number): void {
    clearActiveHeaders();
    clearHeaderSelection();
    const { rowCount } = currentStructure();
    selAnchorRow = 0;
    selFocusRow = Math.max(0, rowCount - 1);
    selAnchorCol = col;
    selFocusCol = col;
    activeCellKey = gridCellKey(activeSheet, selAnchorRow, col);
    renderRangeHighlight();
    headerSelection = { kind: 'col', index: col };
    paintHeaderSelection();
  }
  function selectRow(row: number): void {
    clearActiveHeaders();
    clearHeaderSelection();
    const { colCount } = currentStructure();
    selAnchorRow = row;
    selFocusRow = row;
    selAnchorCol = 0;
    selFocusCol = Math.max(0, colCount - 1);
    activeCellKey = gridCellKey(activeSheet, row, selAnchorCol);
    renderRangeHighlight();
    headerSelection = { kind: 'row', index: row };
    paintHeaderSelection();
  }
  // 잠긴 칸은 focus 리스너 자체가 안 붙으므로(편집 불가) setActiveHeaders를 못 거친다 — 어떤 칸을
  // 클릭하든(잠금 여부와 무관하게) 열/행 전체 선택 표시는 항상 풀리도록 여기서 한 번 더 잡는다.
  tableEl.addEventListener('click', (event) => {
    if ((event.target as HTMLElement).tagName === 'TD') clearHeaderSelection();
  });

  // ---------- 다중 셀 드래그 선택 (엑셀처럼 클릭+드래그로 범위 선택) ----------
  function rangeBounds(): { r0: number; r1: number; c0: number; c1: number } {
    return {
      r0: Math.min(selAnchorRow, selFocusRow),
      r1: Math.max(selAnchorRow, selFocusRow),
      c0: Math.min(selAnchorCol, selFocusCol),
      c1: Math.max(selAnchorCol, selFocusCol),
    };
  }
  function clearRangeHighlight(): void {
    tableEl.querySelectorAll('.g-range-selected').forEach((el) => el.classList.remove('g-range-selected'));
  }
  /**
   * 시트를 바꿀 때 — 선택 상태(드래그 범위/머리글 선택/활성 칸)를 통째로 비운다.
   *
   * selAnchor~selFocus도 headerSelection도 "몇 번 시트의" 선택인지를 들고 있지 않다(activeCellKey만
   * 시트 번호를 갖는다). 그래서 시트를 바꿔도 이 값들이 그대로 남아 있으면 그대로 새 시트에 적용된다:
   *  (1) 스크롤 한 번이면 repaintScrollWindowSelection()이 이전 시트의 범위를 새 시트에 칠한다
   *      — 실측: Sheet1에서 C열 전체를 선택하고 Sheet2로 간 뒤 휠을 한 번 굴리니, 선택한 적도
   *      없는 Sheet2에 125칸이 선택된 것처럼 칠해졌다.
   *  (2) 더 나쁘게는, 툴바 서식 버튼이 전부 rangeBounds() + gridCellKey(activeSheet, r, c)로 키를
   *      만들기 때문에 — 칠하기 버그가 없더라도 이전 시트에서 잡아 둔 범위가 새 시트의 같은 좌표
   *      칸들에 조용히 먹는다(같은 실측에서 굵게가 Sheet2의 125칸에 그대로 먹었다).
   * 엑셀에서도 시트를 바꾸면 보고 있지 않은 시트의 선택이 따라오지 않으므로, 새 시트는 아무것도
   * 선택되지 않은 상태로 시작한다 — 사용자가 칸을 클릭하는 순간 focusin이 전부 다시 채운다.
   */
  function resetSelectionForSheetSwitch(): void {
    clearRangeHighlight();
    clearActiveHeaders();
    clearHeaderSelection();
    activeCellKey = null;
    selAnchorRow = 0;
    selAnchorCol = 0;
    selFocusRow = 0;
    selFocusCol = 0;
    // 활성 칸이 없어졌으니 그 칸 오른쪽 아래에 떠 있던 채우기 핸들도 같이 감춘다 — 안 그러면 새
    // 시트 위에 그대로 남고, 그걸 끌면 (activeCellKey가 null이라) 아무 일도 안 하는 죽은 손잡이가 된다.
    positionFillHandle(null);
  }
  /** 앵커~포커스 사각형에 배경을 칠한다 — 1칸짜리(드래그 안 한 보통 클릭)면 기존 :focus 테두리로
   * 충분하니 별도 배경은 안 칠한다. */
  function renderRangeHighlight(): void {
    clearRangeHighlight();
    const { r0, r1, c0, c1 } = rangeBounds();
    if (r0 === r1 && c0 === c1) return;
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        tdFor(r, c)?.classList.add('g-range-selected');
      }
    }
  }
  /**
   * updateScrollWindow() 전용 — 스크롤로 새로 들어온 행에 선택 관련 "칠하기"를 다시 입힌다.
   *
   * 왜 필요한가: buildRow()는 늘 아무 장식도 없는 <td>를 만든다. 전체 재빌드 경로에서는 마지막에
   * 활성 칸을 다시 focus시켜 focusin 핸들러가 renderRangeHighlight/setActiveHeaders를 도로
   * 불러줬지만, 스크롤 전용 증분 패치에는 그 마무리가 통째로 빠져 있었다. 그 결과 선택 범위를
   * 잡아 둔 채 스크롤로 한 바퀴 돌고 오면 **화면에는 아무것도 선택돼 있지 않은데 내부 선택 상태
   * (selAnchor~selFocus)는 그대로 살아 있는** 상태가 됐다 — 이때 굵게/서식 지우기/채우기색 같은
   * 툴바 버튼(전부 rangeBounds()를 읽는다)을 누르면 "보이지 않는" 범위 수십~수백 칸에 서식이
   * 조용히 적용된다(실측: 화면상 0칸 선택 상태에서 굵게가 44칸에 먹었다).
   *
   * 성능: renderRangeHighlight()는 범위의 칸마다 tableEl.querySelector를 한 번씩 돌아서,
   * selectColumn()이 잡는 "열 전체"(selFocusRow = rowCount-1, 수천 행)를 선택해 둔 채 스크롤하면
   * 스크롤 프레임마다 수천 번의 DOM 질의가 된다 — 이 함수 전체가 매 스크롤 프레임 도는 자리라
   * 그건 81be4d8이 없앤 바로 그 종류의 비용을 되살리는 셈이다. 어차피 창 밖 행에는 칠할 <td>가
   * 존재하지 않으므로, 지금 DOM에 있는 행(vwRowEls)과 범위의 교집합만 훑는다.
   */
  function repaintScrollWindowSelection(): void {
    const { r0, r1, c0, c1 } = rangeBounds();
    // 한 칸짜리 선택이면 칠할 것이 없다(renderRangeHighlight와 같은 규칙 — :focus 테두리로 충분).
    // 이때는 지울 것도 없다: 선택이 한 칸으로 줄어드는 모든 경로가 renderRangeHighlight를 거쳐
    // 이미 전부 지웠고, 새로 만들어진 행에는 애초에 클래스가 없다. 그래서 선택 범위를 안 쓰는
    // 보통의 스크롤에서는 이 함수가 DOM을 단 한 번도 건드리지 않는다.
    if (r0 !== r1 || c0 !== c1) {
      clearRangeHighlight();
      for (const [r, tr] of vwRowEls) {
        if (r < r0 || r > r1) continue;
        // 행 하나당 자식 훑기 한 번 — 칸마다 querySelector를 도는 것보다 훨씬 싸다.
        for (let i = 0; i < tr.children.length; i += 1) {
          const cell = tr.children[i] as HTMLElement;
          const col = cell.dataset.col;
          if (col === undefined) continue; // 행번호 <th>
          const c = Number(col);
          if (c >= c0 && c <= c1) cell.classList.add('g-range-selected');
        }
      }
    }

    // 머리글 클릭으로 열/행을 통째로 선택해 둔 상태(g-col-selected/g-row-selected)도 새로 들어온
    // 행에 다시 입힌다 — 선택을 안 쓰는 보통의 스크롤에서는 headerSelection이 null이라 비용이 없다.
    paintHeaderSelection();

    // 활성 칸이 있는 행의 행번호 머리글 표시(g-row-active)도 그 <tr>과 함께 사라졌다가 돌아온다.
    if (activeCellKey) {
      const parsed = parseGridCellKey(activeCellKey);
      if (parsed && parsed.sheetIdx === activeSheet) {
        tableEl.querySelector(`tbody th[data-row="${parsed.row}"]`)?.classList.add('g-row-active');
      }
    }
  }

  /**
   * updateScrollWindow() 전용 — 창 밖으로 나갔다 돌아온 활성 칸에 실제 DOM 포커스를 되돌린다.
   * 전체 재빌드 경로(renderTableVirtualized)는 맨 끝에서 이미 같은 일을 하지만 스크롤 전용 증분
   * 패치에는 그게 없었다: 활성 칸의 <tr>이 창 밖으로 제거되는 순간 포커스가 <body>로 떨어지는데,
   * 키보드 핸들러는 tableEl에 위임돼 있어 <body>에 포커스가 있는 상태로 친 키는 그리드에 **아예
   * 도달하지 못한다** — 다시 스크롤해 그 칸이 돌아와도 포커스는 <body>에 남아, 사용자가 친 글자가
   * 아무런 피드백 없이 통째로 사라졌다.
   *
   * 포커스가 정말로 "아무 데도 없을 때"(<body>)만 되돌린다 — 수식 입력줄이나 툴바 버튼, 열린
   * 대화상자에 사용자가 옮겨 둔 포커스를 스크롤이 가로채면 안 된다.
   */
  function restoreScrollWindowFocus(firstVisible: number, lastVisible: number, freezeRows: number): void {
    if (!activeCellKey) return;
    // activeElement만으로는 "이 문서 안에서 아무 데도 포커스가 없다"와 "포커스가 아예 이 문서
    // 밖(호스트 페이지·다른 탭·다른 창)에 있다"를 구분할 수 없다 — 둘 다 activeElement가 <body>다.
    // 이 뷰어는 호스트 페이지에 iframe으로 박혀 돌아가는 것을 정식으로 지원하므로
    // (DOCVIEW_ALLOWED_ORIGINS, buildRow의 window.parent.postMessage 참고), 사용자가 iframe을
    // 클릭한 적조차 없이 호스트 쪽 입력칸에 타이핑하면서 그리드 위에 휠만 굴려도 이 함수가 돌 수
    // 있다. 그때 포커스를 되돌리면 그리드가 호스트의 포커스를 소리 없이 빼앗아, 이어지는 타자가
    // 호스트 입력칸이 아니라 스프레드시트 칸에 들어가 그대로 커밋된다 — "키가 사라지는" 원래
    // 버그보다 훨씬 나쁘다. document.hasFocus()는 정확히 이 경우에만 false고, 정작 고치려던
    // 경우(이 문서는 포커스를 갖고 있는데 활성 칸의 <tr>이 스크롤 패치로 잠깐 사라져 <body>로
    // 떨어진 상태)에는 true다.
    if (!document.hasFocus()) return;
    const active = document.activeElement;
    if (active !== null && active !== document.body) return;
    const parsed = parseGridCellKey(activeCellKey);
    if (!parsed || parsed.sheetIdx !== activeSheet) return;
    const { row, col } = parsed;
    // 틀 고정 행은 창과 무관하게 항상 그려져 있다. 그 밖의 행은 지금 창 안에 들어왔을 때만.
    if (row >= freezeRows && (row < firstVisible || row > lastVisible)) return;
    const td = tdFor(row, col);
    if (!td) return;
    // focus()는 위임된 focusin 핸들러를 그 자리에서(동기적으로) 실행하는데, 그 핸들러는 "칸을 새로
    // 클릭했다"는 뜻으로 선택 범위를 그 한 칸으로 접고(selAnchor/selFocus) 머리글 선택도 풀어버린다
    // (setActiveHeaders → clearHeaderSelection). 하지만 스크롤은 선택을 바꾸는 동작이 아니므로
    // (엑셀에서도 스크롤로 선택이 풀리지 않는다) 원래 선택 상태를 스냅숏해 뒀다가 그대로 되돌린다
    // — 실제 칠하기는 호출부가 이어서 repaintScrollWindowSelection()으로 한다.
    const anchorRow = selAnchorRow;
    const anchorCol = selAnchorCol;
    const focusRow = selFocusRow;
    const focusCol = selFocusCol;
    const header = headerSelection;
    td.focus({ preventScroll: true });
    selAnchorRow = anchorRow;
    selAnchorCol = anchorCol;
    selFocusRow = focusRow;
    selFocusCol = focusCol;
    headerSelection = header;
  }

  function onRangeDragMove(event: MouseEvent): void {
    const el = document.elementFromPoint(event.clientX, event.clientY) as HTMLElement | null;
    const td = el?.closest('td[data-row][data-col]') as HTMLTableCellElement | null;
    if (!td) return;
    const row = Number(td.dataset.row);
    const col = Number(td.dataset.col);
    if (row === selFocusRow && col === selFocusCol) return;
    selFocusRow = row;
    selFocusCol = col;
    renderRangeHighlight();
  }
  function onRangeDragEnd(): void {
    isRangeDragging = false;
    document.removeEventListener('mousemove', onRangeDragMove);
    document.removeEventListener('mouseup', onRangeDragEnd);
  }

  /** 지금 선택된 범위를 엑셀과 같은 TSV(탭=열, 줄바꿈=행)로 클립보드에 담는다 — handlePaste가
   * 읽는 형식과 대칭이라 그리드끼리, 또는 실제 엑셀과 오가는 복사/붙여넣기가 그대로 맞는다. */
  function copySelectionToClipboard(): void {
    const { r0, r1, c0, c1 } = rangeBounds();
    const lines: string[] = [];
    for (let r = r0; r <= r1; r += 1) {
      const rowValues: string[] = [];
      for (let c = c0; c <= c1; c += 1) {
        rowValues.push(cells.get(gridCellKey(activeSheet, r, c)) ?? '');
      }
      lines.push(rowValues.join('\t'));
    }
    navigator.clipboard?.writeText(lines.join('\n')).catch(() => {
      /* 클립보드 권한이 없는 환경이면 조용히 무시 — 복사 실패를 사용자에게 굳이 방해되게 알릴 정도는 아니다 */
    });
  }

  // 셀 편집 영역을 완전히 벗어날 때만(다른 셀로 이동하는 경우는 제외) awareness의 cell을 지운다 —
  // focusout은 버블링되므로 표 전체에 한 번만 걸어 둔다.
  // activeCellKey는 일부러 지우지 않는다 — 툴바 버튼(굵게 등)을 누르면 버튼이 포커스를 가져가며
  // 셀이 blur되는데, 이때 activeCellKey까지 지우면 "마지막으로 선택했던 셀에 서식 적용"이라는
  // 스프레드시트의 기본 동작이 깨진다(클릭한 서식 버튼이 조용히 아무 일도 안 하게 됨).
  // awareness의 cell(= "지금 편집 중" 표시)만 셀 영역을 완전히 벗어났을 때 지운다.
  tableEl.addEventListener('focusout', (event) => {
    const related = event.relatedTarget as HTMLElement | null;
    if (!related || related.getAttribute('contenteditable') !== 'true') {
      provider.awareness.setLocalStateField('cell', null);
    }
  });

  function tdFor(row: number, col: number): HTMLTableCellElement | null {
    return tableEl.querySelector<HTMLTableCellElement>(`td[data-row="${row}"][data-col="${col}"]`);
  }

  /** 가상 스크롤 때문에 지금 화면 창(뷰포트+여유분) 밖의 행은 <td>가 DOM에 없을 수 있다 — 키보드로
   * 먼 칸으로 바로 이동하거나(Ctrl+방향키) 검색 결과로 점프할 때처럼, 화면 밖 칸으로 곧장 이동해야
   * 하면 그 행이 보이도록 스크롤 위치를 옮기고 즉시 다시 그린 뒤 다시 찾는다. 병합에 덮인 칸처럼
   * 애초에 <td>가 존재할 수 없는 경우엔 다시 그려도 여전히 null이며, 호출부의 기존 재시도 루프가
   * 그 경우를 처리한다. */
  function tdForEnsureRendered(row: number, col: number): HTMLTableCellElement | null {
    const existing = tdFor(row, col);
    if (existing) return existing;
    const viewportHeight = scrollHostEl.clientHeight;
    scrollHostEl.scrollTop = Math.max(0, row * ROW_HEIGHT_PX - viewportHeight / 2);
    renderTableVirtualized();
    return tdFor(row, col);
  }

  interface RemoteCellState extends AwarenessUserState {
    readonly cell?: { readonly sheetIdx: number; readonly row: number; readonly col: number };
  }

  /** 다른 참여자가 지금 편집 중인 셀에 색 테두리 + 작은 이름표를 띄운다("누가 어디를 수정 중인지"). */
  /** 다른 참여자가 지금 편집 중인 셀에 그 사람 색으로 테두리만 두른다(이름표는 따로 안 보여준다 —
   * 색만으로 "누군가 지금 여기 있다"는 걸 충분히 알 수 있으면 됨). */
  function renderRemoteCellHighlights(): void {
    tableEl.querySelectorAll<HTMLTableCellElement>('.g-remote-cell').forEach((td) => {
      td.classList.remove('g-remote-cell');
      td.style.removeProperty('--remote-color');
    });
    const states = provider.awareness.getStates() as Map<number, RemoteCellState>;
    for (const [clientId, state] of states) {
      if (clientId === ydoc.clientID) continue;
      const cell = state.cell;
      const name = state.user?.name;
      if (!cell || !name || cell.sheetIdx !== activeSheet) continue;
      const td = tdFor(cell.row, cell.col);
      if (!td) continue;
      const color = state.user?.color ?? pickPresenceColorForName(name);
      td.classList.add('g-remote-cell');
      td.style.setProperty('--remote-color', color);
    }
  }

  cells.observe((event) => {
    for (const key of event.keysChanged) {
      const parsed = parseGridCellKey(key);
      if (!parsed || parsed.sheetIdx !== activeSheet) continue;
      const td = tdFor(parsed.row, parsed.col);
      // 선택만 된 칸(편집 중 아님)은 이제 원격 갱신을 그대로 반영해도 된다 — 실제로 로컬에서
      // 타이핑 중인 칸(editingKey === key)만 덮지 않는다. (선택=focus일 뿐 편집과는 분리됐다.)
      if (!td || editingKey === key) continue;
      const value = cells.get(key) ?? '';
      renderCellDisplay(td, key);
      applyStyleToTd(td, styles.get(key));
      if (key === activeCellKey && document.activeElement !== formulaInputEl) formulaInputEl.value = value;
      // 이 칸이 비었다/채워졌다가 바뀌면, 왼쪽에서 이 칸으로 넘쳐 들어오고 있었을 스필 폭도
      // 다시 계산해야 한다(엑셀처럼 — 옆 칸이 채워지면 넘침이 줄고, 비면 다시 늘어난다).
      refreshSpillNeighbor(key);
    }
    recalcFormulaCells(); // 바뀐 셀을 참조하는 다른 수식 셀들도 다시 계산
  });

  styles.observe((event) => {
    for (const key of event.keysChanged) {
      const parsed = parseGridCellKey(key);
      if (!parsed || parsed.sheetIdx !== activeSheet) continue;
      const td = tdFor(parsed.row, parsed.col);
      if (!td || editingKey === key) continue;
      applyStyleToTd(td, styles.get(key));
    }
  });

  /**
   * 서버가 계산해 둔 수식 값(GRID_COMPUTED_MAP)이 도착하면 그 칸만 다시 그린다 — 리뷰 지적 C3.
   *
   * 클라이언트는 더 이상 수식을 직접 계산하지 않고 computed 맵의 값을 읽어 보여줄 뿐인데
   * (computeCellDisplay), 이 맵에는 관찰자가 하나도 없었다. cells.observe가 부르는
   * recalcFormulaCells()는 **셀 값이 바뀐 그 순간** 동기적으로 도는 것이라, 그때는 서버가 아직
   * 계산하기 전이다(계산은 왕복 한 번 뒤에 computed 맵으로 돌아온다). 그래서 수식이 참조하는 칸을
   * 고쳐도 수식 칸의 표시값이 그대로 멈춰 있고, 스크롤·클릭 등 **다른 이유로 재렌더가 일어나야만**
   * 뒤늦게 갱신됐다 — 최종 리뷰가 잡아낸 바로 그 증상이다.
   *
   * 편집 중인 칸(editingKey === key)은 건드리지 않는다: 그 칸에는 계산 결과가 아니라 사용자가
   * 지금 치고 있는 수식 원문이 보여야 하고, 재렌더는 커서/선택 범위를 잃게 한다(cells/styles
   * 관찰자가 쓰는 것과 같은 가드).
   */
  computed.observe((event) => {
    for (const key of event.keysChanged) {
      const parsed = parseGridCellKey(key);
      if (!parsed || parsed.sheetIdx !== activeSheet) continue;
      const td = tdFor(parsed.row, parsed.col);
      if (td && editingKey !== key) renderCellDisplay(td, key);
    }
  });

  /**
   * 편집이 끝난 뒤, 편집 중이라 미뤄 뒀던 구조 변경 재렌더를 따라잡는다 — 리뷰 지적 I5의 뒷부분.
   * 편집 종료 경로(blur 커밋 / commitAndMove / cancelEdit) 끝에서 부른다.
   */
  function flushPendingStructureRender(): void {
    if (!pendingStructureRender || editingKey !== null) return;
    pendingStructureRender = false;
    renderSheetTabs();
    renderTableVirtualized();
  }

  structure.observe((event) => {
    /**
     * 원격에서 온 행/열 삽입·삭제라면 undo 스택을 통째로 버린다 — 리뷰 지적 I-4.
     *
     * UndoManager는 이 구조 변경 자체는 추적하지 않는다(서버가 보낸 원격 트랜잭션이라 trackedOrigins
     * 밖이다 — 위 sendGridStructureOp 주석의 "알려진 대가"). 문제는 **스택에 이미 쌓여 있는 로컬
     * 셀 편집들**이다: 그것들은 편집 당시의 `sheet:row:col` 키로 되돌리는데, 행/열이 밀린 뒤에는 그
     * 키가 전혀 다른 행/열을 가리킨다. 그 상태로 Ctrl+Z를 누르면 사용자가 손댄 적 없는 칸에 옛 값이
     * 조용히 써진다 — 이 프로젝트가 계속 쫓아온 그 손상 클래스와 같은 모양이다.
     *
     * 스택 항목을 새 위치로 다시 매핑하는 것이 이상적이지만 훨씬 복잡하다. 되돌릴 수 없게 되는
     * 것보다 잘못된 칸을 덮어쓰는 것이 훨씬 나쁘므로, 표준적인 완화책대로 스택을 버린다.
     *
     * 판정은 두 조건의 곱이다: (1) 로컬 트랜잭션이 아니다(`transaction.local === false` — Yjs가
     * 원격 업데이트를 적용할 때 쓰는 값), (2) 기존 시트의 행/열 수가 실제로 바뀌었다. 그래서 시트
     * 이름 변경·탭 순서 변경 같은 원격 갱신이나, 아무 칸도 밀지 않는 로컬 행 추가 버튼
     * (btn-add-row 등)은 undo를 잃지 않는다. 내가 누른 삽입·삭제도 서버를 돌아 **원격 트랜잭션으로**
     * 도착하므로 여기서 같이 걸린다(의도한 대로 — 내가 눌렀든 남이 눌렀든 스택 키는 낡는다).
     */
    // 기준값 갱신은 **로컬 이벤트에서도 반드시** 해야 한다(&& 단축 평가로 건너뛰면 로컬 행 추가가
    // 기준값을 낡게 만들어, 이후의 무해한 원격 갱신 하나가 스택을 괜히 날린다).
    const sheetExtentChanged = consumeSheetExtentChange();
    if (!event.transaction.local && sheetExtentChanged) undoManager.clear();
    // 편집 중(contentEditable)에 표를 통째로 다시 그리면 그 <td>가 DOM에서 떨어져 나가며 강제
    // blur가 발생하고, 방금 타이핑하던 값이 커밋되지 못한 채 사라진다(위 blur 핸들러의 세대 판정이
    // 그런 blur를 "버림"으로 처리한다) — 스크롤 리스너가 같은 이유로 쓰는 가드와 같은 판정이다.
    //
    // 다만 스크롤과 달리 구조 변경은 "곧 또 오는" 이벤트가 아니다: 다른 참여자의 행 삽입이나 시트
    // 이름 변경은 한 번뿐이라, 그냥 건너뛰면 편집이 끝나도 그 변경이 화면에 영영 반영되지 않는다
    // (편집 종료 경로는 그 칸 하나만 다시 그린다). 그래서 건너뛰되 반드시 예약해 둔다.
    if (editingKey !== null) {
      pendingStructureRender = true;
      return;
    }
    renderSheetTabs();
    renderTableVirtualized();
  });

  function currentStructure(): SheetStructure {
    return structure.get(String(activeSheet)) ?? { name: 'Sheet1', rowCount: 1, colCount: 1 };
  }

  document.getElementById('btn-add-row')?.addEventListener('click', () => {
    const s = currentStructure();
    structure.set(String(activeSheet), { ...s, rowCount: s.rowCount + 1 });
  });
  document.getElementById('btn-del-row')?.addEventListener('click', () => {
    const s = currentStructure();
    if (s.rowCount <= 1) return;
    structure.set(String(activeSheet), { ...s, rowCount: s.rowCount - 1 });
  });
  document.getElementById('btn-add-col')?.addEventListener('click', () => {
    const s = currentStructure();
    structure.set(String(activeSheet), { ...s, colCount: s.colCount + 1 });
  });
  document.getElementById('btn-del-col')?.addEventListener('click', () => {
    const s = currentStructure();
    if (s.colCount <= 1) return;
    structure.set(String(activeSheet), { ...s, colCount: s.colCount - 1 });
  });

  // ---------- 자동 합계 — 엑셀처럼 위쪽(없으면 왼쪽)으로 연속된 숫자 칸을 찾아 SUM 수식을 채운다 ----------
  document.getElementById('btn-autosum')?.addEventListener('click', () => {
    if (!activeCellKey || !canEdit) return;
    const parsed = parseGridCellKey(activeCellKey);
    if (!parsed) return;
    const { row, col } = parsed;
    const isNumericCell = (r: number, c: number): boolean => {
      const v = cells.get(gridCellKey(activeSheet, r, c));
      return v !== undefined && v !== '' && (NUMERIC_VALUE.test(v.trim()) || v.startsWith('='));
    };
    let formula = '';
    let r = row - 1;
    while (r >= 0 && isNumericCell(r, col)) r -= 1;
    if (r < row - 1) {
      formula = `=SUM(${cellRefLabel(r + 1, col)}:${cellRefLabel(row - 1, col)})`;
    } else {
      let c = col - 1;
      while (c >= 0 && isNumericCell(row, c)) c -= 1;
      if (c < col - 1) formula = `=SUM(${cellRefLabel(row, c + 1)}:${cellRefLabel(row, col - 1)})`;
    }
    if (!formula) return;
    const td = tdFor(row, col);
    if (!td || !td.classList.contains('g-editable')) return;
    enterEditMode(td, activeCellKey, formula);
    commitAndMove(td, activeCellKey, row, col, 0, 0);
  });

  // ---------- 틀 고정 — 선택한 칸의 위쪽 행 전체 + 왼쪽 열 전체를 고정한다(다시 누르면 해제) ----------
  document.getElementById('btn-freeze-panes')?.addEventListener('click', () => {
    if (!activeCellKey) return;
    const parsed = parseGridCellKey(activeCellKey);
    if (!parsed) return;
    const s = currentStructure();
    const isFrozen = (s.freezeRows ?? 0) > 0 || (s.freezeCols ?? 0) > 0;
    structure.set(
      String(activeSheet),
      isFrozen ? { ...s, freezeRows: 0, freezeCols: 0 } : { ...s, freezeRows: parsed.row, freezeCols: parsed.col },
    );
  });

  // ---------- 서식 도구 (선택 범위 전체에 적용, 오피스처럼 눌린 상태로 반영) ----------
  function setPressed(id: string, pressed: boolean): void {
    document.getElementById(id)?.setAttribute('aria-pressed', String(pressed));
  }
  const DEFAULT_FONT_SIZE_PT = 11;
  function syncCellToolbarState(key: string | null): void {
    const style = key ? styles.get(key) : undefined;
    setPressed('btn-cell-bold', style?.b === 1);
    setPressed('btn-cell-italic', style?.i === 1);
    setPressed('btn-cell-underline', style?.u === 1);
    setPressed('btn-cell-strike', style?.st === 1);
    setPressed('btn-cell-align-l', !style?.a || style.a === 'l');
    setPressed('btn-cell-align-c', style?.a === 'c');
    setPressed('btn-cell-align-r', style?.a === 'r');
    setPressed('btn-valign-top', style?.va === 't');
    setPressed('btn-valign-middle', style?.va === 'm');
    setPressed('btn-valign-bottom', !style?.va || style.va === 'b');
    setPressed('btn-wrap-text', style?.wrap === 1);
    const fontFamilySelect = document.getElementById('cell-font-family') as HTMLSelectElement | null;
    if (fontFamilySelect) fontFamilySelect.value = style?.ff ?? '';
    const fontSizeInput = document.getElementById('cell-font-size') as HTMLInputElement | null;
    if (fontSizeInput) {
      const pt = style?.fs ? Math.round(style.fs * PX_TO_PT) : DEFAULT_FONT_SIZE_PT;
      fontSizeInput.value = String(pt);
    }
    const numfmtSelect = document.getElementById('cell-numfmt-select') as HTMLSelectElement | null;
    if (numfmtSelect) numfmtSelect.value = style?.nf ?? 'general';
  }

  /** 지금 선택 범위(다중 드래그 선택이면 그 사각형 전체, 아니면 칸 하나)에 서식을 적용한다 —
   * 셀마다 "지금 그 칸의 서식"을 기준으로 mutate가 따로 호출되므로, 굵게 상태가 섞인 범위를
   * 선택해도 엑셀처럼 칸별로 자연스럽게 토글된다. */
  function updateStyle(mutate: (current: CellStyleAttrs) => CellStyleAttrs): void {
    if (!activeCellKey) return;
    const { r0, r1, c0, c1 } = rangeBounds();
    ydoc.transact(() => {
      for (let r = r0; r <= r1; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          const key = gridCellKey(activeSheet, r, c);
          const current = styles.get(key) ?? {};
          const next = mutate(current);
          if (Object.keys(next).length === 0) styles.delete(key);
          else styles.set(key, next);
          const td = tdFor(r, c);
          if (td) {
            applyStyleToTd(td, styles.get(key));
            renderCellDisplay(td, key); // 숫자 서식은 텍스트 자체가 바뀌므로 다시 그린다
          }
        }
      }
    });
    // 툴바 버튼을 누르면 버튼이 포커스를 가져가 셀은 이미 blur된 상태이므로 :focus로는 못 찾는다
    // — activeCellKey(범위의 기준 칸)를 좌표로 풀어 직접 찾아 눌림 상태를 갱신한다.
    syncCellToolbarState(activeCellKey);
  }
  function toggleFlag(flag: 'b' | 'i' | 'u' | 'st' | 'wrap'): void {
    updateStyle((current) => {
      const next = { ...current };
      if (next[flag]) delete next[flag];
      else next[flag] = 1;
      return next;
    });
  }
  document.getElementById('btn-cell-bold')?.addEventListener('click', () => toggleFlag('b'));
  document.getElementById('btn-cell-italic')?.addEventListener('click', () => toggleFlag('i'));
  document.getElementById('btn-cell-underline')?.addEventListener('click', () => toggleFlag('u'));
  document.getElementById('btn-cell-strike')?.addEventListener('click', () => toggleFlag('st'));
  document.getElementById('btn-wrap-text')?.addEventListener('click', () => toggleFlag('wrap'));
  document.getElementById('btn-cell-clear-format')?.addEventListener('click', () => {
    if (!activeCellKey) return;
    const { r0, r1, c0, c1 } = rangeBounds();
    ydoc.transact(() => {
      for (let r = r0; r <= r1; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          const key = gridCellKey(activeSheet, r, c);
          styles.delete(key);
          const td = tdFor(r, c);
          if (td) {
            applyStyleToTd(td, undefined);
            renderCellDisplay(td, key);
          }
        }
      }
    });
    syncCellToolbarState(activeCellKey);
  });

  // 글자 색/배경 색 — 아이콘을 누르면 막대에 표시된 마지막 색이 바로 적용되고, 캐럿을
  // 눌러야 네이티브 색상 선택기가 뜬다(엑셀과 동일하게 "적용 버튼"과 "색 고르기"를 분리).
  const cellColorInput = document.getElementById('cell-font-color') as HTMLInputElement | null;
  const cellColorBtn = document.getElementById('cell-font-color-btn') as HTMLElement | null;
  cellColorBtn?.style.setProperty('--btn-color', cellColorInput?.value ?? '#c0392b');
  cellColorInput?.addEventListener('input', () => {
    cellColorBtn?.style.setProperty('--btn-color', cellColorInput.value);
    updateStyle((current) => ({ ...current, col: cellColorInput.value }));
  });
  document.getElementById('btn-font-color-apply')?.addEventListener('click', () => {
    if (!cellColorInput) return;
    updateStyle((current) => ({ ...current, col: cellColorInput.value }));
  });
  const cellBgInput = document.getElementById('cell-bg-color') as HTMLInputElement | null;
  const cellBgBtn = document.getElementById('cell-bg-color-btn') as HTMLElement | null;
  cellBgBtn?.style.setProperty('--btn-color', cellBgInput?.value ?? '#ffff00');
  cellBgInput?.addEventListener('input', () => {
    cellBgBtn?.style.setProperty('--btn-color', cellBgInput.value);
    updateStyle((current) => ({ ...current, bg: cellBgInput.value }));
  });
  document.getElementById('btn-bg-color-apply')?.addEventListener('click', () => {
    if (!cellBgInput) return;
    updateStyle((current) => ({ ...current, bg: cellBgInput.value }));
  });

  (['l', 'c', 'r'] as const).forEach((align) => {
    document.getElementById(`btn-cell-align-${align}`)?.addEventListener('click', () => {
      updateStyle((current) => ({ ...current, a: align }));
    });
  });
  (['t', 'm', 'b'] as const).forEach((va) => {
    const id = va === 't' ? 'btn-valign-top' : va === 'm' ? 'btn-valign-middle' : 'btn-valign-bottom';
    document.getElementById(id)?.addEventListener('click', () => {
      updateStyle((current) => ({ ...current, va }));
    });
  });

  // ---------- 글꼴 이름/크기 ----------
  const fontFamilySelect = document.getElementById('cell-font-family') as HTMLSelectElement | null;
  fontFamilySelect?.addEventListener('change', () => {
    updateStyle((current) => {
      const next = { ...current };
      if (fontFamilySelect.value) next.ff = fontFamilySelect.value;
      else delete next.ff;
      return next;
    });
  });
  const fontSizeInput = document.getElementById('cell-font-size') as HTMLInputElement | null;
  function applyFontSizePt(pt: number): void {
    const clamped = Math.max(6, Math.min(96, Math.round(pt)));
    if (fontSizeInput) fontSizeInput.value = String(clamped);
    updateStyle((current) => {
      const next = { ...current };
      if (clamped === DEFAULT_FONT_SIZE_PT) delete next.fs;
      else next.fs = Math.round(clamped * PT_TO_PX);
      return next;
    });
  }
  fontSizeInput?.addEventListener('change', () => {
    const pt = Number(fontSizeInput.value);
    applyFontSizePt(Number.isFinite(pt) && pt > 0 ? pt : DEFAULT_FONT_SIZE_PT);
  });
  document.getElementById('btn-font-size-up')?.addEventListener('click', () => {
    applyFontSizePt(Number(fontSizeInput?.value ?? DEFAULT_FONT_SIZE_PT) + 1);
  });
  document.getElementById('btn-font-size-down')?.addEventListener('click', () => {
    applyFontSizePt(Number(fontSizeInput?.value ?? DEFAULT_FONT_SIZE_PT) - 1);
  });

  // ---------- 표시 형식(숫자 서식) ----------
  const numfmtSelect = document.getElementById('cell-numfmt-select') as HTMLSelectElement | null;
  numfmtSelect?.addEventListener('change', () => {
    updateStyle((current) => {
      const next = { ...current };
      if (numfmtSelect.value === 'general') {
        delete next.nf;
        delete next.nd;
      } else {
        next.nf = numfmtSelect.value as CellStyleAttrs['nf'];
        if (!next.nd) next.nd = next.nf === 'currency' ? 2 : 0;
      }
      return next;
    });
  });
  function applyQuickNumFmt(nf: NonNullable<CellStyleAttrs['nf']>, nd: number): void {
    updateStyle((current) => ({ ...current, nf, nd }));
  }
  document.getElementById('btn-numfmt-currency')?.addEventListener('click', () => applyQuickNumFmt('currency', 2));
  document.getElementById('btn-numfmt-percent')?.addEventListener('click', () => applyQuickNumFmt('percent', 0));
  document.getElementById('btn-numfmt-comma')?.addEventListener('click', () => applyQuickNumFmt('number', 0));
  function stepDecimals(delta: number): void {
    updateStyle((current) => {
      const nextDecimals = Math.max(0, (current.nd ?? 0) + delta);
      return { ...current, nf: current.nf ?? 'number', nd: nextDecimals };
    });
  }
  document.getElementById('btn-numfmt-decimal-up')?.addEventListener('click', () => stepDecimals(1));
  document.getElementById('btn-numfmt-decimal-down')?.addEventListener('click', () => stepDecimals(-1));

  // ---------- 테두리 ----------
  type BorderPreset = 'none' | 'all' | 'outside' | 'outside-thick' | 'top' | 'bottom' | 'left' | 'right';
  function applyBorderPreset(preset: BorderPreset): void {
    if (!activeCellKey) return;
    const { r0, r1, c0, c1 } = rangeBounds();
    ydoc.transact(() => {
      for (let r = r0; r <= r1; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          const key = gridCellKey(activeSheet, r, c);
          const next: { -readonly [K in keyof CellStyleAttrs]?: CellStyleAttrs[K] } = { ...(styles.get(key) ?? {}) };
          const isTop = r === r0;
          const isBottom = r === r1;
          const isLeft = c === c0;
          const isRight = c === c1;
          const weight = preset === 'outside-thick' ? 2 : 1;
          switch (preset) {
            case 'none':
              delete next.bt; delete next.br; delete next.bb; delete next.bl;
              break;
            case 'all':
              next.bt = 1; next.br = 1; next.bb = 1; next.bl = 1;
              break;
            case 'outside':
            case 'outside-thick':
              if (isTop) next.bt = weight; else delete next.bt;
              if (isBottom) next.bb = weight; else delete next.bb;
              if (isLeft) next.bl = weight; else delete next.bl;
              if (isRight) next.br = weight; else delete next.br;
              break;
            case 'top':
              if (isTop) next.bt = 1;
              break;
            case 'bottom':
              if (isBottom) next.bb = 1;
              break;
            case 'left':
              if (isLeft) next.bl = 1;
              break;
            case 'right':
              if (isRight) next.br = 1;
              break;
          }
          if (Object.keys(next).length === 0) styles.delete(key);
          else styles.set(key, next as CellStyleAttrs);
          const td = tdFor(r, c);
          if (td) applyStyleToTd(td, styles.get(key));
        }
      }
    });
  }
  const borderDropdown = document.getElementById('border-dropdown') as HTMLElement | null;
  const borderMenu = document.getElementById('border-menu') as HTMLElement | null;
  function closeBorderMenu(): void {
    if (borderMenu) borderMenu.hidden = true;
  }
  borderDropdown?.querySelector('.dropdown-caret')?.addEventListener('click', (event) => {
    event.stopPropagation();
    if (borderMenu) borderMenu.hidden = !borderMenu.hidden;
  });
  borderMenu?.addEventListener('click', (event) => {
    const target = (event.target as HTMLElement).closest('button[data-border]') as HTMLButtonElement | null;
    if (!target) return;
    applyBorderPreset(target.dataset.border as BorderPreset);
    closeBorderMenu();
  });
  document.addEventListener('click', (event) => {
    if (borderDropdown && !borderDropdown.contains(event.target as Node)) closeBorderMenu();
  });
  document.getElementById('btn-border')?.addEventListener('click', () => applyBorderPreset('all'));

  // ---------- 확대/축소 (엑셀처럼 드래그 슬라이더 포함) ----------
  const zoomLevelEl = document.getElementById('zoom-level') as HTMLElement;
  const zoomSliderTrackEl = document.getElementById('zoom-slider-track') as HTMLElement | null;
  const zoomSliderFillEl = document.getElementById('zoom-slider-fill') as HTMLElement | null;
  const zoomSliderHandleEl = document.getElementById('zoom-slider-handle') as HTMLElement | null;
  const ZOOM_MIN = 0.6;
  const ZOOM_MAX = 2;
  let zoom = 1;
  // 확대율이 낮아질수록 화면에 보이는 실제 영역이 넓어지므로, 그 영역을 채우기에 지금 그려둔
  // 열 수(가상 스크롤 대상이 아님)와 가상 스크롤 범위(뷰포트 높이 기준)가 여전히 맞는지 다시
  // 확인한다(rAF로 묶어서 슬라이더를 드래그하는 동안 매 프레임 전체 재렌더가 도는 걸 막는다).
  // renderTableVirtualized()는 이제 화면에 보이는 범위만 그리므로(시트 전체 크기와 무관하게
  // 항상 가볍다), 예전처럼 "이미 충분히 그려져 있는지" 미리 비교할 필요 없이 그냥 다시 부른다.
  let viewportRefitScheduled = false;
  function scheduleViewportRefit(): void {
    if (viewportRefitScheduled) return;
    viewportRefitScheduled = true;
    requestAnimationFrame(() => {
      viewportRefitScheduled = false;
      renderTableVirtualized();
    });
  }
  function applyZoom(): void {
    zoomHost.style.zoom = String(zoom);
    zoomLevelEl.textContent = `${Math.round(zoom * 100)}%`;
    const ratio = (zoom - ZOOM_MIN) / (ZOOM_MAX - ZOOM_MIN);
    if (zoomSliderFillEl) zoomSliderFillEl.style.width = `${ratio * 100}%`;
    if (zoomSliderHandleEl) zoomSliderHandleEl.style.left = `${ratio * 100}%`;
    scheduleViewportRefit();
  }
  window.addEventListener('resize', scheduleViewportRefit);
  function setZoomFromRatio(ratio: number): void {
    const clamped = Math.max(0, Math.min(1, ratio));
    zoom = Math.round((ZOOM_MIN + clamped * (ZOOM_MAX - ZOOM_MIN)) * 20) / 20; // 5%p 단위로 스냅
    applyZoom();
  }
  document.getElementById('zoom-in')?.addEventListener('click', () => {
    zoom = Math.min(ZOOM_MAX, Math.round((zoom + 0.1) * 10) / 10);
    applyZoom();
  });
  document.getElementById('zoom-out')?.addEventListener('click', () => {
    zoom = Math.max(ZOOM_MIN, Math.round((zoom - 0.1) * 10) / 10);
    applyZoom();
  });
  function ratioFromClientX(clientX: number): number {
    const rect = zoomSliderTrackEl!.getBoundingClientRect();
    return (clientX - rect.left) / rect.width;
  }
  zoomSliderTrackEl?.addEventListener('mousedown', (event) => {
    if (event.target === zoomSliderHandleEl) return; // 핸들 자체의 드래그는 아래 리스너가 처리
    setZoomFromRatio(ratioFromClientX(event.clientX));
  });
  zoomSliderHandleEl?.addEventListener('mousedown', (event) => {
    event.preventDefault();
    const onMove = (moveEvent: MouseEvent) => setZoomFromRatio(ratioFromClientX(moveEvent.clientX));
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
  applyZoom();

  // ---------- 엑셀처럼 시트 탭 옆에 붙는 커스텀 가로 스크롤바 ----------
  // 네이티브 가로 스크롤바는 CSS에서 숨기고(그리드 바로 아래에 뜨는 게 엑셀과 달라서), 이 위젯이
  // scrollHostEl.scrollLeft를 그대로 읽고 쓴다 — 장식이 아니라 실제 스크롤을 대신한다.
  const hscrollEl = document.querySelector('.grid-hscroll') as HTMLElement | null;
  const hscrollTrackEl = document.getElementById('hscroll-track') as HTMLElement | null;
  const hscrollThumbEl = document.getElementById('hscroll-thumb') as HTMLElement | null;
  function updateHScrollbar(): void {
    if (!hscrollEl || !hscrollTrackEl || !hscrollThumbEl) return;
    const { scrollWidth, clientWidth, scrollLeft } = scrollHostEl;
    const scrollable = scrollWidth > clientWidth + 1;
    hscrollEl.dataset.scrollable = String(scrollable);
    if (!scrollable) {
      hscrollThumbEl.style.width = '100%';
      hscrollThumbEl.style.left = '0';
      return;
    }
    const trackWidth = hscrollTrackEl.clientWidth;
    const thumbWidth = Math.max(16, (clientWidth / scrollWidth) * trackWidth);
    const maxThumbLeft = trackWidth - thumbWidth;
    const maxScrollLeft = scrollWidth - clientWidth;
    const thumbLeft = maxScrollLeft > 0 ? (scrollLeft / maxScrollLeft) * maxThumbLeft : 0;
    hscrollThumbEl.style.width = `${thumbWidth}px`;
    hscrollThumbEl.style.left = `${thumbLeft}px`;
  }
  scrollHostEl.addEventListener('scroll', updateHScrollbar);
  window.addEventListener('resize', updateHScrollbar);

  // ---------- 가상 스크롤 재렌더 + 페이지 구독 전송(윈도잉) ----------
  // 서버는 이제 클라이언트가 MESSAGE_PAGE_SUBSCRIBE로 알려준 500행 단위 "페이지"만 Yjs에 채운다
  // (room.ts 참고) — 스크롤로 보이는 범위가 바뀔 때마다 그 범위를 포함하는 페이지 번호를 계산해
  // 서버에 알려야, 스크롤해서 새로 보이는 행에도 실제 데이터가 채워진다.
  const MESSAGE_PAGE_SUBSCRIBE = 2;
  let subscriptionDebounceTimer: ReturnType<typeof setTimeout> | null = null;

  /** room.ts가 받는 그대로의 형식(varUint 메시지 타입 + varString JSON 페이로드)으로 같은 WebSocket
   * 연결에 얹어 보낸다. y-websocket의 WebsocketProvider는 이 커스텀 메시지 타입을 몰라 자기
   * messageHandlers로 처리하지 못하지만, 그건 수신 쪽 얘기고(서버는 room.ts에서 messageType으로
   * 직접 분기해 처리한다) — 전송은 provider가 이미 들고 있는 실제 WebSocket 인스턴스
   * (node_modules/y-websocket/src/y-websocket.js 확인: `provider.ws`가 `WebSocket | null`이고,
   * 라이브러리 자신도 내부적으로 `this.ws.send(encoding.toUint8Array(encoder))`로 syncStep1을
   * 보낸다 — 같은 패턴을 그대로 따른다)를 통해 그대로 내보내면 된다. */
  function sendPageSubscription(sheetIdx: number, pages: readonly number[]): void {
    if (!provider.ws || provider.ws.readyState !== WebSocket.OPEN) return;
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
    encoding.writeVarString(encoder, JSON.stringify({ sheet: sheetIdx, pages }));
    provider.ws.send(encoding.toUint8Array(encoder));
  }

  /** 지금 스크롤 위치(+여유분) 기준으로 필요한 페이지 번호를 계산해 서버에 알린다. 스크롤 중
   * 매 프레임 보낼 필요는 없으므로(서버가 그때마다 페이지를 채우고 비우는 건 낭비다) 200ms
   * 디바운스한다 — renderTableVirtualized()가 호출되는 모든 경로(스크롤, 시트 전환, 구조 변경,
   * 최초 동기화)에서 공통으로 이 함수가 불린다. */
  function scheduleSubscriptionUpdate(): void {
    if (subscriptionDebounceTimer) clearTimeout(subscriptionDebounceTimer);
    subscriptionDebounceTimer = setTimeout(() => {
      subscriptionDebounceTimer = null;
      const scrollTop = scrollHostEl.scrollTop;
      const viewportHeight = scrollHostEl.clientHeight;
      const firstRow = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT_PX) - OVERSCAN_ROWS);
      const lastRow = Math.ceil((scrollTop + viewportHeight) / ROW_HEIGHT_PX) + OVERSCAN_ROWS;
      const firstPage = Math.floor(firstRow / PAGE_SIZE_ROWS);
      const lastPage = Math.floor(lastRow / PAGE_SIZE_ROWS);
      const pages = new Set<number>();
      for (let p = firstPage; p <= lastPage; p += 1) pages.add(p);
      // 틀 고정된 위쪽 행은 스크롤 위치와 무관하게 항상 화면에 그려지므로(renderTableVirtualized),
      // 그 데이터가 담긴 페이지(freezeRows는 항상 시트 맨 위부터라 언제나 0번 페이지)도 항상 같이
      // 구독해 둔다 — 안 그러면 멀리 스크롤했을 때 고정된 헤더 행이 빈 칸으로 보일 수 있다.
      const freezeRows = structure.get(String(activeSheet))?.freezeRows ?? 0;
      if (freezeRows > 0) pages.add(0);
      sendPageSubscription(activeSheet, [...pages].sort((a, b) => a - b));
    }, 200);
  }

  // 세로 스크롤 때마다 화면에 보이는 범위만 다시 그린다(rAF로 묶어 스크롤 중 매 프레임 다시
  // 그리는 걸 막는다) — 기존 커스텀 가로 스크롤바 리스너(updateHScrollbar)는 그대로 두고 이
  // 리스너는 추가로 붙는다(대체 아님). 편집 중(contentEditable)에는 재렌더가 그 <td>를 DOM에서
  // 떼어내며 강제로 blur를 발생시켜 방금 타이핑한 값을 잃을 수 있어(위 블러 핸들러의 세대 판정이
  // 그런 blur를 "커밋하지 않고 버림" 처리한다) 재렌더 자체를 건너뛴다 — 스크롤은 계속되고, 편집이
  // 끝나면(다음 스크롤이나 다른 재렌더 트리거에서) 화면이 다시 맞춰진다.
  //
  // renderTableVirtualized() 전체 재빌드 대신 updateScrollWindow()를 부른다 — 실측(Chrome
  // longtask) 기준 전체 재빌드는 스크롤 프레임 하나에 54~108ms(16ms 프레임 예산의 3~6배)가 걸려
  // 스크롤이 버벅이며 의도한 만큼 이동하지 않는 원인이었다. updateScrollWindow는 같은 문맥(시트/
  // 열 폭/틀고정/병합)에서 스크롤만 일어난 경우 <colgroup>/<thead>/창 안에서 안 바뀐 행은 그대로
  // 두고 창 가장자리에서 나가고 들어온 행만 patch한다 — 문맥이 조금이라도 의심되면 알아서
  // renderTableVirtualized()로 물러나므로 안전하다.
  let scrollRenderScheduled = false;
  scrollHostEl.addEventListener('scroll', () => {
    scheduleSubscriptionUpdate();
    if (editingKey !== null) return;
    if (scrollRenderScheduled) return;
    scrollRenderScheduled = true;
    requestAnimationFrame(() => {
      scrollRenderScheduled = false;
      if (editingKey !== null) return;
      updateScrollWindow();
    });
  });
  document.getElementById('hscroll-left')?.addEventListener('click', () => {
    scrollHostEl.scrollBy({ left: -100 });
  });
  document.getElementById('hscroll-right')?.addEventListener('click', () => {
    scrollHostEl.scrollBy({ left: 100 });
  });
  hscrollThumbEl?.addEventListener('mousedown', (event) => {
    event.preventDefault();
    const startX = event.clientX;
    const startScrollLeft = scrollHostEl.scrollLeft;
    const trackWidth = hscrollTrackEl!.clientWidth;
    const thumbWidth = hscrollThumbEl.getBoundingClientRect().width;
    const maxScrollLeft = scrollHostEl.scrollWidth - scrollHostEl.clientWidth;
    const onMove = (moveEvent: MouseEvent) => {
      const deltaPx = moveEvent.clientX - startX;
      const deltaScroll = (deltaPx / (trackWidth - thumbWidth)) * maxScrollLeft;
      scrollHostEl.scrollLeft = Math.max(0, Math.min(maxScrollLeft, startScrollLeft + deltaScroll));
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
  // 트랙의 빈 부분을 클릭하면 그 방향으로 한 화면씩 스크롤한다(엑셀과 동일).
  hscrollTrackEl?.addEventListener('mousedown', (event) => {
    if (event.target !== hscrollTrackEl) return;
    const trackRect = hscrollTrackEl.getBoundingClientRect();
    const clickedBeforeThumb = event.clientX < trackRect.left + (hscrollThumbEl?.offsetLeft ?? 0);
    scrollHostEl.scrollBy({ left: clickedBeforeThumb ? -scrollHostEl.clientWidth : scrollHostEl.clientWidth });
  });

  // ---------- 문서(시트 전체) 내 검색 ----------
  interface CellMatch {
    readonly sheetIdx: number;
    readonly row: number;
    readonly col: number;
  }
  const searchBar = document.getElementById('search-bar') as HTMLElement;
  const searchInput = document.getElementById('search-input') as HTMLInputElement;
  const searchCount = document.getElementById('search-count') as HTMLElement;
  const searchNextBtn = document.getElementById('search-next') as HTMLButtonElement;
  const searchScopeHintEl = document.getElementById('search-scope-hint') as HTMLElement | null;
  const searchReplaceRow = document.getElementById('search-replace-row') as HTMLElement | null;
  const searchReplaceInput = document.getElementById('search-replace-input') as HTMLInputElement | null;
  let searchMatches: CellMatch[] = [];
  let searchCurrent = -1;
  let lastHighlighted: HTMLTableCellElement | null = null;
  let allHighlighted: HTMLTableCellElement[] = [];
  // 엑셀처럼 — 다중 셀 범위를 선택한 채로 찾기/바꾸기를 열면 그 범위 안에서만 찾는다(셀 하나만
  // 선택돼 있었으면 시트 전체). 대화상자를 여는 시점의 선택을 스냅샷으로 고정해 둔다.
  let searchScope: { sheetIdx: number; r0: number; r1: number; c0: number; c1: number } | null = null;

  function findCellMatches(query: string): CellMatch[] {
    if (!query) return [];
    const lower = query.toLowerCase();
    const matches: CellMatch[] = [];
    for (const [key, value] of cells) {
      if (!value.toLowerCase().includes(lower)) continue;
      const parsed = parseGridCellKey(key);
      if (!parsed) continue;
      if (
        searchScope &&
        (parsed.sheetIdx !== searchScope.sheetIdx ||
          parsed.row < searchScope.r0 ||
          parsed.row > searchScope.r1 ||
          parsed.col < searchScope.c0 ||
          parsed.col > searchScope.c1)
      ) {
        continue;
      }
      matches.push(parsed);
    }
    matches.sort((a, b) => a.sheetIdx - b.sheetIdx || a.row - b.row || a.col - b.col);
    return matches;
  }

  function clearHighlight(): void {
    lastHighlighted?.classList.remove('g-hit-current');
    lastHighlighted = null;
  }
  function clearAllHighlight(): void {
    allHighlighted.forEach((td) => td.classList.remove('g-hit-all'));
    allHighlighted = [];
  }

  function jumpToMatch(index: number): void {
    clearHighlight();
    const match = searchMatches[index];
    if (!match) return;
    if (match.sheetIdx !== activeSheet) {
      activeSheet = match.sheetIdx;
      // 검색 결과로 다른 시트로 건너뛰는 것도 시트 전환이다 — 이전 시트에서 잡아 둔 선택 범위를
      // 그대로 들고 가면 안 된다(탭 클릭 전환과 같은 이유, resetSelectionForSheetSwitch 주석 참고).
      resetSelectionForSheetSwitch();
      renderSheetTabs();
      renderTableVirtualized();
    }
    // 가상 스크롤 범위 밖(다른 페이지)의 매치일 수 있다 — 필요하면 스크롤해 그 행을 그린 뒤 찾는다.
    const td = tdForEnsureRendered(match.row, match.col);
    if (td) {
      td.classList.add('g-hit-current');
      td.scrollIntoView({ block: 'center', inline: 'center' });
      lastHighlighted = td;
    }
  }

  function applySearch(jumpTo: number): void {
    clearAllHighlight();
    searchMatches = findCellMatches(searchInput.value);
    searchCurrent =
      searchMatches.length > 0 ? ((jumpTo % searchMatches.length) + searchMatches.length) % searchMatches.length : -1;
    searchCount.textContent = searchMatches.length > 0 ? `${searchCurrent + 1}/${searchMatches.length}` : searchInput.value ? '0/0' : '';
    searchNextBtn.disabled = searchMatches.length === 0;
    if (searchCurrent >= 0) jumpToMatch(searchCurrent);
    else clearHighlight();
  }
  /** "모두 찾기" — 엑셀의 결과 목록 대신, 매치되는 칸을 전부(연한 노랑) 칠하고 첫 매치로 이동해
   * (진한 노랑) 개수를 보여준다 — 지금 화면(활성 시트)에 실제로 그려진 칸만 칠할 수 있다. */
  function findAllMatches(): void {
    clearAllHighlight();
    searchMatches = findCellMatches(searchInput.value);
    searchCurrent = searchMatches.length > 0 ? 0 : -1;
    searchCount.textContent = searchMatches.length > 0 ? `${searchMatches.length}개 찾음` : searchInput.value ? '0개 찾음' : '';
    searchNextBtn.disabled = searchMatches.length === 0;
    if (searchMatches.length === 0) {
      clearHighlight();
      return;
    }
    searchMatches.forEach((m) => {
      if (m.sheetIdx !== activeSheet) return;
      const td = tdFor(m.row, m.col);
      if (td) {
        td.classList.add('g-hit-all');
        allHighlighted.push(td);
      }
    });
    jumpToMatch(0);
  }
  // ---------- 찾기/바꾸기 창 드래그 이동(엑셀처럼 제목줄을 잡고 옮길 수 있게) ----------
  const searchTitlebar = searchBar.querySelector('.search-titlebar') as HTMLElement | null;
  if (searchTitlebar) {
    let dragStart: { pointerX: number; pointerY: number; left: number; top: number } | null = null;

    searchTitlebar.addEventListener('mousedown', (event) => {
      event.preventDefault();
      const rect = searchBar.getBoundingClientRect();
      // 중앙 정렬(left:50%+transform)로 열려 있던 상태라면, 드래그 시작 시점의 실제 화면 위치를
      // 고정된 left/top 픽셀값으로 바꿔둔다 — 그래야 이후 이동량을 그대로 더해도 어긋나지 않는다.
      searchBar.style.left = `${rect.left}px`;
      searchBar.style.top = `${rect.top}px`;
      searchBar.style.transform = 'none';
      dragStart = { pointerX: event.clientX, pointerY: event.clientY, left: rect.left, top: rect.top };
    });

    document.addEventListener('mousemove', (event) => {
      if (!dragStart) return;
      const rect = searchBar.getBoundingClientRect();
      const maxLeft = Math.max(window.innerWidth - rect.width, 0);
      const maxTop = Math.max(window.innerHeight - rect.height, 0);
      const nextLeft = dragStart.left + (event.clientX - dragStart.pointerX);
      const nextTop = dragStart.top + (event.clientY - dragStart.pointerY);
      searchBar.style.left = `${Math.min(Math.max(nextLeft, 0), maxLeft)}px`;
      searchBar.style.top = `${Math.min(Math.max(nextTop, 0), maxTop)}px`;
    });

    document.addEventListener('mouseup', () => {
      dragStart = null;
    });
  }

  function openSearch(withReplace = false): void {
    searchBar.hidden = false;
    const range = rangeBounds();
    const hasRealSelection = range.r0 !== range.r1 || range.c0 !== range.c1;
    searchScope = hasRealSelection ? { sheetIdx: activeSheet, ...range } : null;
    if (searchScopeHintEl) searchScopeHintEl.hidden = !hasRealSelection;
    if (withReplace && searchReplaceRow) searchReplaceRow.hidden = false;
    searchInput.focus();
    searchInput.select();
  }
  function closeSearch(): void {
    searchBar.hidden = true;
    searchInput.value = '';
    clearHighlight();
    clearAllHighlight();
    searchMatches = [];
    searchCurrent = -1;
    searchCount.textContent = '';
    searchScope = null;
    if (searchScopeHintEl) searchScopeHintEl.hidden = true;
    if (searchReplaceRow) searchReplaceRow.hidden = true;
  }
  document.getElementById('search-btn')?.addEventListener('click', () => {
    if (searchBar.hidden) openSearch();
    else closeSearch();
  });
  document.getElementById('btn-find-replace')?.addEventListener('click', () => {
    if (!canEdit) return;
    if (searchBar.hidden) openSearch(true);
    else closeSearch();
  });
  document.getElementById('search-close')?.addEventListener('click', closeSearch);
  document.getElementById('search-find-all')?.addEventListener('click', () => findAllMatches());
  searchInput.addEventListener('input', () => applySearch(0));
  searchInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') applySearch(searchCurrent + (event.shiftKey ? -1 : 1));
    if (event.key === 'Escape') closeSearch();
  });
  searchNextBtn.addEventListener('click', () => applySearch(searchCurrent + 1));

  // ---------- 찾아서 바꾸기 ----------
  function escapeForRegExp(text: string): string {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  function replaceInText(text: string, find: string, replace: string): string {
    if (!find) return text;
    return text.replace(new RegExp(escapeForRegExp(find), 'gi'), replace);
  }
  document.getElementById('search-replace-one')?.addEventListener('click', () => {
    if (!canEdit) return;
    const match = searchMatches[searchCurrent];
    if (!match) return;
    const key = gridCellKey(match.sheetIdx, match.row, match.col);
    if (styles.get(key)?.lk === 1) {
      applySearch(searchCurrent + 1);
      return;
    }
    const next = replaceInText(cells.get(key) ?? '', searchInput.value, searchReplaceInput?.value ?? '');
    ydoc.transact(() => {
      if (next === '') cells.delete(key);
      else cells.set(key, next);
    });
    applySearch(searchCurrent); // 방금 바꾼 칸이 더 이상 검색어와 안 맞으면 자연히 다음 매치로 넘어간다
  });
  document.getElementById('search-replace-all')?.addEventListener('click', () => {
    if (!canEdit || !searchInput.value) return;
    const findText = searchInput.value;
    const replaceText = searchReplaceInput?.value ?? '';
    const matches = findCellMatches(findText);
    ydoc.transact(() => {
      matches.forEach((m) => {
        const key = gridCellKey(m.sheetIdx, m.row, m.col);
        if (styles.get(key)?.lk === 1) return;
        const current = cells.get(key) ?? '';
        const next = replaceInText(current, findText, replaceText);
        if (next === current) return;
        if (next === '') cells.delete(key);
        else cells.set(key, next);
      });
    });
    applySearch(0);
  });

  window.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      openSearch();
      return;
    }
    // 엑셀과 동일하게 Ctrl+H는 "찾기 및 바꾸기"를 바꾸기 탭이 펼쳐진 채로 연다.
    if (canEdit && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'h') {
      event.preventDefault();
      openSearch(true);
      return;
    }
    // Ctrl/Cmd+Z(실행 취소), Ctrl/Cmd+Y(다시 실행) — 셀 편집 중(입력창에 포커스)이 아닐 때만
    // 그리드 단위 undo/redo로 처리한다. 편집 중이면 브라우저의 기본 텍스트 입력 undo가 자연스럽다.
    const target = event.target as HTMLElement | null;
    const isTextInput = target?.tagName === 'INPUT' || target?.tagName === 'TEXTAREA' || target?.isContentEditable;
    if (isTextInput) return;
    if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === 'z') {
      event.preventDefault();
      undoManager.undo();
    } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y') {
      event.preventDefault();
      undoManager.redo();
    }
  });

  // ---------- 수식 입력줄: 선택된 셀 값을 그대로 편집할 수 있는 별도 입력창 ----------
  formulaInputEl.addEventListener('input', () => {
    if (!activeCellKey) return;
    const parsed = parseGridCellKey(activeCellKey);
    const td = parsed ? tdFor(parsed.row, parsed.col) : null;
    if (td) td.textContent = formulaInputEl.value;
  });
  function commitFormulaBar(): void {
    if (!activeCellKey) return;
    const key = activeCellKey;
    const parsed = parseGridCellKey(key);
    const td = parsed ? tdFor(parsed.row, parsed.col) : null;
    if (!td) return;
    commitCell(td);
    renderCellDisplay(td, key);
    recalcFormulaCells();
  }
  formulaInputEl.addEventListener('change', commitFormulaBar);
  formulaInputEl.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      commitFormulaBar();
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
