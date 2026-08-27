import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import FormulaParser, { FormulaError, type CellRef, type RangeRef } from 'fast-formula-parser';
import {
  EDIT_SOFT_COL_LIMIT,
  EDIT_SOFT_ROW_LIMIT,
  GRID_CELLS_MAP,
  GRID_STRUCTURE_MAP,
  GRID_STYLES_MAP,
  gridCellKey,
  parseGridCellKey,
  type CellStyleAttrs,
  type SheetStructure,
} from '../collab/gridShared.js';
import { formulaFunctions, formulaFunctionsNeedContext } from './formulaOverrides.js';
import { pickPresenceColor, renderPresenceAvatars, type AwarenessUserState } from './presence.js';

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

const params = new URLSearchParams(location.search);
const token = params.get('token') ?? '';

const titleEl = document.getElementById('doc-title') as HTMLElement;
const statusEl = document.getElementById('conn-status') as HTMLElement;
const presenceEl = document.getElementById('presence-list') as HTMLElement;
const warningEl = document.getElementById('unsupported-warning') as HTMLElement;
const editorHost = document.getElementById('grid-editor') as HTMLElement;
const zoomHost = document.getElementById('grid-zoom-host') as HTMLElement;
const tableEl = document.getElementById('grid-edit-table') as HTMLTableElement;
const sheetTabsEl = document.getElementById('sheet-tabs') as HTMLElement;
const cellRefBoxEl = document.getElementById('cell-ref-box') as HTMLElement;
const formulaInputEl = document.getElementById('formula-input') as HTMLInputElement;
const toolbar = document.getElementById('toolbar') as HTMLElement;
const errorScreen = document.getElementById('error-screen') as HTMLElement;
const errorMessage = document.getElementById('error-message') as HTMLElement;

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

function applyStyleToTd(td: HTMLTableCellElement, style: CellStyleAttrs | undefined): void {
  td.style.fontWeight = style?.b ? '700' : '400';
  td.style.fontStyle = style?.i ? 'italic' : 'normal';
  td.style.textDecoration = [style?.u ? 'underline' : '', style?.st ? 'line-through' : ''].filter(Boolean).join(' ');
  td.style.color = style?.col ?? '';
  td.style.background = style?.bg ?? '';
  if (style?.a) {
    td.style.textAlign = ALIGN_TO_CSS[style.a];
  } else {
    td.style.textAlign = NUMERIC_VALUE.test((td.textContent ?? '').trim()) ? 'right' : '';
  }
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

  const ydoc = new Y.Doc();
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);

  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, {
    params: { token },
  });

  const localColor = pickPresenceColor(ydoc.clientID);
  provider.awareness.setLocalStateField('user', { name: info.viewer, color: localColor });

  provider.on('status', (event: { status: string }) => {
    statusEl.textContent =
      event.status === 'connected' ? '연결됨' : event.status === 'disconnected' ? '연결 끊김 — 재연결 중…' : event.status;
    statusEl.dataset.status = event.status;
  });
  provider.on('sync', (isSynced: boolean) => {
    if (isSynced) {
      statusEl.dataset.status = 'synced';
      renderSheetTabs();
      renderTable();
    }
  });
  provider.awareness.on('change', () => {
    renderPresenceAvatars(presenceEl, provider.awareness);
    renderRemoteCellHighlights();
  });
  renderPresenceAvatars(presenceEl, provider.awareness);

  const undoManager = new Y.UndoManager([cells, structure, styles]);

  let activeSheet = 0;
  let activeCellKey: string | null = null;

  // ---------- 수식 계산 엔진: '='로 시작하는 셀은 실시간으로 계산 결과를 보여준다 ----------
  // fast-formula-parser(MIT)를 쓴다 — 자체 수식 파서/계산기를 새로 만들면 SUM/IF/VLOOKUP/범위
  // 참조까지 정확히 구현하기엔 너무 크고 위험한 작업이라, 검증된 라이브러리에 셀 조회만 연결한다.
  // 이 라이브러리는 VLOOKUP/HLOOKUP/INDEX는 실제로 구현돼 있지만, MAX/MIN/MATCH/CHOOSE/SWITCH/
  // SUMIFS/AVERAGEIFS를 포함한 일부는 내부적으로 본문이 빈 스텁이라(소스 직접 확인) 자주 쓰는
  // 것들만 formulaOverrides.ts에서 직접 구현해 덮어썼다. LARGE/SMALL/MEDIAN/TEXTJOIN 등 나머지
  // 스텁 함수는 v1 스코프 밖으로 남겨둔다.
  const evalGuard = new Set<string>();
  // 이 라이브러리는 셀/범위 좌표를 1-based로 주고받는다 — 우리 내부 스킴(0-based)과 여기서만 변환한다.
  const formulaParser = new FormulaParser({
    functions: formulaFunctions,
    functionsNeedContext: formulaFunctionsNeedContext,
    onCell: (ref: CellRef) => computeCellDisplay(activeSheet, ref.row - 1, ref.col - 1),
    onRange: (ref: RangeRef) => {
      const rows: unknown[][] = [];
      for (let r = ref.from.row; r <= ref.to.row; r += 1) {
        const row: unknown[] = [];
        for (let c = ref.from.col; c <= ref.to.col; c += 1) row.push(computeCellDisplay(activeSheet, r - 1, c - 1));
        rows.push(row);
      }
      return rows;
    },
  });

  /** 셀 하나의 "표시용" 값을 계산한다 — 수식이면 계산 결과(에러면 FormulaError.name, 예: #DIV/0!),
   * 아니면 원본 텍스트. 다른 수식이 이 값을 참조할 때는 계산 실패를 빈 값으로 흡수해 전체가
   * 깨지지 않게 한다(예: SUM 범위 안에 에러 셀이 하나 있어도 나머지는 정상 합산). */
  function computeCellDisplay(sheetIdx: number, row: number, col: number): unknown {
    const key = gridCellKey(sheetIdx, row, col);
    const raw = cells.get(key) ?? '';
    if (!raw.startsWith('=')) {
      if (raw.trim() === '') return '';
      const num = Number(raw);
      return Number.isFinite(num) ? num : raw;
    }
    if (evalGuard.has(key)) return FormulaError.REF.name; // 순환 참조
    evalGuard.add(key);
    try {
      return formulaParser.parse(raw.slice(1), { row: row + 1, col: col + 1, sheet: 'Sheet1' });
    } catch (error) {
      return error instanceof FormulaError ? error.name : '#ERROR!';
    } finally {
      evalGuard.delete(key);
    }
  }

  function formatComputedValue(value: unknown): string {
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (typeof value !== 'number') return value === null || value === undefined ? '' : String(value);
    // 부동소수점 오차(0.1+0.2=0.30000000000000004류)를 다듬는다
    return String(Math.round(value * 1e10) / 1e10);
  }

  /** 셀 하나를 "지금 이 셀이 포커스 중인지"에 맞춰 원본 수식 또는 계산 결과로 그린다. */
  function renderCellDisplay(td: HTMLTableCellElement, key: string): void {
    const raw = cells.get(key) ?? '';
    if (document.activeElement === td) {
      td.textContent = raw; // 편집 중에는 계산 결과가 아니라 실제 입력한 수식이 보여야 한다
      return;
    }
    if (!raw.startsWith('=')) {
      td.textContent = raw;
      return;
    }
    const parsed = parseGridCellKey(key);
    td.textContent = parsed ? formatComputedValue(computeCellDisplay(parsed.sheetIdx, parsed.row, parsed.col)) : raw;
  }

  /** 어떤 셀이든 값이 바뀌면, 그 값을 참조할 수 있는 다른 수식 셀들도 다시 계산해 보여준다
   * (의존성 그래프 없이 "지금 그려진 수식 셀 전부 다시 계산"으로 단순하게 처리 — 편집기 규모에서
   * 충분히 빠르다). */
  function recalcFormulaCells(): void {
    tableEl.querySelectorAll<HTMLTableCellElement>('td[contenteditable]').forEach((td) => {
      if (document.activeElement === td) return;
      const row = Number(td.dataset.row);
      const col = Number(td.dataset.col);
      const key = gridCellKey(activeSheet, row, col);
      if ((cells.get(key) ?? '').startsWith('=')) renderCellDisplay(td, key);
    });
  }

  function sortedSheetIndices(): number[] {
    return [...structure.keys()].map(Number).sort((a, b) => a - b);
  }

  function renderSheetTabs(): void {
    const indices = sortedSheetIndices();
    sheetTabsEl.replaceChildren(
      ...indices.map((idx) => {
        const sheet = structure.get(String(idx))!;
        const tab = document.createElement('button');
        tab.type = 'button';
        tab.className = 'grid-sheet-tab' + (idx === activeSheet ? ' active' : '');
        tab.textContent = sheet.name;
        tab.addEventListener('click', () => {
          if (activeSheet === idx) return;
          activeSheet = idx;
          renderSheetTabs();
          renderTable();
        });
        return tab;
      }),
    );
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

  /** 데이터가 있는 범위 너머로도 빈 칸을 넉넉히 그려서, 실제 엑셀처럼 "빈 시트도 격자가 있는"
   * 느낌을 준다 — 그 빈 칸에 입력하면 commitCell이 자동으로 범위를 넓혀준다. */
  const GRID_ROW_PADDING = 15;
  const GRID_COL_PADDING = 6;
  const GRID_MIN_RENDER_ROWS = 25;
  const GRID_MIN_RENDER_COLS = 12;

  function renderTable(): void {
    const sheet = structure.get(String(activeSheet));
    tableEl.replaceChildren();
    if (!sheet) return;

    warningEl.hidden = !(sheet.rowCount > EDIT_SOFT_ROW_LIMIT || sheet.colCount > EDIT_SOFT_COL_LIMIT);
    const renderRowCount = Math.max(sheet.rowCount + GRID_ROW_PADDING, GRID_MIN_RENDER_ROWS);
    const renderColCount = Math.max(sheet.colCount + GRID_COL_PADDING, GRID_MIN_RENDER_COLS);

    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    headRow.appendChild(document.createElement('th'));
    for (let c = 0; c < renderColCount; c += 1) {
      const th = document.createElement('th');
      th.textContent = colLetter(c);
      th.dataset.col = String(c);
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    tableEl.appendChild(thead);

    const tbody = document.createElement('tbody');
    for (let r = 0; r < renderRowCount; r += 1) {
      const tr = document.createElement('tr');
      const rowHead = document.createElement('th');
      rowHead.className = 'g-edit-rownum';
      rowHead.textContent = String(r + 1);
      rowHead.dataset.row = String(r);
      tr.appendChild(rowHead);
      for (let c = 0; c < renderColCount; c += 1) {
        const key = gridCellKey(activeSheet, r, c);
        const td = document.createElement('td');
        td.contentEditable = 'true';
        td.dataset.row = String(r);
        td.dataset.col = String(c);
        renderCellDisplay(td, key);
        applyStyleToTd(td, styles.get(key));
        td.addEventListener('focus', () => {
          activeCellKey = key;
          provider.awareness.setLocalStateField('cell', { sheetIdx: activeSheet, row: r, col: c });
          syncCellToolbarState(key);
          setActiveHeaders(r, c);
          cellRefBoxEl.textContent = cellRefLabel(r, c);
          formulaInputEl.value = cells.get(key) ?? '';
          renderCellDisplay(td, key); // 편집 모드로 진입 — 계산 결과 대신 원본 수식/값을 보여준다
        });
        td.addEventListener('input', () => {
          if (activeCellKey === key) formulaInputEl.value = td.textContent ?? '';
        });
        td.addEventListener('blur', () => {
          commitCell(td);
          renderCellDisplay(td, key); // 편집 종료 — 수식이면 계산 결과로 되돌린다
          recalcFormulaCells(); // 이 셀 값이 바뀌었으니 이를 참조하는 다른 수식 셀도 갱신
        });
        td.addEventListener('keydown', (event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            td.blur();
          }
        });
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    tableEl.appendChild(tbody);
    renderRemoteCellHighlights();
    if (activeCellKey) {
      const parsed = parseGridCellKey(activeCellKey);
      if (parsed && parsed.sheetIdx === activeSheet) setActiveHeaders(parsed.row, parsed.col);
    }
  }

  function clearActiveHeaders(): void {
    tableEl.querySelectorAll('.g-col-active').forEach((el) => el.classList.remove('g-col-active'));
    tableEl.querySelectorAll('.g-row-active').forEach((el) => el.classList.remove('g-row-active'));
  }
  function setActiveHeaders(row: number, col: number): void {
    clearActiveHeaders();
    tableEl.querySelector(`thead th[data-col="${col}"]`)?.classList.add('g-col-active');
    tableEl.querySelector(`tbody th[data-row="${row}"]`)?.classList.add('g-row-active');
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

  interface RemoteCellState extends AwarenessUserState {
    readonly cell?: { readonly sheetIdx: number; readonly row: number; readonly col: number };
  }

  /** 다른 참여자가 지금 편집 중인 셀에 색 테두리 + 작은 이름표를 띄운다("누가 어디를 수정 중인지"). */
  function renderRemoteCellHighlights(): void {
    tableEl.querySelectorAll<HTMLTableCellElement>('.g-remote-cell').forEach((td) => {
      td.classList.remove('g-remote-cell');
      td.style.removeProperty('--remote-color');
      td.querySelector('.g-remote-tag')?.remove();
    });
    const states = provider.awareness.getStates() as Map<number, RemoteCellState>;
    for (const [clientId, state] of states) {
      if (clientId === ydoc.clientID) continue;
      const cell = state.cell;
      const name = state.user?.name;
      if (!cell || !name || cell.sheetIdx !== activeSheet) continue;
      const td = tdFor(cell.row, cell.col);
      if (!td) continue;
      const color = state.user?.color ?? pickPresenceColor(clientId);
      td.classList.add('g-remote-cell');
      td.style.setProperty('--remote-color', color);
      const tag = document.createElement('span');
      tag.className = 'g-remote-tag';
      tag.style.setProperty('--remote-color', color);
      tag.textContent = name;
      td.appendChild(tag);
    }
  }

  cells.observe((event) => {
    for (const key of event.keysChanged) {
      const parsed = parseGridCellKey(key);
      if (!parsed || parsed.sheetIdx !== activeSheet) continue;
      const td = tdFor(parsed.row, parsed.col);
      if (!td || document.activeElement === td) continue; // 로컬에서 입력 중인 셀은 원격 갱신으로 덮지 않는다
      const value = cells.get(key) ?? '';
      renderCellDisplay(td, key);
      applyStyleToTd(td, styles.get(key));
      if (key === activeCellKey && document.activeElement !== formulaInputEl) formulaInputEl.value = value;
    }
    recalcFormulaCells(); // 바뀐 셀을 참조하는 다른 수식 셀들도 다시 계산
  });

  styles.observe((event) => {
    for (const key of event.keysChanged) {
      const parsed = parseGridCellKey(key);
      if (!parsed || parsed.sheetIdx !== activeSheet) continue;
      const td = tdFor(parsed.row, parsed.col);
      if (!td || document.activeElement === td) continue;
      applyStyleToTd(td, styles.get(key));
    }
  });

  structure.observe(() => {
    renderSheetTabs();
    renderTable();
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
  document.getElementById('btn-undo')?.addEventListener('click', () => undoManager.undo());
  document.getElementById('btn-redo')?.addEventListener('click', () => undoManager.redo());

  // ---------- 서식 도구 (현재 포커스된 셀에 적용, 오피스처럼 눌린 상태로 반영) ----------
  function setPressed(id: string, pressed: boolean): void {
    document.getElementById(id)?.setAttribute('aria-pressed', String(pressed));
  }
  function syncCellToolbarState(key: string | null): void {
    const style = key ? styles.get(key) : undefined;
    setPressed('btn-cell-bold', style?.b === 1);
    setPressed('btn-cell-italic', style?.i === 1);
    setPressed('btn-cell-underline', style?.u === 1);
    setPressed('btn-cell-strike', style?.st === 1);
    setPressed('btn-cell-align-l', !style?.a || style.a === 'l');
    setPressed('btn-cell-align-c', style?.a === 'c');
    setPressed('btn-cell-align-r', style?.a === 'r');
  }

  function updateStyle(mutate: (current: CellStyleAttrs) => CellStyleAttrs): void {
    if (!activeCellKey) return;
    const current = styles.get(activeCellKey) ?? {};
    const next = mutate(current);
    if (Object.keys(next).length === 0) styles.delete(activeCellKey);
    else styles.set(activeCellKey, next);
    // 툴바 버튼을 누르면 버튼이 포커스를 가져가 셀은 이미 blur된 상태이므로 :focus로는 못 찾는다
    // — activeCellKey(마지막으로 선택했던 셀)를 좌표로 풀어 직접 찾는다.
    const parsed = parseGridCellKey(activeCellKey);
    const td = parsed ? tdFor(parsed.row, parsed.col) : null;
    if (td) applyStyleToTd(td, styles.get(activeCellKey));
    syncCellToolbarState(activeCellKey);
  }
  function toggleFlag(flag: 'b' | 'i' | 'u' | 'st'): void {
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
  document.getElementById('btn-cell-clear-format')?.addEventListener('click', () => {
    if (!activeCellKey) return;
    styles.delete(activeCellKey);
    const parsed = parseGridCellKey(activeCellKey);
    const td = parsed ? tdFor(parsed.row, parsed.col) : null;
    if (td) applyStyleToTd(td, undefined);
    syncCellToolbarState(activeCellKey);
  });

  const cellColorInput = document.getElementById('cell-font-color') as HTMLInputElement | null;
  const cellColorBtn = document.getElementById('cell-font-color-btn') as HTMLElement | null;
  cellColorBtn?.style.setProperty('--btn-color', cellColorInput?.value ?? '#1b1a17');
  cellColorInput?.addEventListener('input', () => {
    cellColorBtn?.style.setProperty('--btn-color', cellColorInput.value);
    updateStyle((current) => ({ ...current, col: cellColorInput.value }));
  });
  const cellBgInput = document.getElementById('cell-bg-color') as HTMLInputElement | null;
  const cellBgBtn = document.getElementById('cell-bg-color-btn') as HTMLElement | null;
  cellBgBtn?.style.setProperty('--btn-color', cellBgInput?.value ?? '#ffff00');
  cellBgInput?.addEventListener('input', () => {
    cellBgBtn?.style.setProperty('--btn-color', cellBgInput.value);
    updateStyle((current) => ({ ...current, bg: cellBgInput.value }));
  });

  (['l', 'c', 'r'] as const).forEach((align) => {
    document.getElementById(`btn-cell-align-${align}`)?.addEventListener('click', () => {
      updateStyle((current) => ({ ...current, a: align }));
    });
  });

  // ---------- 확대/축소 ----------
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

  // ---------- 문서(시트 전체) 내 검색 ----------
  interface CellMatch {
    readonly sheetIdx: number;
    readonly row: number;
    readonly col: number;
  }
  const searchBar = document.getElementById('search-bar') as HTMLElement;
  const searchInput = document.getElementById('search-input') as HTMLInputElement;
  const searchCount = document.getElementById('search-count') as HTMLElement;
  const searchPrevBtn = document.getElementById('search-prev') as HTMLButtonElement;
  const searchNextBtn = document.getElementById('search-next') as HTMLButtonElement;
  let searchMatches: CellMatch[] = [];
  let searchCurrent = -1;
  let lastHighlighted: HTMLTableCellElement | null = null;

  function findCellMatches(query: string): CellMatch[] {
    if (!query) return [];
    const lower = query.toLowerCase();
    const matches: CellMatch[] = [];
    for (const [key, value] of cells) {
      if (!value.toLowerCase().includes(lower)) continue;
      const parsed = parseGridCellKey(key);
      if (parsed) matches.push(parsed);
    }
    matches.sort((a, b) => a.sheetIdx - b.sheetIdx || a.row - b.row || a.col - b.col);
    return matches;
  }

  function clearHighlight(): void {
    lastHighlighted?.classList.remove('g-hit-current');
    lastHighlighted = null;
  }

  function jumpToMatch(index: number): void {
    clearHighlight();
    const match = searchMatches[index];
    if (!match) return;
    if (match.sheetIdx !== activeSheet) {
      activeSheet = match.sheetIdx;
      renderSheetTabs();
      renderTable();
    }
    const td = tdFor(match.row, match.col);
    if (td) {
      td.classList.add('g-hit-current');
      td.scrollIntoView({ block: 'center', inline: 'center' });
      lastHighlighted = td;
    }
  }

  function applySearch(jumpTo: number): void {
    searchMatches = findCellMatches(searchInput.value);
    searchCurrent =
      searchMatches.length > 0 ? ((jumpTo % searchMatches.length) + searchMatches.length) % searchMatches.length : -1;
    searchCount.textContent = searchMatches.length > 0 ? `${searchCurrent + 1}/${searchMatches.length}` : searchInput.value ? '0/0' : '';
    searchPrevBtn.disabled = searchMatches.length === 0;
    searchNextBtn.disabled = searchMatches.length === 0;
    if (searchCurrent >= 0) jumpToMatch(searchCurrent);
    else clearHighlight();
  }
  function openSearch(): void {
    searchBar.hidden = false;
    searchInput.focus();
    searchInput.select();
  }
  function closeSearch(): void {
    searchBar.hidden = true;
    searchInput.value = '';
    clearHighlight();
    searchMatches = [];
    searchCurrent = -1;
    searchCount.textContent = '';
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
