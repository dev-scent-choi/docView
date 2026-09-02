'use strict';

const ZOOM_STEPS = [50, 67, 80, 100, 125, 150, 200, 250, 300, 400, 500];
const DEFAULT_ZOOM_INDEX = 3;

const token = new URLSearchParams(location.search).get('token');
const pagesContainer = document.getElementById('pages');
const docTitle = document.getElementById('doc-title');
const pageIndicator = document.getElementById('page-indicator');
const zoomLevelLabel = document.getElementById('zoom-level');
const printBtn = document.getElementById('print-btn');
const downloadBtn = document.getElementById('download-btn');

let zoomIndex = DEFAULT_ZOOM_INDEX;
let pageCount = 0;
let permissions = { canPrint: false, canDownload: false };
let docName = 'document';
let pageDimensions = null; // [{w,h}] — 로딩 전 프레임 크기 고정용

/** content-length 기반 진행률과 함께 응답 본문을 받아온다. */
async function fetchWithProgress(url, onProgress) {
  const res = await fetch(url);
  if (!res.ok) {
    const json = await res.json().catch(() => null);
    throw new Error((json && json.error) || `요청 실패 (HTTP ${res.status})`);
  }
  const total = Number(res.headers.get('content-length')) || 0;
  const reader = res.body.getReader();
  const chunks = [];
  let received = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    received += value.length;
    onProgress(received, total);
  }
  return chunks;
}

function formatProgress(received, total) {
  if (total > 0) return `${Math.round((received / total) * 100)}%`;
  return `${(received / 1048576).toFixed(1)}MB`;
}

/** 버튼을 진행률 표시기로 전환했다가 끝나면 복원한다. */
async function runButtonTask(button, progressPrefix, task) {
  const label = button.querySelector('.btn-label') ?? button;
  const originalText = label.textContent;
  button.disabled = true;
  try {
    await task((received, total) => {
      label.textContent = `${progressPrefix} ${formatProgress(received, total)}`;
    });
    label.textContent = originalText;
  } catch (error) {
    label.textContent = '실패';
    setTimeout(() => {
      label.textContent = originalText;
    }, 2500);
    throw error;
  } finally {
    button.disabled = false;
  }
}

function showError(message) {
  document.getElementById('toolbar').hidden = true;
  pagesContainer.hidden = true;
  document.getElementById('error-screen').hidden = false;
  document.getElementById('error-message').textContent = message;
}

function applyZoom() {
  const percent = ZOOM_STEPS[zoomIndex];
  zoomLevelLabel.textContent = `${percent}%`;
  pagesContainer.style.setProperty('--page-width', `${percent}%`);
  document.getElementById('zoom-out').disabled = zoomIndex === 0;
  document.getElementById('zoom-in').disabled = zoomIndex === ZOOM_STEPS.length - 1;
  scheduleQualityReload();
}

const tokenQuery = `token=${encodeURIComponent(token)}`;

/* 진행 중인 페이지·그리드 로딩을 뒤로가기 시 즉시 취소해 네비게이션 지연을 없앤다 */
let viewAbort = new AbortController();

function isAbort(error) {
  return error && error.name === 'AbortError';
}

/**
 * 확대 수준 × 디스플레이 픽셀 밀도에 맞는 서버 렌더링 배율.
 * 고해상도(Windows 배율 125~150%) 모니터에서는 CSS 픽셀보다 실제 픽셀이
 * 많아 배율을 더 올려야 원본처럼 선명하게 보인다.
 */
function desiredScale() {
  const percent = ZOOM_STEPS[zoomIndex];
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const scale = Math.ceil((percent / 100) * 2 * dpr);
  return Math.min(16, Math.max(2, scale));
}

/* 대용량 페이지 이미지를 동시에 너무 많이 받으면 메인 스레드가 밀려
   스크롤·버튼 반응이 늦어지므로 동시 로딩 수를 제한한다 */
const MAX_CONCURRENT_LOADS = 3;
let activeLoads = 0;
const loadWaiters = [];

async function withLoadSlot(task) {
  if (activeLoads >= MAX_CONCURRENT_LOADS) {
    await new Promise((resolve) => loadWaiters.push(resolve));
  }
  activeLoads += 1;
  try {
    return await task();
  } finally {
    activeLoads -= 1;
    const next = loadWaiters.shift();
    if (next) next();
  }
}

async function loadPageInto(canvas, pageNumber) {
  const scale = desiredScale();
  canvas.dataset.requestedScale = String(scale);
  await withLoadSlot(async () => {
    if (viewAbort.signal.aborted) return;
    let res;
    let bitmap;
    try {
      res = await fetch(`/api/view/pages/${pageNumber}?${tokenQuery}&scale=${scale}`, {
        signal: viewAbort.signal,
      });
      if (!res.ok) {
        canvas.closest('.page-frame').classList.add('page-error');
        return;
      }
      bitmap = await createImageBitmap(await res.blob());
    } catch (error) {
      if (isAbort(error)) return; // 뒤로가기 등으로 취소됨
      canvas.closest('.page-frame').classList.add('page-error');
      return;
    }
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    canvas.getContext('2d').drawImage(bitmap, 0, 0);
    bitmap.close();
    canvas.dataset.loadedScale = String(scale);
    const frame = canvas.closest('.page-frame');
    frame.style.aspectRatio = `${canvas.width} / ${canvas.height}`;
    frame.classList.remove('page-loading');
    frame.classList.add('page-ready');
  });
}

/**
 * 화면에서 멀리 벗어난 페이지의 비트맵을 해제한다.
 * 페이지가 많을 때 GPU 메모리가 쌓여 화질 저하·뒤로가기 지연이
 * 생기는 것을 막는다. 종횡비를 고정해 스크롤 위치는 유지된다.
 */
function releaseCanvas(canvas) {
  if (!canvas.dataset.loadedScale) return;
  const frame = canvas.closest('.page-frame');
  if (canvas.width > 0) {
    frame.style.aspectRatio = `${canvas.width} / ${canvas.height}`;
  }
  canvas.width = 0;
  canvas.height = 0;
  delete canvas.dataset.loadedScale;
  delete canvas.dataset.requestedScale;
  frame.classList.remove('page-ready');
  frame.classList.add('page-loading');
}

function ensureLoaded(canvas) {
  const target = desiredScale();
  const requested = Number(canvas.dataset.requestedScale || 0);
  if (requested >= target) return;
  loadPageInto(canvas, Number(canvas.dataset.page));
}

let qualityReloadTimer = null;

/** 확대 후 화면 근처에 보이는 페이지들을 더 높은 배율로 다시 렌더링한다. */
function scheduleQualityReload() {
  if (qualityReloadTimer !== null) clearTimeout(qualityReloadTimer);
  qualityReloadTimer = setTimeout(() => {
    qualityReloadTimer = null;
    const nearViewport = window.innerHeight * 2;
    for (const canvas of pagesContainer.querySelectorAll('canvas')) {
      if (!canvas.dataset.loadedScale) continue; // 화면 밖(해제됨)은 접근 시 로딩된다
      const rect = canvas.getBoundingClientRect();
      if (rect.bottom < -nearViewport || rect.top > window.innerHeight + nearViewport) continue;
      ensureLoaded(canvas);
    }
  }, 350);
}

function buildPages() {
  // 화면 근처(±1200px)에 들어오면 로딩, 멀어지면 비트맵 해제 (가상화)
  const observer = new IntersectionObserver((entries) => {
    for (const entry of entries) {
      const canvas = entry.target.querySelector('canvas');
      if (entry.isIntersecting) {
        ensureLoaded(canvas);
      } else {
        releaseCanvas(canvas);
      }
    }
  }, { rootMargin: '1200px 0px' });

  const frames = [];
  for (let page = 1; page <= pageCount; page += 1) {
    const frame = document.createElement('div');
    frame.className = 'page-frame page-loading';
    // 페이지 크기를 미리 알면 로딩 전에도 프레임 높이가 정확해 스크롤 점프가 어긋나지 않는다
    const dims = pageDimensions && pageDimensions[page - 1];
    if (dims && dims.w > 0) {
      frame.style.aspectRatio = `${dims.w} / ${dims.h}`;
      frame.style.minHeight = '0';
    }
    const canvas = document.createElement('canvas');
    canvas.dataset.page = String(page);
    const label = document.createElement('span');
    label.className = 'page-number';
    label.textContent = String(page);
    frame.append(canvas, label);
    frames.push(frame);
    observer.observe(frame);
  }
  pagesContainer.replaceChildren(...frames);
}

function trackCurrentPage() {
  const observer = new IntersectionObserver((entries) => {
    const visible = entries.find((entry) => entry.isIntersecting);
    if (visible) {
      const page = visible.target.querySelector('canvas').dataset.page;
      pageIndicator.textContent = `${page} / ${pageCount}`;
    }
  }, { threshold: 0.4 });
  for (const frame of pagesContainer.children) {
    observer.observe(frame);
  }
}

const INVERT_STORAGE_KEY = 'docview-invert';

function applyInvert(enabled) {
  pagesContainer.classList.toggle('pages-invert', enabled);
  document.getElementById('invert-btn').setAttribute('aria-pressed', String(enabled));
  try {
    localStorage.setItem(INVERT_STORAGE_KEY, enabled ? '1' : '0');
  } catch {
    /* 시크릿 모드 등에서 저장 실패해도 무시 */
  }
}

function navigateBack() {
  // 진행 중인 이미지·그리드 로딩을 즉시 취소해야 대형 문서에서도 뒤로가기가 바로 먹힌다
  viewAbort.abort();
  if (qualityReloadTimer !== null) clearTimeout(qualityReloadTimer);
  window.location.assign('/');
}

function setupToolbar() {
  const backBtn = document.getElementById('back-btn');
  backBtn.addEventListener('click', (event) => {
    event.preventDefault();
    navigateBack();
  });

  document.getElementById('zoom-in').addEventListener('click', () => {
    zoomIndex = Math.min(zoomIndex + 1, ZOOM_STEPS.length - 1);
    applyZoom();
  });
  document.getElementById('zoom-out').addEventListener('click', () => {
    zoomIndex = Math.max(zoomIndex - 1, 0);
    applyZoom();
  });

  const invertBtn = document.getElementById('invert-btn');
  invertBtn.addEventListener('click', () => {
    applyInvert(!pagesContainer.classList.contains('pages-invert'));
  });
  try {
    if (localStorage.getItem(INVERT_STORAGE_KEY) === '1') applyInvert(true);
  } catch {
    /* 저장소 접근 불가 시 기본값 유지 */
  }

  printBtn.disabled = !permissions.canPrint;
  printBtn.title = permissions.canPrint ? '워터마크가 포함된 출력용 PDF' : '출력 권한이 없습니다';
  if (permissions.canPrint) {
    printBtn.addEventListener('click', () => {
      const includeMarkup = document.getElementById('markup-include-checkbox')?.checked ? '&includeMarkup=1' : '';
      // 팝업 차단을 피하려면 사용자 클릭 스택 안에서 창을 먼저 열어둬야 한다
      const printWindow = window.open('', '_blank');
      runButtonTask(printBtn, '생성', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/print?${tokenQuery}${includeMarkup}`, onProgress);
        const blob = new Blob(chunks, { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        if (printWindow && !printWindow.closed) {
          printWindow.location.href = url;
        } else {
          const anchor = document.createElement('a');
          anchor.href = url;
          anchor.target = '_blank';
          anchor.rel = 'noopener';
          anchor.click();
        }
        setTimeout(() => URL.revokeObjectURL(url), 60_000);
      }).catch(() => {
        if (printWindow && !printWindow.closed) printWindow.close();
      });
    });
  }

  downloadBtn.disabled = !permissions.canDownload;
  downloadBtn.title = permissions.canDownload ? '원본 파일 다운로드' : '저장 권한이 없습니다';
  if (permissions.canDownload) {
    downloadBtn.addEventListener('click', () => {
      const includeMarkup = document.getElementById('markup-include-checkbox')?.checked ? '&includeMarkup=1' : '';
      runButtonTask(downloadBtn, '저장', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/download?${tokenQuery}${includeMarkup}`, onProgress);
        const blob = new Blob(chunks);
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = docName;
        anchor.click();
        setTimeout(() => URL.revokeObjectURL(url), 60_000);
      }).catch(() => {});
    });
  }
}

/* 우클릭/단축키 차단은 억지력일 뿐이며 실제 권한 강제는 서버에서 이루어진다. */
function setupDeterrents() {
  document.addEventListener('contextmenu', (event) => event.preventDefault());
  document.addEventListener('keydown', (event) => {
    const key = event.key.toLowerCase();
    if ((event.ctrlKey || event.metaKey) && (key === 'p' || key === 's')) {
      event.preventDefault();
      if (key === 'p' && permissions.canPrint) {
        printBtn.click();
      }
    }
    // 브라우저 기본 Ctrl+F는 가상화된 문서를 찾지 못하므로 자체 검색으로 대체한다
    if ((event.ctrlKey || event.metaKey) && key === 'f') {
      event.preventDefault();
      openSearchBar();
    }
    if (key === 'f3') {
      event.preventDefault();
      if (searchState.hits.length > 0) gotoHit(searchState.current + (event.shiftKey ? -1 : 1));
      else openSearchBar();
    }
    if (key === 'escape') {
      if (measure.mode) setMeasureMode(false);
      else if (markup.drawMode) setMarkupDrawMode(false);
      else closeSearchBar();
    }
  });
}

/* ==================== 문서 내 검색 ==================== */

const searchState = { hits: [], current: -1, kind: 'pages', lastQuery: '' };

function openSearchBar() {
  const bar = document.getElementById('search-bar');
  bar.hidden = false;
  document.getElementById('search-input').focus();
  document.getElementById('search-input').select();
}

function closeSearchBar() {
  document.getElementById('search-bar').hidden = true;
  clearSearchHighlights();
  searchState.hits = [];
  searchState.current = -1;
  updateSearchCount('');
}

function updateSearchCount(text) {
  document.getElementById('search-count').textContent = text;
  const hasHits = searchState.hits.length > 0;
  document.getElementById('search-prev').disabled = !hasHits;
  document.getElementById('search-next').disabled = !hasHits;
}

function clearSearchHighlights() {
  for (const el of document.querySelectorAll('.search-hit-overlay')) el.remove();
  gridSearchHits.clear();
  gridCurrentHit = null;
  if (typeof renderGridViewport === 'function' && gridState.sheets.length > 0) renderGridViewport();
}

async function runSearch() {
  const query = document.getElementById('search-input').value.trim();
  if (query === '' || query === searchState.lastQuery) {
    if (query === searchState.lastQuery && searchState.hits.length > 0) gotoHit(searchState.current + 1);
    return;
  }
  searchState.lastQuery = query;
  clearSearchHighlights();
  updateSearchCount('검색 중…');
  try {
    const data = await fetchJsonView(
      `/api/view/search?${tokenQuery}&q=${encodeURIComponent(query)}&limit=500`,
    );
    searchState.kind = data.kind;
    searchState.hits = data.hits;
    searchState.current = -1;
    if (data.hits.length === 0) {
      updateSearchCount('결과 없음');
      return;
    }
    if (data.kind === 'grid') {
      gridSearchHits = new Set(data.hits.map((h) => `${h.sheet}:${h.r}:${h.c}`));
    }
    gotoHit(0);
  } catch (error) {
    updateSearchCount(error.message || '검색 실패');
  }
}

function gotoHit(index) {
  if (searchState.hits.length === 0) return;
  const i = ((index % searchState.hits.length) + searchState.hits.length) % searchState.hits.length;
  searchState.current = i;
  const total = searchState.hits.length;
  const totalLabel = total >= 500 ? '500+' : String(total);
  updateSearchCount(`${i + 1} / ${totalLabel}`);
  const hit = searchState.hits[i];
  if (searchState.kind === 'grid') {
    gotoGridHit(hit);
  } else {
    gotoPageHit(hit);
  }
}

/** 모든 검색 결과를 연노랑으로, 현재 결과를 진한 주황 테두리로 표시한다 */
function renderPageHitOverlays() {
  for (const el of document.querySelectorAll('.search-hit-overlay')) el.remove();
  if (searchState.kind !== 'pages') return;
  searchState.hits.forEach((hit, index) => {
    const frame = pagesContainer.children[hit.page - 1];
    if (!frame) return;
    for (const rect of hit.rects) {
      const overlay = document.createElement('div');
      overlay.className =
        index === searchState.current ? 'search-hit-overlay current' : 'search-hit-overlay';
      overlay.style.left = `${rect.x * 100}%`;
      overlay.style.top = `${rect.y * 100}%`;
      overlay.style.width = `${Math.max(rect.w * 100, 0.5)}%`;
      overlay.style.height = `${Math.max(rect.h * 100, 0.5)}%`;
      frame.append(overlay);
    }
  });
}

/** 하이라이트가 이 크기(px)보다 작으면 읽을 수 있는 배율까지 자동 확대한다 */
const MIN_READABLE_HIT_PX = 11;
const TARGET_HIT_PX = 16;

function gotoPageHit(hit) {
  const frame = pagesContainer.children[hit.page - 1];
  if (!frame) return;
  renderPageHitOverlays();
  const current = frame.querySelector('.search-hit-overlay.current');

  // 도면처럼 결과 문자가 아주 작으면(전체 보기 기준 수 px) 자동으로 확대해 보여준다
  if (current) {
    const hitHeight = current.getBoundingClientRect().height;
    if (hitHeight > 0 && hitHeight < MIN_READABLE_HIT_PX) {
      const targetPercent = Math.min(
        ZOOM_STEPS[ZOOM_STEPS.length - 1],
        ZOOM_STEPS[zoomIndex] * (TARGET_HIT_PX / hitHeight),
      );
      const nextIndex = ZOOM_STEPS.findIndex((step) => step >= targetPercent);
      const resolved = nextIndex === -1 ? ZOOM_STEPS.length - 1 : nextIndex;
      if (resolved > zoomIndex) {
        zoomIndex = resolved;
        applyZoom();
      }
    }
  }

  centerOnElement(current ?? frame);
}

/**
 * 요소를 화면 정중앙으로 스크롤한다.
 * 확대 시 pages-container가 가로·세로 스크롤을 모두 갖게 되어 scrollIntoView가
 * 어긋나므로, 컨테이너 → 윈도우 순으로 직접 보정한다.
 */
function centerOnElement(el) {
  const rect = el.getBoundingClientRect();
  const containerRect = pagesContainer.getBoundingClientRect();

  const dx = rect.left + rect.width / 2 - (containerRect.left + pagesContainer.clientWidth / 2);
  pagesContainer.scrollLeft += dx;

  const dy = rect.top + rect.height / 2 - window.innerHeight / 2;
  const beforeTop = pagesContainer.scrollTop;
  pagesContainer.scrollTop += dy;
  const consumed = pagesContainer.scrollTop - beforeTop;
  if (Math.abs(dy - consumed) > 1) {
    window.scrollBy(0, dy - consumed);
  }
}

let gridSearchHits = new Set();
let gridCurrentHit = null;

function gotoGridHit(hit) {
  if (gridState.active !== hit.sheet) activateSheet(hit.sheet);
  gridCurrentHit = `${hit.sheet}:${hit.r}:${hit.c}`;
  const scroll = document.getElementById('grid-scroll');
  scroll.scrollTop = Math.max(0, (hit.r - 1) * GRID_ROW_H - scroll.clientHeight / 2);
  renderGridViewport();
}

function setupSearchBar() {
  document.getElementById('search-btn').addEventListener('click', openSearchBar);
  document.getElementById('search-close').addEventListener('click', closeSearchBar);
  document.getElementById('search-prev').addEventListener('click', () => gotoHit(searchState.current - 1));
  document.getElementById('search-next').addEventListener('click', () => gotoHit(searchState.current + 1));
  document.getElementById('search-input').addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      if (event.shiftKey) gotoHit(searchState.current - 1);
      else runSearch();
    }
  });
}

/* ==================== 그리드 뷰 (스프레드시트) ==================== */

const GRID_ROW_H = 28;
const GRID_CHUNK = 200;
const GRID_OVERSCAN = 12;
const GRID_ROWNUM_W = 52;
// 실제 데이터 범위 너머로도 빈 칸을 넉넉히 그려서 엑셀처럼 "빈 시트도 격자가 있는" 느낌을 준다
// (열람 전용이라 편집기처럼 입력 시 확장되지는 않지만, 화면 톤은 편집기와 맞춘다).
const GRID_ROW_PADDING = 30;
const GRID_COL_PADDING = 8;
const GRID_MIN_DISPLAY_ROWS = 40;
const GRID_MIN_DISPLAY_COLS = 14;
const GRID_DEFAULT_COL_WIDTH = 100;

function displayRowCount(sheet) {
  return Math.max(sheet.rowCount + GRID_ROW_PADDING, GRID_MIN_DISPLAY_ROWS);
}
function displayColCount(sheet) {
  return Math.max(sheet.colCount + GRID_COL_PADDING, GRID_MIN_DISPLAY_COLS);
}
function colWidthAt(sheet, c) {
  return sheet.colWidths[c - 1] ?? GRID_DEFAULT_COL_WIDTH;
}

const gridState = {
  sheets: [],
  styles: [],
  active: 0,
  chunks: new Map(), // `${sheet}:${chunkIdx}` -> rows | 'loading'
  mergeStarts: new Map(),
  mergeCovered: new Set(),
};

function colLetter(index) {
  let name = '';
  let n = index;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function injectGridStyles(styles) {
  const rules = styles.map((style, i) => {
    const parts = [];
    if (style.b) parts.push('font-weight:600');
    if (style.i) parts.push('font-style:italic');
    const lines = [];
    if (style.u) lines.push('underline');
    if (style.st) lines.push('line-through');
    if (lines.length > 0) parts.push(`text-decoration:${lines.join(' ')}`);
    if (style.fs) parts.push(`font-size:${style.fs}px`);
    if (style.col) parts.push(`color:${style.col}`);
    if (style.bg) parts.push(`background:${style.bg}`);
    if (style.a === 'c') parts.push('text-align:center');
    if (style.a === 'r') parts.push('text-align:right');
    if (style.w) parts.push('white-space:normal');
    return `.gs${i}{${parts.join(';')}}`;
  });
  const styleEl = document.createElement('style');
  styleEl.textContent = rules.join('\n');
  document.head.append(styleEl);
}

function buildMergeIndex(sheet) {
  gridState.mergeStarts = new Map();
  gridState.mergeCovered = new Set();
  for (const m of sheet.merges) {
    gridState.mergeStarts.set(`${m.r}:${m.c}`, m);
    for (let r = m.r; r < m.r + m.rs; r += 1) {
      for (let c = m.c; c < m.c + m.cs; c += 1) {
        if (r !== m.r || c !== m.c) gridState.mergeCovered.add(`${r}:${c}`);
      }
    }
  }
}

async function fetchChunk(sheetIdx, chunkIdx) {
  const key = `${sheetIdx}:${chunkIdx}`;
  if (gridState.chunks.has(key)) return;
  gridState.chunks.set(key, 'loading');
  try {
    const from = chunkIdx * GRID_CHUNK + 1;
    const data = await fetchJsonView(`/api/view/grid/rows?${tokenQuery}&sheet=${sheetIdx}&from=${from}&count=${GRID_CHUNK}`);
    gridState.chunks.set(key, data.rows);
    renderGridViewport();
  } catch {
    gridState.chunks.delete(key);
  }
}

async function fetchJsonView(url, options = {}) {
  const res = await fetch(url, { signal: viewAbort.signal, ...options });
  const json = await res.json();
  if (!res.ok || !json.success) throw new Error(json.error || `요청 실패 (HTTP ${res.status})`);
  return json.data;
}

function cellAt(sheetIdx, rowNumber, colIdx) {
  const chunkIdx = Math.floor((rowNumber - 1) / GRID_CHUNK);
  const chunk = gridState.chunks.get(`${sheetIdx}:${chunkIdx}`);
  if (!chunk || chunk === 'loading') return undefined;
  const row = chunk[(rowNumber - 1) % GRID_CHUNK];
  return row ? row[colIdx] : 0;
}

function escapeHtml(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 셀 값 중 검색어와 일치하는 부분만 <mark>로 감싼다 (대소문자 무시) */
function markMatches(rawText, query) {
  if (!query) return escapeHtml(rawText);
  const lower = rawText.toLowerCase();
  const needle = query.toLowerCase();
  const parts = [];
  let from = 0;
  for (;;) {
    const at = lower.indexOf(needle, from);
    if (at === -1) break;
    parts.push(escapeHtml(rawText.slice(from, at)));
    parts.push(`<mark>${escapeHtml(rawText.slice(at, at + needle.length))}</mark>`);
    from = at + needle.length;
  }
  parts.push(escapeHtml(rawText.slice(from)));
  return parts.join('');
}

function renderGridViewport() {
  const sheet = gridState.sheets[gridState.active];
  if (!sheet) return;
  const scroll = document.getElementById('grid-scroll');
  const firstVisible = Math.max(1, Math.floor(scroll.scrollTop / GRID_ROW_H) + 1 - GRID_OVERSCAN);
  const visibleCount = Math.ceil(scroll.clientHeight / GRID_ROW_H) + GRID_OVERSCAN * 2;
  const lastVisible = Math.min(displayRowCount(sheet), firstVisible + visibleCount);
  const colCount = displayColCount(sheet);

  // 필요한 청크를 비동기로 요청 — 실제 데이터 범위 너머(패딩 영역)는 서버가 빈 배열을 돌려준다
  const firstChunk = Math.floor((firstVisible - 1) / GRID_CHUNK);
  const lastChunk = Math.floor((Math.min(lastVisible, sheet.rowCount) - 1) / GRID_CHUNK);
  for (let c = firstChunk; c <= lastChunk; c += 1) fetchChunk(gridState.active, c);

  const cols = `<colgroup><col style="width:${GRID_ROWNUM_W}px">${Array.from({ length: colCount }, (_, i) =>
    `<col style="width:${colWidthAt(sheet, i + 1)}px">`,
  ).join('')}</colgroup>`;

  const rowsHtml = [];
  for (let r = firstVisible; r <= lastVisible; r += 1) {
    const cells = [`<td class="g-rownum">${r}</td>`];
    for (let c = 1; c <= colCount; c += 1) {
      if (gridState.mergeCovered.has(`${r}:${c}`)) continue;
      const merge = gridState.mergeStarts.get(`${r}:${c}`);
      const cell = cellAt(gridState.active, r, c - 1);
      const classes = [];
      if (cell && cell !== 0 && cell.s !== undefined) classes.push(`gs${cell.s}`);
      const hitKey = `${gridState.active}:${r}:${c}`;
      const isHit = gridSearchHits.has(hitKey);
      if (isHit) classes.push('g-hit');
      if (hitKey === gridCurrentHit) classes.push('g-hit-current');
      const value =
        cell && cell !== 0
          ? isHit
            ? markMatches(cell.v, searchState.lastQuery)
            : escapeHtml(cell.v)
          : '';
      const cls = classes.length > 0 ? ` class="${classes.join(' ')}"` : '';
      const span = merge ? ` colspan="${merge.cs}"${merge.rs > 1 ? ` rowspan="${merge.rs}"` : ''}` : '';
      cells.push(`<td${cls}${span}>${value}</td>`);
    }
    rowsHtml.push(`<tr style="height:${GRID_ROW_H}px">${cells.join('')}</tr>`);
  }

  const table = document.getElementById('grid-table');
  table.innerHTML = cols + rowsHtml.join('');
  table.style.transform = `translateY(${(firstVisible - 1) * GRID_ROW_H}px)`;

  pageIndicator.textContent = `${Math.min(sheet.rowCount, Math.floor(scroll.scrollTop / GRID_ROW_H) + 1).toLocaleString()} / ${sheet.rowCount.toLocaleString()}행`;
}

function renderGridHead() {
  const sheet = gridState.sheets[gridState.active];
  const head = document.getElementById('grid-head');
  const scroll = document.getElementById('grid-scroll');
  const cells = [`<div class="g-head-cell g-head-corner" style="width:${GRID_ROWNUM_W}px"></div>`];
  for (let c = 1; c <= displayColCount(sheet); c += 1) {
    cells.push(`<div class="g-head-cell" style="width:${colWidthAt(sheet, c)}px">${colLetter(c)}</div>`);
  }
  head.innerHTML = `<div class="g-head-track">${cells.join('')}</div>`;
  head.firstChild.style.transform = `translateX(${-scroll.scrollLeft}px)`;
}

function activateSheet(index) {
  gridState.active = index;
  const sheet = gridState.sheets[index];
  buildMergeIndex(sheet);
  const inner = document.getElementById('grid-inner');
  inner.style.height = `${displayRowCount(sheet) * GRID_ROW_H}px`;
  let totalWidth = 0;
  for (let c = 1; c <= displayColCount(sheet); c += 1) totalWidth += colWidthAt(sheet, c);
  inner.style.width = `${GRID_ROWNUM_W + totalWidth}px`;
  document.getElementById('grid-scroll').scrollTop = 0;
  for (const [i, tab] of [...document.querySelectorAll('.sheet-tab')].entries()) {
    tab.classList.toggle('active', i === index);
  }
  renderGridHead();
  renderGridViewport();
}

async function initGridView() {
  document.getElementById('pages').hidden = true;
  document.getElementById('grid').hidden = false;
  document.getElementById('zoom-out').hidden = true;
  document.getElementById('zoom-in').hidden = true;
  document.getElementById('zoom-level').hidden = true;
  document.getElementById('invert-btn').hidden = true;
  document.getElementById('measure-btn').hidden = true;
  document.getElementById('markup-toggle-btn').hidden = true;
  document.getElementById('markup-add-btn').hidden = true;
  document.getElementById('markup-include-wrap').hidden = true;

  const meta = await fetchJsonView(`/api/view/grid/meta?${tokenQuery}`);
  gridState.sheets = meta.sheets;
  injectGridStyles(meta.styles);

  const tabs = document.getElementById('sheet-tabs');
  tabs.replaceChildren(
    ...meta.sheets.map((sheet, i) => {
      const tab = document.createElement('button');
      tab.className = 'sheet-tab';
      tab.role = 'tab';
      tab.textContent = sheet.name;
      tab.addEventListener('click', () => activateSheet(i));
      return tab;
    }),
  );

  const scroll = document.getElementById('grid-scroll');
  let ticking = false;
  scroll.addEventListener('scroll', () => {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(() => {
      ticking = false;
      renderGridViewport();
      document.querySelector('.g-head-track').style.transform = `translateX(${-scroll.scrollLeft}px)`;
    });
  });

  // 복사 권한이 없으면 선택·복사·우클릭을 차단한다 (서버가 원본을 주지 않는 것이 근본 방어)
  if (!permissions.canCopy) {
    const grid = document.getElementById('grid');
    grid.classList.add('grid-noselect');
    grid.addEventListener('copy', (event) => event.preventDefault());
  }

  activateSheet(0);
}

/* ==================== 도면 도구: 커서 기준 줌 · 패닝 · 거리 측정 ==================== */

/** 커서가 가리키는 지점을 유지한 채 배율을 바꾼다 (AutoCAD 휠 줌) */
function zoomAtPoint(nextIndex, clientX, clientY) {
  const bounded = Math.max(0, Math.min(ZOOM_STEPS.length - 1, nextIndex));
  if (bounded === zoomIndex) return;
  const ratio = ZOOM_STEPS[bounded] / ZOOM_STEPS[zoomIndex];
  const rect = pagesContainer.getBoundingClientRect();
  const containerDocTop = rect.top + window.scrollY;
  const contentX = pagesContainer.scrollLeft + (clientX - rect.left);
  const contentY = clientY + window.scrollY - containerDocTop;
  zoomIndex = bounded;
  applyZoom();
  pagesContainer.scrollLeft = contentX * ratio - (clientX - rect.left);
  window.scrollTo(window.scrollX, containerDocTop + contentY * ratio - clientY);
}

const measure = {
  mode: false,
  calibrating: false,
  pending: null, // { frame, fx, fy, dot }
  items: [], // { frame, a, b, paperMm, els: [] }
  factor: 1,
  unit: 'mm(용지)',
};

function measureLayer(frame) {
  let layer = frame.querySelector('.measure-layer');
  if (!layer) {
    layer = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    layer.setAttribute('class', 'measure-layer');
    layer.setAttribute('viewBox', '0 0 100 100');
    layer.setAttribute('preserveAspectRatio', 'none');
    frame.append(layer);
  }
  return layer;
}

function measureDot(frame, fx, fy) {
  const dot = document.createElement('div');
  dot.className = 'measure-dot';
  dot.style.left = `${fx * 100}%`;
  dot.style.top = `${fy * 100}%`;
  frame.append(dot);
  return dot;
}

function measurementLabelText(paperMm) {
  const value = paperMm * measure.factor;
  return `${value >= 100 ? value.toFixed(0) : value.toFixed(1)} ${measure.unit}`;
}

function addMeasurement(a, b) {
  const frame = a.frame;
  const pageIdx = [...pagesContainer.children].indexOf(frame);
  const dims = pageDimensions && pageDimensions[pageIdx];
  if (!dims) return;
  const dxPt = (b.fx - a.fx) * dims.w;
  const dyPt = (b.fy - a.fy) * dims.h;
  const paperMm = (Math.hypot(dxPt, dyPt) * 25.4) / 72;

  const layer = measureLayer(frame);
  const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  line.setAttribute('x1', String(a.fx * 100));
  line.setAttribute('y1', String(a.fy * 100));
  line.setAttribute('x2', String(b.fx * 100));
  line.setAttribute('y2', String(b.fy * 100));
  layer.append(line);

  const dotB = measureDot(frame, b.fx, b.fy);
  const label = document.createElement('div');
  label.className = 'measure-label';
  label.style.left = `${((a.fx + b.fx) / 2) * 100}%`;
  label.style.top = `${((a.fy + b.fy) / 2) * 100}%`;
  frame.append(label);

  const item = { frame, paperMm, label, els: [line, a.dot, dotB, label] };
  label.textContent = measurementLabelText(paperMm);
  measure.items.push(item);

  if (measure.calibrating) {
    measure.calibrating = false;
    const input = window.prompt('방금 잰 구간의 실제 길이를 입력하세요 (예: 1500 → mm)', '');
    const value = Number((input || '').replace(/[^\d.]/g, ''));
    if (value > 0) {
      measure.factor = value / paperMm;
      measure.unit = 'mm';
      for (const it of measure.items) it.label.textContent = measurementLabelText(it.paperMm);
      document.getElementById('measure-hint').textContent = '보정 완료 — 실제 치수(mm)로 표시됩니다';
    } else {
      document.getElementById('measure-hint').textContent = '보정이 취소되었습니다';
    }
  }
}

function clearMeasurements() {
  for (const item of measure.items) for (const el of item.els) el?.remove();
  measure.items = [];
  if (measure.pending) {
    measure.pending.dot.remove();
    measure.pending = null;
  }
}

function setMeasureMode(on) {
  measure.mode = on;
  measure.calibrating = false;
  document.getElementById('measure-btn').setAttribute('aria-pressed', String(on));
  document.getElementById('measure-bar').hidden = !on;
  pagesContainer.classList.toggle('measuring', on);
  if (!on) clearMeasurements();
  else document.getElementById('measure-hint').textContent = '두 점을 클릭해 거리를 측정하세요';
}

/* ==================== 마크업(리뷰 코멘트) ==================== */

const markup = {
  visible: true,
  drawMode: false,
  items: [],
  dragging: null, // {frame, page, fx, fy, rectEl, box}
  draftRect: null, // 팝오버가 열려 있는 동안 유지되는 사각형 엘리먼트
};

function markupLayer(frame) {
  let layer = frame.querySelector('.markup-layer');
  if (!layer) {
    layer = document.createElement('div');
    layer.className = 'markup-layer';
    frame.append(layer);
  }
  return layer;
}

function clearMarkupOverlays() {
  for (const frame of pagesContainer.children) {
    const layer = frame.querySelector('.markup-layer');
    if (layer) layer.replaceChildren();
  }
}

async function deleteMarkupItem(item, rectEl) {
  if (!confirm('이 마크업을 삭제할까요?')) return;
  try {
    await fetchJsonView(`/api/view/markup/${item.id}?${tokenQuery}`, { method: 'DELETE' });
    markup.items = markup.items.filter((it) => it.id !== item.id);
    rectEl.remove();
    updateMarkupIncludeVisibility();
  } catch (error) {
    alert(error.message || '삭제에 실패했습니다.');
  }
}

function renderMarkupItem(item) {
  const frame = pagesContainer.children[item.page - 1];
  if (!frame) return;
  const layer = markupLayer(frame);
  const rectEl = document.createElement('div');
  rectEl.className = 'markup-rect';
  rectEl.style.left = `${item.x * 100}%`;
  rectEl.style.top = `${item.y * 100}%`;
  rectEl.style.width = `${item.w * 100}%`;
  rectEl.style.height = `${item.h * 100}%`;
  rectEl.tabIndex = 0;
  rectEl.dataset.markupId = item.id;

  const tooltip = document.createElement('div');
  tooltip.className = 'markup-tooltip';
  const metaLine = document.createElement('div');
  metaLine.className = 'markup-tooltip-meta';
  metaLine.textContent = `${item.author} · ${new Date(item.createdAt).toLocaleString('ko-KR')}`;
  const body = document.createElement('div');
  body.className = 'markup-tooltip-text';
  body.textContent = item.text;
  tooltip.append(metaLine, body);

  if (permissions.canEdit) {
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'markup-tooltip-delete';
    del.textContent = '삭제';
    del.addEventListener('click', (event) => {
      event.stopPropagation();
      deleteMarkupItem(item, rectEl);
    });
    tooltip.append(del);
  }

  rectEl.append(tooltip);
  layer.append(rectEl);
}

function renderAllMarkup() {
  clearMarkupOverlays();
  if (!markup.visible) return;
  for (const item of markup.items) renderMarkupItem(item);
}

function updateMarkupIncludeVisibility() {
  document.getElementById('markup-include-wrap').hidden = markup.items.length === 0;
}

async function loadMarkup() {
  try {
    markup.items = await fetchJsonView(`/api/view/markup?${tokenQuery}`);
    renderAllMarkup();
    updateMarkupIncludeVisibility();
  } catch {
    /* 마크업 로딩 실패는 열람 자체를 막지 않는다 */
  }
}

function setMarkupVisible(visible) {
  markup.visible = visible;
  document.getElementById('markup-toggle-btn').setAttribute('aria-pressed', String(visible));
  renderAllMarkup();
}

function closeMarkupPopover() {
  const existing = document.querySelector('.markup-popover');
  if (existing) existing.remove();
}

function cancelMarkupDraft() {
  if (markup.dragging) {
    markup.dragging.rectEl.remove();
    markup.dragging = null;
  }
  if (markup.draftRect) {
    markup.draftRect.remove();
    markup.draftRect = null;
  }
  closeMarkupPopover();
}

function setMarkupDrawMode(on) {
  markup.drawMode = on;
  document.getElementById('markup-add-btn').setAttribute('aria-pressed', String(on));
  pagesContainer.classList.toggle('markup-drawing', on);
  if (!on) cancelMarkupDraft();
}

function openMarkupPopover(frame, page, box) {
  closeMarkupPopover();
  const popover = document.createElement('div');
  popover.className = 'markup-popover';
  popover.style.left = `${Math.min(box.x + box.w, 0.7) * 100}%`;
  popover.style.top = `${box.y * 100}%`;

  const textarea = document.createElement('textarea');
  textarea.placeholder = '코멘트를 입력하세요';
  textarea.maxLength = 2000;

  const actions = document.createElement('div');
  actions.className = 'markup-popover-actions';
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'btn-pill';
  cancelBtn.textContent = '취소';
  const saveBtn = document.createElement('button');
  saveBtn.type = 'button';
  saveBtn.className = 'btn-pill btn-pill-primary';
  saveBtn.textContent = '저장';
  actions.append(cancelBtn, saveBtn);
  popover.append(textarea, actions);
  frame.append(popover);
  textarea.focus();

  cancelBtn.addEventListener('click', () => cancelMarkupDraft());

  saveBtn.addEventListener('click', async () => {
    const text = textarea.value.trim();
    if (!text) {
      textarea.focus();
      return;
    }
    saveBtn.disabled = true;
    try {
      const item = await fetchJsonView(`/api/view/markup?${tokenQuery}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ page, x: box.x, y: box.y, w: box.w, h: box.h, text }),
      });
      markup.items.push(item);
      cancelMarkupDraft();
      renderAllMarkup();
      updateMarkupIncludeVisibility();
    } catch (error) {
      alert(error.message || '마크업 저장에 실패했습니다.');
      saveBtn.disabled = false;
    }
  });
}

function setupMarkupTool() {
  document.getElementById('markup-toggle-btn').addEventListener('click', () => {
    setMarkupVisible(!markup.visible);
  });

  const addBtn = document.getElementById('markup-add-btn');
  if (permissions.canEdit) {
    addBtn.hidden = false;
    addBtn.addEventListener('click', () => setMarkupDrawMode(!markup.drawMode));
  }

  pagesContainer.addEventListener('mousedown', (event) => {
    if (!markup.drawMode) return;
    const frame = event.target.closest('.page-frame');
    if (!frame) return;
    event.preventDefault();
    const rect = frame.getBoundingClientRect();
    const fx = (event.clientX - rect.left) / rect.width;
    const fy = (event.clientY - rect.top) / rect.height;
    const rectEl = document.createElement('div');
    rectEl.className = 'markup-rect markup-rect-draft';
    markupLayer(frame).append(rectEl);
    markup.dragging = {
      frame,
      page: [...pagesContainer.children].indexOf(frame) + 1,
      fx,
      fy,
      rectEl,
      box: null,
    };
  });

  document.addEventListener('mousemove', (event) => {
    if (!markup.dragging) return;
    const { frame, fx: startFx, fy: startFy, rectEl } = markup.dragging;
    const rect = frame.getBoundingClientRect();
    const curFx = Math.min(Math.max((event.clientX - rect.left) / rect.width, 0), 1);
    const curFy = Math.min(Math.max((event.clientY - rect.top) / rect.height, 0), 1);
    const x = Math.min(startFx, curFx);
    const y = Math.min(startFy, curFy);
    const w = Math.abs(curFx - startFx);
    const h = Math.abs(curFy - startFy);
    rectEl.style.left = `${x * 100}%`;
    rectEl.style.top = `${y * 100}%`;
    rectEl.style.width = `${w * 100}%`;
    rectEl.style.height = `${h * 100}%`;
    markup.dragging.box = { x, y, w, h };
  });

  document.addEventListener('mouseup', () => {
    if (!markup.dragging) return;
    const { frame, page, rectEl, box } = markup.dragging;
    markup.dragging = null;
    const MIN_SIZE = 0.01;
    if (!box || box.w < MIN_SIZE || box.h < MIN_SIZE) {
      rectEl.remove();
      return;
    }
    markup.draftRect = rectEl;
    openMarkupPopover(frame, page, box);
  });
}

function setupDrawingTools() {
  // Ctrl+휠: 커서 기준 확대/축소
  pagesContainer.addEventListener(
    'wheel',
    (event) => {
      if (!event.ctrlKey) return;
      event.preventDefault();
      zoomAtPoint(zoomIndex + (event.deltaY < 0 ? 1 : -1), event.clientX, event.clientY);
    },
    { passive: false },
  );

  // 더블클릭: 해당 지점으로 두 단계 확대 (측정 중에는 비활성)
  pagesContainer.addEventListener('dblclick', (event) => {
    if (measure.mode) return;
    zoomAtPoint(zoomIndex + 2, event.clientX, event.clientY);
  });

  // 드래그 패닝 (왼쪽/휠 버튼) — AutoCAD의 팬처럼 문서를 끌어서 이동
  let pan = null;
  pagesContainer.addEventListener('mousedown', (event) => {
    if (measure.mode) return;
    if (event.button !== 0 && event.button !== 1) return;
    // 컨테이너 스크롤바 영역 클릭은 네이티브 동작(스크롤바 드래그)에 맡긴다
    if (
      event.target === pagesContainer &&
      (event.offsetX > pagesContainer.clientWidth || event.offsetY > pagesContainer.clientHeight)
    ) {
      return;
    }
    pan = {
      x: event.clientX,
      y: event.clientY,
      left: pagesContainer.scrollLeft,
      top: window.scrollY,
    };
    pagesContainer.classList.add('panning');
    event.preventDefault();
  });
  document.addEventListener('mousemove', (event) => {
    if (!pan) return;
    pagesContainer.scrollLeft = pan.left - (event.clientX - pan.x);
    window.scrollTo(window.scrollX, pan.top - (event.clientY - pan.y));
  });
  document.addEventListener('mouseup', () => {
    pan = null;
    pagesContainer.classList.remove('panning');
  });

  // 측정 모드
  document.getElementById('measure-btn').addEventListener('click', () => setMeasureMode(!measure.mode));
  document.getElementById('measure-close').addEventListener('click', () => setMeasureMode(false));
  document.getElementById('measure-clear').addEventListener('click', clearMeasurements);
  document.getElementById('measure-calibrate').addEventListener('click', () => {
    measure.calibrating = true;
    document.getElementById('measure-hint').textContent =
      '실제 길이를 아는 구간(치수선 등)의 양 끝을 클릭하세요';
  });

  pagesContainer.addEventListener('click', (event) => {
    if (!measure.mode) return;
    const frame = event.target.closest('.page-frame');
    if (!frame) return;
    const rect = frame.getBoundingClientRect();
    const point = {
      frame,
      fx: (event.clientX - rect.left) / rect.width,
      fy: (event.clientY - rect.top) / rect.height,
    };
    if (!measure.pending || measure.pending.frame !== frame) {
      if (measure.pending) measure.pending.dot.remove();
      measure.pending = { ...point, dot: measureDot(frame, point.fx, point.fy) };
    } else {
      addMeasurement(measure.pending, point);
      measure.pending = null;
    }
  });
}

async function init() {
  if (!token) {
    showError('열람 토큰이 없습니다. 발급받은 링크로 접속해 주세요.');
    return;
  }
  let res;
  try {
    res = await fetch(`/api/view/info?${tokenQuery}`);
  } catch {
    showError('서버에 연결할 수 없습니다.');
    return;
  }
  const json = await res.json();
  if (!res.ok || !json.success) {
    showError(json.error || '문서를 열 수 없습니다.');
    return;
  }

  docTitle.textContent = json.data.name;
  document.title = `DocView — ${json.data.name}`;
  docName = json.data.name;
  pageCount = json.data.pageCount;
  pageDimensions = json.data.pages || null;
  permissions = json.data.permissions;

  setupToolbar();
  setupDeterrents();
  setupSearchBar();

  if (json.data.viewKind === 'grid') {
    try {
      await initGridView();
    } catch (error) {
      showError(error.message || '스프레드시트를 불러오지 못했습니다.');
    }
    return;
  }

  pageIndicator.textContent = `1 / ${pageCount}`;
  applyZoom();
  buildPages();
  trackCurrentPage();
  setupDrawingTools();
  setupMarkupTool();
  loadMarkup();
}

init();
