'use strict';

const DEFAULT_LABEL = '파일을 선택하면 바로 업로드되고 뷰어가 열립니다';

const fileInput = document.getElementById('file-input');
const fileLabel = document.getElementById('file-label');
const uploadStatus = document.getElementById('upload-status');
const docList = document.getElementById('doc-list');
const emptyMessage = document.getElementById('empty-message');

function setStatus(message, kind) {
  uploadStatus.textContent = message;
  uploadStatus.dataset.kind = kind;
}

async function fetchJson(url, options) {
  const res = await fetch(url, options);
  const json = await res.json();
  if (!res.ok || !json.success) {
    throw new Error(json.error || `요청 실패 (HTTP ${res.status})`);
  }
  return json.data;
}

async function loadFormats() {
  try {
    const formats = await fetchJson('/api/formats');
    fileInput.accept = formats.map((ext) => `.${ext}`).join(',');
    fileLabel.textContent = DEFAULT_LABEL;
    document.getElementById('format-hint').textContent = `지원 형식: ${formats.join(', ')}`;
  } catch {
    /* 형식 목록을 못 가져와도 업로드 자체는 서버가 검증한다 */
  }
}

async function issueToken(docId, viewer, canPrint, canDownload, canCopy, canEdit) {
  return fetchJson(`/api/documents/${docId}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer, canPrint, canDownload, canCopy, canEdit: canEdit === true }),
  });
}

async function deleteDocument(docId) {
  return fetchJson(`/api/documents/${docId}`, { method: 'DELETE' });
}

/** 업로더 본인 열람용 — 전체 권한 토큰을 발급해 이동한다. 항상 viewerUrl(업로드 시점에 미리
 * 만들어 둔 렌더링 캐시를 그냥 HTTP로 읽는 가벼운 화면)로 연다 — xlsx/pptx도 마찬가지다.
 * 예전에는 xlsx/pptx의 "바로 보기"도 편집 화면(editor-grid.html/editor-slides.html)을 그대로
 * 재사용했는데(server가 canEdit 여부와 무관하게 항상 editorUrl을 내려주므로 가능은 했다), 그
 * 화면은 열람뿐이어도 Yjs Y.Doc·WebSocket·LazyGridStore 전체 스트리밍 파싱·전체 시트 서식
 * 동기화까지 실제 동시편집과 똑같이 다 거쳐서 열람 속도가 불필요하게 느렸다 — 편집 툴바·수식
 * 입력창 등은 어차피 안 보이니 얻는 것 없이 비용만 지불하고 있었다. 동시편집이 실제로 필요한
 * 경우는 "동시편집 열기" 버튼(openEditor)의 몫이다.
 * docx/xlsx/pptx가 아닌 문서(PDF 등 pages 뷰)는 동시편집 대상이 아니므로 canEdit이 곧 마크업
 * (리뷰 코멘트) 추가/삭제 권한이 된다 — 업로더 본인이 보는 화면이니 기본으로 켜서 발급한다. */
async function openViewer(docId, viewer, ext) {
  const canEdit = !supportsEdit({ ext });
  const issued = await issueToken(docId, viewer, true, true, true, canEdit);
  location.href = issued.viewerUrl;
}

/** 업로더 본인 편집용 — 전체 권한 + canEdit 토큰을 발급해 동시편집기로 이동한다(docx/xlsx). */
async function openEditor(docId, viewer) {
  const issued = await issueToken(docId, viewer, true, true, true, true);
  if (!issued.editorUrl) {
    throw new Error('이 문서는 동시편집을 지원하지 않습니다 (docx/xlsx만 지원).');
  }
  location.href = issued.editorUrl;
}

const progressBox = document.getElementById('progress');
const progressLabel = document.getElementById('progress-label');
const progressPercent = document.getElementById('progress-percent');
const progressFill = document.getElementById('progress-fill');

function showProgress(label, ratio) {
  progressBox.hidden = false;
  progressLabel.textContent = label;
  if (ratio === null) {
    progressFill.classList.add('indeterminate');
    progressFill.style.transform = '';
    progressPercent.textContent = '';
  } else {
    progressFill.classList.remove('indeterminate');
    progressFill.style.transform = `scaleX(${ratio})`;
    progressPercent.textContent = `${Math.round(ratio * 100)}%`;
  }
}

function hideProgress() {
  progressBox.hidden = true;
  progressFill.classList.remove('indeterminate');
  progressFill.style.transform = 'scaleX(0)';
}

/** XHR로 업로드해 전송 진행률(%)을 노출한다. fetch는 업로드 진행 이벤트를 지원하지 않는다. */
function uploadWithProgress(file, onProgress) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/api/documents');
    xhr.upload.addEventListener('progress', (event) => {
      if (event.lengthComputable) onProgress(event.loaded / event.total);
    });
    xhr.addEventListener('load', () => {
      try {
        const json = JSON.parse(xhr.responseText);
        if (xhr.status >= 200 && xhr.status < 300 && json.success) {
          resolve(json.data);
        } else {
          reject(new Error(json.error || `업로드 실패 (HTTP ${xhr.status})`));
        }
      } catch {
        reject(new Error(`업로드 응답을 해석할 수 없습니다 (HTTP ${xhr.status})`));
      }
    });
    xhr.addEventListener('error', () => reject(new Error('네트워크 오류로 업로드에 실패했습니다.')));
    const body = new FormData();
    body.append('file', file);
    xhr.send(body);
  });
}

const POLL_INTERVAL_MS = 700;
const POLL_TIMEOUT_MS = 5 * 60 * 1000;
/* LibreOffice 변환 구간(25~85%)은 서버가 중간값을 못 주므로 시간 기반으로 게이지를 부드럽게 채운다 */
const CONVERT_CREEP_CAP = 84;
const CONVERT_CREEP_RATE = 0.07;

async function waitForConversion(docId) {
  const startedAt = Date.now();
  let displayed = 10;
  for (;;) {
    const meta = await fetchJson(`/api/documents/${docId}`);
    if (meta.status === 'ready') return meta;
    if (meta.status === 'failed') throw new Error(meta.error || '문서 변환에 실패했습니다.');
    if (Date.now() - startedAt > POLL_TIMEOUT_MS) {
      throw new Error('변환이 너무 오래 걸립니다. 잠시 후 문서 목록에서 다시 확인해 주세요.');
    }
    displayed = Math.max(displayed, meta.progress ?? 10);
    if (displayed < CONVERT_CREEP_CAP) {
      displayed = Math.min(CONVERT_CREEP_CAP, displayed + (CONVERT_CREEP_CAP - displayed) * CONVERT_CREEP_RATE);
    }
    showProgress('업로드 중…', displayed / 100);
    await new Promise((resolve) => setTimeout(resolve, POLL_INTERVAL_MS));
  }
}

fileInput.addEventListener('change', async () => {
  const file = fileInput.files[0];
  if (!file) return;
  fileLabel.textContent = file.name;
  fileInput.disabled = true;
  setStatus('', 'pending');
  try {
    showProgress('업로드 중…', 0);
    const doc = await uploadWithProgress(file, (ratio) => showProgress('업로드 중…', ratio));
    showProgress('업로드 중…', 0.1);
    const ready = await waitForConversion(doc.id);
    // 업로드는 목록에만 반영한다 — 열람/편집은 목록에서 사용자가 직접 "바로보기"/"동시편집"을
    // 눌러야 시작된다(자동으로 편집기·뷰어로 이동하지 않는다).
    showProgress('업로드 완료', 1);
    setStatus(`${ready.pageCount}페이지 업로드 완료`, 'ok');
    hideProgress();
    fileInput.disabled = false;
    fileInput.value = '';
    fileLabel.textContent = DEFAULT_LABEL;
    await refreshList().catch(() => {});
  } catch (error) {
    hideProgress();
    setStatus(error.message, 'error');
    fileInput.disabled = false;
    fileInput.value = '';
    fileLabel.textContent = DEFAULT_LABEL;
    await refreshList().catch(() => {});
  }
});

const LIST_POLL_INTERVAL_MS = 2000;
let listPollTimer = null;

async function refreshList() {
  const documents = await fetchJson('/api/documents');
  // 토큰 발급 폼에 입력 중이면 이번 갱신은 건너뛴다 (입력값이 지워지지 않도록)
  if (!docList.contains(document.activeElement)) {
    docList.replaceChildren(...documents.map(renderDocItem));
    emptyMessage.hidden = documents.length > 0;
  }
  // 변환 중인 문서가 있으면 완료될 때까지 목록을 주기적으로 갱신한다
  const hasProcessing = documents.some((doc) => doc.status === 'processing');
  if (hasProcessing && listPollTimer === null) {
    listPollTimer = setTimeout(() => {
      listPollTimer = null;
      refreshList().catch(() => {});
    }, LIST_POLL_INTERVAL_MS);
  }
}

/** docx/xlsx/pptx가 동시편집을 지원한다 (src/routes/documents.ts의 EDITABLE_EXTS와 일치) */
function supportsEdit(doc) {
  return doc.ext === 'docx' || doc.ext === 'xlsx' || doc.ext === 'pptx';
}

/** 동시편집 대상이 아닌 pages 뷰 문서(PDF 등)는 canEdit이 마크업(리뷰 코멘트) 추가/삭제 권한이 된다. */
function supportsMarkup(doc) {
  return !supportsEdit(doc) && (doc.viewKind ?? 'pages') === 'pages';
}

function renderDocItem(doc) {
  const li = document.createElement('li');
  li.className = 'doc-item';

  const info = document.createElement('div');
  info.className = 'doc-info';
  const name = document.createElement('strong');
  name.textContent = doc.name;
  const detail = document.createElement('span');
  detail.className = 'doc-detail';
  const when = new Date(doc.createdAt).toLocaleString('ko-KR');
  if (doc.status === 'ready') {
    detail.textContent = `${doc.pageCount}페이지 · ${when}`;
  } else if (doc.status === 'processing') {
    detail.textContent = `업로드 중 ${doc.progress ?? 0}% · ${when}`;
  } else {
    detail.textContent = `변환 실패 · ${when}`;
  }
  info.append(name, detail);

  function makeDeleteButton() {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'btn btn-danger';
    button.textContent = '삭제';
    button.addEventListener('click', async () => {
      if (!confirm(`"${doc.name}" 문서를 삭제할까요? 이 작업은 되돌릴 수 없습니다.`)) return;
      // 삭제 요청이 끝나기 전부터 즉시 반응이 보이도록, 버튼만이 아니라 카드 전체를 바로 흐리게
      // 한다 — 예전에는 DELETE 요청과 그 뒤 목록 재조회(refreshList, 별도의 GET 왕복) 두 번의
      // 네트워크 왕복이 끝나야만 화면이 바뀌어 "눌러도 한참 반응이 없다"고 느껴졌다.
      button.disabled = true;
      li.classList.add('is-deleting');
      try {
        await deleteDocument(doc.id);
        // 성공하면 전체 목록을 다시 받아오는 대신 이 카드만 바로 지운다 — 왕복 한 번을 아낀다.
        li.remove();
        emptyMessage.hidden = docList.children.length > 0;
      } catch (error) {
        li.classList.remove('is-deleting');
        button.disabled = false;
        setStatus(error.message, 'error');
      }
    });
    return button;
  }

  if (doc.status !== 'ready') {
    info.append(makeDeleteButton());
    li.append(info);
    if (doc.status === 'processing') {
      const gauge = document.createElement('div');
      gauge.className = 'progress-track doc-gauge';
      const fill = document.createElement('div');
      fill.className = 'progress-fill';
      fill.style.transform = `scaleX(${(doc.progress ?? 0) / 100})`;
      gauge.append(fill);
      li.append(gauge);
    }
    if (doc.error) {
      const errorLine = document.createElement('p');
      errorLine.className = 'token-result';
      errorLine.textContent = doc.error;
      li.append(errorLine);
    }
    return li;
  }

  const quickView = document.createElement('button');
  quickView.type = 'button';
  quickView.className = 'btn';
  quickView.textContent = '바로 보기';
  quickView.addEventListener('click', () => {
    quickView.disabled = true;
    openViewer(doc.id, '업로더', doc.ext).catch((error) => {
      quickView.disabled = false;
      setStatus(error.message, 'error');
    });
  });
  info.append(quickView);

  if (supportsEdit(doc)) {
    const quickEdit = document.createElement('button');
    quickEdit.type = 'button';
    quickEdit.className = 'btn';
    quickEdit.textContent = '동시편집 열기';
    quickEdit.addEventListener('click', () => {
      quickEdit.disabled = true;
      openEditor(doc.id, '업로더').catch((error) => {
        quickEdit.disabled = false;
        setStatus(error.message, 'error');
      });
    });
    info.append(quickEdit);
  }

  info.append(makeDeleteButton());

  const form = document.createElement('form');
  form.className = 'token-form';
  form.innerHTML = `
    <input type="text" name="viewer" placeholder="열람자 이름" required maxlength="100" />
    <label><input type="checkbox" name="canPrint" /> 출력</label>
    <label><input type="checkbox" name="canDownload" /> 저장</label>
    <label><input type="checkbox" name="canCopy" /> 복사</label>
    ${supportsEdit(doc) ? '<label><input type="checkbox" name="canEdit" /> 동시편집</label>' : ''}
    ${supportsMarkup(doc) ? '<label><input type="checkbox" name="canEdit" /> 마크업 편집(코멘트 추가/삭제)</label>' : ''}
    <button type="submit" class="btn btn-primary">열람 링크 발급</button>
  `;
  const result = document.createElement('p');
  result.className = 'token-result';

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const data = new FormData(form);
    try {
      const issued = await issueToken(
        doc.id,
        data.get('viewer'),
        data.get('canPrint') === 'on',
        data.get('canDownload') === 'on',
        data.get('canCopy') === 'on',
        data.get('canEdit') === 'on',
      );
      result.replaceChildren();
      const open = document.createElement('a');
      open.href = issued.viewerUrl;
      open.target = '_blank';
      open.rel = 'noopener';
      open.textContent = '뷰어 열기';
      open.className = 'btn';
      result.append(open);
      if (issued.editorUrl) {
        const openEdit = document.createElement('a');
        openEdit.href = issued.editorUrl;
        openEdit.target = '_blank';
        openEdit.rel = 'noopener';
        openEdit.textContent = '편집기 열기';
        openEdit.className = 'btn';
        result.append(openEdit);
      }
      const copy = document.createElement('button');
      copy.type = 'button';
      copy.textContent = '링크 복사';
      copy.className = 'btn btn-ghost';
      copy.addEventListener('click', async () => {
        const link = issued.editorUrl || issued.viewerUrl;
        await navigator.clipboard.writeText(new URL(link, location.origin).href);
        copy.textContent = '복사됨';
      });
      result.append(copy);
    } catch (error) {
      result.textContent = error.message;
    }
  });

  li.append(info, form, result);
  return li;
}

loadFormats();
refreshList().catch((error) => setStatus(`목록을 불러오지 못했습니다: ${error.message}`, 'error'));
