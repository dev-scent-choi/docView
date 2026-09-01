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

/** 업로더 본인 열람용 — 전체 권한(수정 제외) 토큰을 발급해 이동한다.
 * xlsx는 서버가 canEdit 여부와 무관하게 항상 editorUrl을 내려주므로(collab/room.ts가 쓰기만
 * 서버 쪽에서 막는다), "바로 보기"도 편집 화면과 완전히 같은 그리드 UI로 열되 수정만 못 하게
 * 한다 — 읽기 전용이라고 다른 화면(viewer.html)을 따로 보여줄 이유가 없다. docx는 editor.ts가
 * 아직 읽기 전용 모드를 지원하지 않아(server가 canEdit:false면 editorUrl 자체를 안 줌) 기존대로
 * viewer.html을 쓴다. */
async function openViewer(docId, viewer, ext) {
  const issued = await issueToken(docId, viewer, true, true, true, false);
  location.href = ext === 'xlsx' && issued.editorUrl ? issued.editorUrl : issued.viewerUrl;
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
    progressFill.style.width = '';
    progressPercent.textContent = '';
  } else {
    progressFill.classList.remove('indeterminate');
    progressFill.style.width = `${Math.round(ratio * 100)}%`;
    progressPercent.textContent = `${Math.round(ratio * 100)}%`;
  }
}

function hideProgress() {
  progressBox.hidden = true;
  progressFill.classList.remove('indeterminate');
  progressFill.style.width = '0%';
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
    showProgress('변환 중…', displayed / 100);
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
    showProgress('변환 중…', 0.1);
    const ready = await waitForConversion(doc.id);
    // docx/xlsx처럼 동시편집을 지원하는 형식은 업로드 직후 바로 편집기로 들어간다 — 굳이
    // 읽기 전용 뷰어를 거쳤다가 다시 "편집" 버튼을 눌러야 하는 두 단계를 거칠 이유가 없다.
    const canOpenEditor = supportsEdit(ready);
    showProgress(canOpenEditor ? '변환 완료 — 편집기로 이동합니다' : '변환 완료 — 뷰어로 이동합니다', 1);
    setStatus(`${ready.pageCount}페이지 변환 완료`, 'ok');
    if (canOpenEditor) {
      await openEditor(ready.id, '업로더');
    } else {
      await openViewer(ready.id, '업로더');
    }
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

/** docx/xlsx가 동시편집을 지원한다 (src/routes/documents.ts의 EDITABLE_EXTS와 일치) */
function supportsEdit(doc) {
  return doc.ext === 'docx' || doc.ext === 'xlsx';
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
    detail.textContent = `변환 중 ${doc.progress ?? 0}% · ${when}`;
  } else {
    detail.textContent = `변환 실패 · ${when}`;
  }
  info.append(name, detail);

  if (doc.status !== 'ready') {
    li.append(info);
    if (doc.status === 'processing') {
      const gauge = document.createElement('div');
      gauge.className = 'progress-track doc-gauge';
      const fill = document.createElement('div');
      fill.className = 'progress-fill';
      fill.style.width = `${doc.progress ?? 0}%`;
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

  const form = document.createElement('form');
  form.className = 'token-form';
  form.innerHTML = `
    <input type="text" name="viewer" placeholder="열람자 이름" required maxlength="100" />
    <label><input type="checkbox" name="canPrint" /> 출력</label>
    <label><input type="checkbox" name="canDownload" /> 저장</label>
    <label><input type="checkbox" name="canCopy" /> 복사</label>
    ${supportsEdit(doc) ? '<label><input type="checkbox" name="canEdit" /> 동시편집</label>' : ''}
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
