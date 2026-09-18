import { createEditor, type RhwpEditor } from '@rhwp/editor';

interface Permissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
  readonly canEdit: boolean;
}

interface ViewInfo {
  readonly docId: string;
  readonly name: string;
  readonly ext: string;
  readonly permissions: Permissions;
  readonly viewer: string;
}

const params = new URLSearchParams(location.search);
const token = params.get('token') ?? '';

const titleEl = document.getElementById('doc-title') as HTMLElement;
const readonlyBadgeEl = document.getElementById('readonly-badge') as HTMLElement;
const saveBtn = document.getElementById('save-btn') as HTMLButtonElement;
const saveStatusEl = document.getElementById('save-status') as HTMLElement;
const containerEl = document.getElementById('hwp-editor-container') as HTMLElement;
const toolbar = document.getElementById('toolbar') as HTMLElement;
const errorScreen = document.getElementById('error-screen') as HTMLElement;
const errorMessage = document.getElementById('error-message') as HTMLElement;

function showFatalError(message: string): void {
  containerEl.hidden = true;
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

async function fetchSourceBytes(): Promise<ArrayBuffer> {
  const res = await fetch(`/api/view/hwp-source?token=${encodeURIComponent(token)}`);
  if (!res.ok) {
    const json = (await res.json().catch(() => null)) as { error?: string } | null;
    throw new Error(json?.error || '문서를 불러오지 못했습니다.');
  }
  return res.arrayBuffer();
}

/** exportHwp()/exportHwpx() 중 원본 확장자에 맞는 쪽을 고른다 — hwpx로 열었으면 hwpx로,
 * hwp로 열었으면 hwp로 저장해야 확장자와 실제 바이트 포맷이 어긋나지 않는다. */
async function exportBytes(editor: RhwpEditor, ext: string): Promise<Uint8Array> {
  return ext === 'hwpx' ? editor.exportHwpx() : editor.exportHwp();
}

async function saveToServer(editor: RhwpEditor, ext: string): Promise<void> {
  saveStatusEl.textContent = '저장 중…';
  try {
    const bytes = await exportBytes(editor, ext);
    const res = await fetch(`/api/view/hwp-save?token=${encodeURIComponent(token)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream' },
      // exportHwp()/exportHwpx()가 돌려주는 Uint8Array는 실제로는 항상 일반 ArrayBuffer를
      // 감싼 것이지만(WASM 힙에서 막 복사된 값 — SharedArrayBuffer일 수 없다), TS 5.7의 DOM
      // lib 타입이 제네릭 ArrayBufferLike를 BlobPart로 좁히지 못해 캐스트가 필요하다.
      body: new Blob([bytes as BlobPart]),
    });
    const json = (await res.json()) as { success: boolean; error?: string };
    if (!res.ok || !json.success) {
      throw new Error(json.error || '저장에 실패했습니다.');
    }
    // 업로드가 성공한 뒤에만 통지한다 — 실패 시 통지하지 않으면 스튜디오가 자동복구 draft를
    // 보존해 재시도/복구가 가능하다(@rhwp/editor README "저장 계약" 참고).
    await editor.notifySaved();
    saveStatusEl.textContent = `저장됨 ${new Date().toLocaleTimeString('ko-KR')}`;
  } catch (error) {
    saveStatusEl.textContent = '저장 실패';
    console.error('hwp 저장 실패:', error);
    window.alert(error instanceof Error ? error.message : '저장에 실패했습니다.');
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

  if (info.ext !== 'hwp' && info.ext !== 'hwpx') {
    showFatalError('한글(hwp/hwpx) 문서가 아닙니다.');
    return;
  }

  titleEl.textContent = info.name;
  document.title = `${info.name} — DocView`;

  const canEdit = info.permissions.canEdit;
  readonlyBadgeEl.hidden = canEdit;
  saveBtn.hidden = !canEdit;

  let sourceBytes: ArrayBuffer;
  try {
    sourceBytes = await fetchSourceBytes();
  } catch (error) {
    showFatalError(error instanceof Error ? error.message : '문서를 불러오지 못했습니다.');
    return;
  }

  const editor = await createEditor(containerEl, { studioUrl: '/rhwp/' });
  await editor.loadFile(sourceBytes, info.name);

  if (canEdit) {
    saveBtn.addEventListener('click', () => {
      void saveToServer(editor, info.ext);
    });
  }
}

main().catch((error) => {
  console.error('hwp 편집기 초기화 실패:', error);
  showFatalError('편집기를 초기화하지 못했습니다.');
});
