import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSampleDocx, makeSamplePptx } from './helpers.js';
import type { SlideEntry } from '../src/collab/pptxShared.js';
import type { PerSlideContentsMap } from '../src/pptx/collab.js';

/**
 * changeLogPersistence.ts — 룸이 "정상적으로 닫히는"(마지막 참여자 퇴장, closeConn의
 * room.conns.size === 0 분기) 바로 그 순간, 그때까지 쌓인 변경 이력을 config.changeLogDir 아래
 * 날짜별 JSONL로 남기는지 검증한다. getCellChangeLog/getParagraphChangeLog/getPptxChangeLog(메모리
 * 전용, 룸이 닫히면 즉시 사라짐)와 달리, 이건 룸이 닫힌 "뒤"에도 디스크에서 확인할 수 있어야 한다는
 * 것이 핵심 계약 — 세 형식(xlsx/docx/pptx) 모두에서 실제 fs.appendFile까지 확인한다(실시간 타이머
 * 없이 실제 시각으로 대기 — fake timer는 libuv poll phase를 타는 디스크 I/O 완료를 안정적으로
 * 기다리지 못한다, 다른 테스트들에서 이미 확인된 제약).
 */
const MESSAGE_SYNC = 0;
const MESSAGE_SLIDE_CONTENT_SUBSCRIBE = 5;
const XML_FRAGMENT_NAME = 'prosemirror';
const PPTX_SLIDES = 'slides';
const PPTX_CONTENTS = 'slideContents';

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();

  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    if (decoding.readVarUint(decoder) === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) readUpdate(decoder, this.mirror, this);
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}

  requestSync(): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeSyncStep1(encoder, this.mirror);
    this.emit('message', encoding.toUint8Array(encoder));
  }

  subscribeSlide(slideEntryId: string): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SLIDE_CONTENT_SUBSCRIBE);
    encoding.writeVarString(encoder, JSON.stringify({ slideEntryId }));
    this.emit('message', encoding.toUint8Array(encoder));
  }

  /** mirror에서 트랜잭션을 실행하고 그 델타만 서버로 보낸다(문서 종류에 관계없이 재사용). */
  edit(action: (doc: Y.Doc) => void): void {
    let captured: Uint8Array | null = null;
    const onUpdate = (update: Uint8Array) => {
      captured = captured === null ? update : Y.mergeUpdates([captured, update]);
    };
    this.mirror.on('update', onUpdate);
    this.mirror.transact(() => action(this.mirror));
    this.mirror.off('update', onUpdate);
    if (!captured) throw new Error('edit produced no update — was requestSync/subscribeSlide called first?');
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, captured);
    this.emit('message', encoding.toUint8Array(encoder));
  }
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('A2').value = 'Alice';
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

function sendCellEdit(socket: FakeSocket, key: string, value: string): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  const scratch = new Y.Doc();
  scratch.getMap<string>('cells').set(key, value);
  writeUpdate(encoder, Y.encodeStateAsUpdate(scratch));
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** fs.appendFile은 closeConn에서 await되지 않는(fire-and-forget) 호출이라, 실제 디스크에 쓰일 때까지
 * 실시간으로 폴링해서 기다린다(가짜 타이머 금지 — 위 파일 주석 참고). */
async function waitFor(check: () => Promise<boolean>, timeoutMs = 4000): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (await check()) return;
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
  throw new Error('waitFor: condition never became true');
}

describe('collab/changeLogPersistence.ts: room-close persistence of change logs', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-changelog-persist-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
    // DOCVIEW_STORAGE=memory는 originals/converted/meta 디렉터리를 건너뛰지만 audit/changeLog
    // 디렉터리는 항상 만든다(storage.ts의 ensureDataDirs) — 실제 서버는 부팅 시 이걸 부르므로
    // persistChangeLog가 그 전제를 그대로 믿고 자기 디렉터리를 스스로 만들지 않는다(audit.ts의
    // recordAudit과 대칭). 여기선 그 부팅 절차를 직접 흉내낸다.
    const { ensureDataDirs } = await import('../src/storage.js');
    await ensureDataDirs();
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  /** 오늘 날짜의 changeLogDir JSONL 파일에서, 주어진 docId의 줄만 파싱해 돌려준다. 여러 테스트가
   * 같은(오늘) 날짜 파일에 함께 append하므로 docId로 걸러야 한다. */
  async function readPersistedEntries(docId: string): Promise<Record<string, unknown>[]> {
    const { config } = await import('../src/config.js');
    let files: string[];
    try {
      files = await fs.readdir(config.changeLogDir);
    } catch {
      return [];
    }
    const out: Record<string, unknown>[] = [];
    for (const file of files) {
      const content = await fs.readFile(path.join(config.changeLogDir, file), 'utf8');
      for (const line of content.split('\n')) {
        if (!line.trim()) continue;
        const parsed = JSON.parse(line) as Record<string, unknown>;
        if (parsed.docId === docId) out.push(parsed);
      }
    }
    return out;
  }

  test('docx: a paragraph edit is written to disk once the last editor leaves', async () => {
    const { store } = await import('../src/storage.js');
    const room = await import('../src/collab/room.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'docx', makeSampleDocx(['First']));
    await store.writeMeta({
      id: docId, name: 'x.docx', ext: 'docx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100,
    });

    const socket = new FakeSocket();
    await room.joinEditSession(socket as never, {} as never, docId, 'docx', [], false, 'viewer-a');
    socket.requestSync();
    socket.edit((doc) => {
      const block = doc.getXmlFragment(XML_FRAGMENT_NAME).toArray()[0] as Y.XmlElement;
      const text = block.toArray().find((c): c is Y.XmlText => c instanceof Y.XmlText)!;
      text.insert(text.length, '!');
    });

    const startedAt = Date.now();
    socket.emit('close'); // 마지막 참여자 퇴장 → 룸이 닫힌다

    await waitFor(async () => (await readPersistedEntries(docId)).length > 0);
    const entries = await readPersistedEntries(docId);
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({
      docId, docKind: 'docx', paragraph: 1, before: 'First', after: 'First!', viewer: 'viewer-a',
    });
    expect(new Date(entries[0].closedAt as string).getTime()).toBeGreaterThanOrEqual(startedAt);

    // 룸은 이제 없다 — 실시간 API는 더 이상 이 변경을 알지 못한다(디스크만이 유일한 기록).
    expect(room.getParagraphChangeLog(docId)).toBeNull();
  });

  test('xlsx: a cell edit is written to disk once the last editor leaves', async () => {
    const { store } = await import('../src/storage.js');
    const room = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'xlsx', await makeSampleXlsx());
    await store.writeMeta({
      id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid',
    });

    const socket = new FakeSocket();
    await room.joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false, 'viewer-b');
    sendCellEdit(socket, gridCellKey(0, 1, 0), 'Bob'); // (row=1,col=0) internal = A2 = "Alice"
    await waitFor(async () => (room.getCellChangeLog(docId)?.length ?? 0) > 0);

    socket.emit('close');

    await waitFor(async () => (await readPersistedEntries(docId)).length > 0);
    const entries = await readPersistedEntries(docId);
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({
      docId, docKind: 'xlsx', sheet: 0, row: 2, col: 1, before: 'Alice', after: 'Bob', viewer: 'viewer-b',
    });

    expect(room.getCellChangeLog(docId)).toBeNull();
  });

  test('pptx: a text box edit is written to disk once the last editor leaves', async () => {
    const { store } = await import('../src/storage.js');
    const room = await import('../src/collab/room.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Slide A']));
    await store.writeMeta({
      id: docId, name: 'x.pptx', ext: 'pptx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100, viewKind: 'pages',
    });

    const socket = new FakeSocket();
    await room.joinEditSession(socket as never, {} as never, docId, 'pptx', [], false, 'viewer-c');
    socket.requestSync();
    const slideId = socket.mirror.getArray<SlideEntry>(PPTX_SLIDES).get(0).id;
    socket.subscribeSlide(slideId);
    await new Promise((resolve) => setTimeout(resolve, 100)); // 지연 시딩(비동기 unzip) 완료를 기다린다

    socket.edit((doc) => {
      const perSlide = doc.getMap<PerSlideContentsMap>(PPTX_CONTENTS).get(slideId)!;
      const content = perSlide.get('2')!.get('content') as Y.XmlFragment;
      const text = (content.toArray()[0] as Y.XmlElement).toArray().find((c): c is Y.XmlText => c instanceof Y.XmlText)!;
      text.insert(text.length, '!');
    });

    socket.emit('close');

    await waitFor(async () => (await readPersistedEntries(docId)).length > 0);
    const entries = await readPersistedEntries(docId);
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({
      // docKind(문서 형식)와 kind(pptx 항목 자신의 종류 — text/slide_added/slide_removed)는 이름이
      // 같지 않아야 한다: 전자가 "pptx", 후자가 "text"로 둘 다 살아남아야 한다(처음엔 이름 충돌로
      // 후자가 전자를 덮어써 "pptx"가 아니라 "text"만 기록되는 버그가 있었다).
      docId, docKind: 'pptx', kind: 'text', slideId, shape: '2', before: 'Slide A', after: 'Slide A!', viewer: 'viewer-c',
    });

    expect(room.getPptxChangeLog(docId)).toBeNull();
  });

  test('an edit-free room leaves nothing on disk', async () => {
    const { store } = await import('../src/storage.js');
    const room = await import('../src/collab/room.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'docx', makeSampleDocx(['First']));
    await store.writeMeta({
      id: docId, name: 'x.docx', ext: 'docx', status: 'ready', pageCount: 1,
      createdAt: new Date().toISOString(), progress: 100,
    });

    const socket = new FakeSocket();
    await room.joinEditSession(socket as never, {} as never, docId, 'docx', [], false, 'viewer-a');
    socket.requestSync(); // 접속·동기화만 하고 편집은 하지 않는다
    socket.emit('close');

    // 음성 검증: "결국 아무것도 안 남는다"는 유한 시간 안에 확인할 수 없으므로, flushToStore와
    // recordAudit처럼 이미 그 시점에 완료되는 다른 부수효과가 끝날 때까지 기다린 뒤 스냅샷을 본다.
    await waitFor(async () => room.getParagraphChangeLog(docId) === null);
    await new Promise((resolve) => setTimeout(resolve, 200));
    expect(await readPersistedEntries(docId)).toEqual([]);
  });
});
