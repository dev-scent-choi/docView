import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep1, messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSampleDocx } from './helpers.js';

/**
 * docx의 동시편집 화면(editor.ts/editor.html)이 xlsx/pptx처럼 canEdit:false에서도 읽기 전용으로
 * 열리도록 확장하면서, 클라이언트가 편집 UI를 숨기는 것과는 **별개로** 서버가 실제 쓰기를 직접
 * 막는지를 검증한다.
 *
 * src/routes/edit.ts는 파일 확장자와 무관하게 `readOnly = !claims.perms.canEdit`를 계산해
 * joinEditSession에 넘기고, src/collab/room.ts의 handleMessage(MESSAGE_SYNC 분기)는
 * room.readOnlyConns.has(conn)이면 syncStep2/update(실제 쓰기)만 조용히 버리고 syncStep1(읽기
 * 요청)은 그대로 처리한다 — 이 로직은 ext별 분기 없이 xlsx/pptx/docx가 전부 같은 코드를 탄다.
 * xlsx는 roomRowColShift.test.ts가 (MESSAGE_GRID_STRUCTURE_OP 경로로) 이미 다루고 있으므로,
 * 여기서는 **모든 문서 형식이 공유하는 더 아래 계층**인 MESSAGE_SYNC/messageYjsUpdate 경로 자체를
 * docx로 직접 검증한다 — 악의적/버그 있는 클라이언트가 UI를 거치지 않고 원시 Yjs update 메시지를
 * 소켓에 직접 흘려보내는 상황을 그대로 흉내낸다.
 */

const MESSAGE_SYNC = 0;
const XML_FRAGMENT_NAME = 'prosemirror';

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();

  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    const messageType = decoding.readVarUint(decoder);
    if (messageType === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep1) {
        // 서버가 먼저 보낸 자기 상태 벡터 — 이 테스트는 그 자리에서 곧바로 회신하지 않고
        // requestSync()로 명시적으로 요청한다.
      } else if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) {
        readUpdate(decoder, this.mirror, this);
      }
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}

  /** 실제 클라이언트가 연결 직후 하듯 자기(빈) 상태 벡터를 보내 서버가 가진 내용을 회신받는다. */
  requestSync(): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeSyncStep1(encoder, this.mirror);
    this.emit('message', encoding.toUint8Array(encoder));
  }
}

/** 별도의(공격자가 조작한) Y.Doc에서 만든 전체 상태 업데이트를, 실제 y-websocket 클라이언트가
 * 보내는 것과 완전히 같은 MESSAGE_SYNC/messageYjsUpdate 프레임으로 소켓에 직접 흘려보낸다 —
 * editor.ts의 ProseMirror/yUndoPlugin을 전혀 거치지 않는, 순수 프로토콜 레벨 쓰기 시도다. */
function sendRawUpdate(socket: FakeSocket, text: string): void {
  const scratch = new Y.Doc();
  const fragment = scratch.getXmlFragment(XML_FRAGMENT_NAME);
  const paragraph = new Y.XmlElement('paragraph');
  const textNode = new Y.XmlText();
  textNode.insert(0, text);
  paragraph.insert(0, [textNode]);
  fragment.insert(0, [paragraph]);
  const update = Y.encodeStateAsUpdate(scratch);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, update);
  socket.emit('message', encoding.toUint8Array(encoder));
}

async function flushMicrotasks(times = 50): Promise<void> {
  for (let i = 0; i < times; i += 1) await Promise.resolve();
}

async function setUpTestDoc(): Promise<{ docId: string }> {
  const { store } = await import('../src/storage.js');
  const docId = randomUUID();
  const original = makeSampleDocx(['Hello']);
  await store.saveOriginal(docId, 'docx', original);
  await store.writeMeta({
    id: docId,
    name: 'x.docx',
    ext: 'docx',
    status: 'ready',
    pageCount: 1,
    createdAt: new Date().toISOString(),
    progress: 100,
  });
  return { docId };
}

/** getCurrentBytes()가 돌려주는 canonical docx 바이트에서 본문 텍스트(XML)를 꺼낸다 — 실제로
 * 저장/다운로드될 내용을 기준으로 검증하기 위해(화면상의 Yjs 미러만 보는 것보다 신뢰도가 높다). */
function documentXmlOf(bytes: Uint8Array): string {
  const entries = unzipSync(bytes);
  return strFromU8(entries['word/document.xml']);
}

describe('room.ts: a read-only docx connection cannot write, independent of the client UI', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-docx-readonly-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('a read-only connection still receives the full document (syncStep1 is not blocked)', async () => {
    const { joinEditSession } = await import('../src/collab/room.js');
    const { docId } = await setUpTestDoc();

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'docx', [], true);
    socket.requestSync();
    await flushMicrotasks();

    const mirrorFragment = socket.mirror.getXmlFragment(XML_FRAGMENT_NAME);
    expect(mirrorFragment.toString()).toContain('Hello');
  });

  test('a raw Yjs update from a read-only connection is dropped — canonical content on disk/export is unchanged, and no other connection ever sees it', async () => {
    const { joinEditSession, getCurrentBytes } = await import('../src/collab/room.js');
    const { docId } = await setUpTestDoc();

    const attacker = new FakeSocket();
    await joinEditSession(attacker as never, {} as never, docId, 'docx', [], true);
    attacker.requestSync();
    await flushMicrotasks();

    // 클라이언트 UI(editor.ts)를 완전히 우회해 원시 update 프레임을 직접 보낸다 — 악의적이거나
    // 버그 있는 클라이언트를 흉내낸다.
    sendRawUpdate(attacker, 'INJECTED-BY-READONLY-CLIENT');
    await flushMicrotasks();

    const exported = await getCurrentBytes(docId);
    expect(exported).not.toBeNull();
    const xml = documentXmlOf(exported!);
    expect(xml).not.toContain('INJECTED-BY-READONLY-CLIENT');
    expect(xml).toContain('Hello');

    // 공격자 자신의 미러에 로컬로 반영된 것과는 별개로, canonical room.ydoc 자체가 바뀌지
    // 않았어야 한다 — 뒤늦게 들어온(정상 권한) 다른 연결도 그 주입된 내용을 절대 보면 안 된다.
    const observer = new FakeSocket();
    await joinEditSession(observer as never, {} as never, docId, 'docx', [], false);
    observer.requestSync();
    await flushMicrotasks();
    const observerXml = observer.mirror.getXmlFragment(XML_FRAGMENT_NAME).toString();
    expect(observerXml).not.toContain('INJECTED-BY-READONLY-CLIENT');
    expect(observerXml).toContain('Hello');
  });

  test('sanity check: the same raw update IS applied for a normal (non-read-only) connection', async () => {
    // 위 거부 테스트가 실제로 무언가를 검증하고 있는지(메커니즘 자체가 죽어서 항상 "안 바뀜"으로
    // 보이는 거짓 양성이 아닌지) 확인하는 대조군 — canEdit:true 연결에서는 같은 시도가 그대로 반영돼야 한다.
    const { joinEditSession, getCurrentBytes } = await import('../src/collab/room.js');
    const { docId } = await setUpTestDoc();

    const editor = new FakeSocket();
    await joinEditSession(editor as never, {} as never, docId, 'docx', [], false);
    editor.requestSync();
    await flushMicrotasks();

    sendRawUpdate(editor, 'APPLIED-BY-EDITOR-CLIENT');
    await flushMicrotasks();

    const exported = await getCurrentBytes(docId);
    expect(exported).not.toBeNull();
    const xml = documentXmlOf(exported!);
    expect(xml).toContain('APPLIED-BY-EDITOR-CLIENT');
  });
});
