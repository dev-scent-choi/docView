import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep1, messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSampleDocx } from './helpers.js';

const MESSAGE_SYNC = 0;
const XML_FRAGMENT_NAME = 'prosemirror';

/**
 * fake timer 한 번을 진행시키면 그 콜백(flushToStore)이 시작되지만, 그 안의 await 체인
 * (readMeta → readOriginal → saveOriginal → recordAudit → docMeta 쓰기 → 브로드캐스트)은
 * 타이머가 아닌 일반 프라미스라 advanceTimersByTimeAsync가 시간만 감아서는 끝까지 다 풀리지
 * 않을 수 있다 — 남은 마이크로태스크 큐를 몇 번 더 비워 체인이 완전히 끝난 뒤 상태를 읽는다.
 */
async function flushMicrotasks(times = 50): Promise<void> {
  for (let i = 0; i < times; i += 1) await Promise.resolve();
}

/**
 * joinEditSession이 필요로 하는 최소한의 WebSocket 인터페이스만 흉내 낸다 — 실제 네트워크나
 * 인증(routes/edit.ts) 없이 room.ts의 룸 생성·메시지 처리 로직만 직접 거친다. 서버가 보내는
 * sync 메시지를 그대로 mirror(별도 Y.Doc)에 적용해, 실제 y-websocket 클라이언트가 보는 것과
 * 같은 상태(docMeta 포함)를 테스트에서 그대로 읽을 수 있게 한다.
 */
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
        // 서버가 먼저 보낸 자기 상태 벡터 — mirror가 회신할 syncStep2는(mirror가 아직 아무것도
        // 안 만들었으므로) 이 테스트에서 의미 있는 내용이 없다. 조용히 무시한다.
      } else if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) {
        readUpdate(decoder, this.mirror, this);
      }
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}

  /** 실제 클라이언트가 연결 직후 하듯 자기(빈) 상태 벡터를 보내 서버가 가진 내용을 받는다 */
  requestSync(): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeSyncStep1(encoder, this.mirror);
    this.emit('message', encoding.toUint8Array(encoder));
  }
}

function sendEdit(socket: FakeSocket, text: string): void {
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

/**
 * room.ts의 flushToStore가 끝나며 스스로 기록하는 저장 시각(docMeta)이 다시 자동저장을
 * 예약해 버리면, 실제 사용자 편집이 하나도 없어도 8초마다 영원히 재저장(및 PDF 재변환)이
 * 반복된다 — 이 세션에서 실제로 겪은 버그. 진짜 편집 한 번 → 저장 한 번만 일어나고, 그
 * 이후로는(추가 편집 없이 시간만 흘러도) 더 이상 저장이 반복되지 않아야 한다.
 */
describe('room.ts: a real edit triggers exactly one autosave, never a perpetual loop', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-loop-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    // PDF 변환기는 이 테스트의 관심사가 아니다(LibreOffice 등 외부 도구 없이도 돌게) — findConverter가
    // null을 돌려주게 해 flushToStore가 변환 단계를 건너뛰게 한다.
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    // recordAudit는 실제 디스크 I/O(fs.appendFile)를 한다 — 그 완료 콜백은 libuv 폴 단계를 타서
    // 순수 마이크로태스크 반복(flushMicrotasks)이나 advanceTimersByTimeAsync로는 안정적으로
    // 기다려지지 않는다(가짜 타이머가 실제 이벤트 루프 I/O 단계를 앞당겨주지 않기 때문). 이
    // 테스트의 관심사가 아니므로 동기적으로 즉시 끝나는 걸로 교체해 체인 전체를 마이크로태스크만
    // 으로 완주 가능하게 만든다.
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('autosave does not perpetually re-fire after its own save-status write', async () => {
    vi.useFakeTimers();

    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const { docId } = await setUpTestDoc();

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'docx', [], false);

    const saveSpy = vi.spyOn(store, 'saveOriginal');

    // 실제 원격 편집 한 번을 시뮬레이션한다 — 별도의 독립된 Y.Doc에서 만든 업데이트를 서버
    // 소켓의 'message' 이벤트로 그대로 보낸다(CRDT라 서로 다른 Y.Doc 간 업데이트도 그대로 병합됨).
    sendEdit(socket, 'edited concurrently');

    // 8초 유휴 자동저장이 한 번 돈다.
    await vi.advanceTimersByTimeAsync(8_100);
    await flushMicrotasks();
    expect(saveSpy).toHaveBeenCalledTimes(1);

    // 그 저장 자신이 기록한 저장 시각(docMeta) 쓰기가 또 다른 자동저장을 예약했다면, 추가
    // 편집이 전혀 없어도 이 시점에 saveOriginal이 다시 불렸을 것이다 — 그러면 버그(무한 루프)다.
    await vi.advanceTimersByTimeAsync(30_000);
    await flushMicrotasks();
    expect(saveSpy).toHaveBeenCalledTimes(1);
  });
});

/**
 * 저장이 실패하면(디스크 오류, 되풀이되는 내보내기 실패 등) 예전에는 서버 콘솔에만 로그를
 * 남기고 클라이언트에는 아무 신호도 가지 않아 "저장 중…" 표시가 영원히 멈춰 보였다 —
 * queueFlush의 실패 처리가 docMeta.saveErrorAt을 기록해 다른 참여자 화면까지 실시간으로
 * "저장 실패"를 알 수 있어야 한다.
 */
describe('room.ts: a failed save is surfaced to clients via docMeta', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-saveerror-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    // recordAudit는 실제 디스크 I/O(fs.appendFile)를 한다 — 그 완료 콜백은 libuv 폴 단계를 타서
    // 순수 마이크로태스크 반복(flushMicrotasks)이나 advanceTimersByTimeAsync로는 안정적으로
    // 기다려지지 않는다(가짜 타이머가 실제 이벤트 루프 I/O 단계를 앞당겨주지 않기 때문). 이
    // 테스트의 관심사가 아니므로 동기적으로 즉시 끝나는 걸로 교체해 체인 전체를 마이크로태스크만
    // 으로 완주 가능하게 만든다.
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('a save failure sets saveErrorAt (newer than savedAt) without looping', async () => {
    vi.useFakeTimers();

    const { joinEditSession } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');
    const { DOC_META_MAP, DOC_META_SAVED_AT_KEY, DOC_META_SAVE_ERROR_AT_KEY } = await import(
      '../src/collab/gridShared.js'
    );
    const { docId } = await setUpTestDoc();

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'docx', [], false);
    socket.requestSync();
    const mirrorMeta = socket.mirror.getMap<number>(DOC_META_MAP);

    vi.spyOn(store, 'saveOriginal').mockRejectedValue(new Error('디스크 꽉 참'));
    const saveSpy = vi.mocked(store.saveOriginal);

    sendEdit(socket, 'will fail to save');
    await vi.advanceTimersByTimeAsync(8_100);
    await flushMicrotasks();

    expect(saveSpy).toHaveBeenCalledTimes(1);
    const savedAt = mirrorMeta.get(DOC_META_SAVED_AT_KEY) ?? 0;
    const saveErrorAt = mirrorMeta.get(DOC_META_SAVE_ERROR_AT_KEY) ?? 0;
    expect(saveErrorAt).toBeGreaterThan(savedAt);

    // 실패 신호를 기록하는 쓰기 자신이 또 다른 자동저장을 예약했다면(자동저장 무한 루프 버그의
    // 재발), 추가 편집 없이도 이 시점에 saveOriginal이 다시 불렸을 것이다.
    await vi.advanceTimersByTimeAsync(30_000);
    await flushMicrotasks();
    expect(saveSpy).toHaveBeenCalledTimes(1);
  });

  test('a save that succeeds writes a fresh savedAt visible to clients', async () => {
    vi.useFakeTimers();

    const { joinEditSession } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');
    const { DOC_META_MAP, DOC_META_SAVED_AT_KEY } = await import('../src/collab/gridShared.js');
    const { docId } = await setUpTestDoc();

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'docx', [], false);
    socket.requestSync();
    const mirrorMeta = socket.mirror.getMap<number>(DOC_META_MAP);

    const saveSpy = vi.spyOn(store, 'saveOriginal');

    sendEdit(socket, 'should save successfully');
    await vi.advanceTimersByTimeAsync(8_100);
    await flushMicrotasks();

    expect(saveSpy).toHaveBeenCalledTimes(1);
    expect(mirrorMeta.get(DOC_META_SAVED_AT_KEY) ?? 0).toBeGreaterThan(0);
  });
});
