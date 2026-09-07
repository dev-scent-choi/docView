import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSampleDocx } from './helpers.js';

const MESSAGE_SYNC = 0;
const XML_FRAGMENT_NAME = 'prosemirror';

/**
 * joinEditSession이 필요로 하는 최소한의 WebSocket 인터페이스만 흉내 낸다 — 실제 네트워크나
 * 인증(routes/edit.ts) 없이 room.ts의 룸 생성·메시지 처리 로직만 직접 거친다.
 */
class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  send(_data: Uint8Array, cb?: (err?: Error) => void): void {
    cb?.();
  }
  close(): void {}
  ping(): void {}
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
    await fs.mkdir(path.join(tmpDir, 'audit'), { recursive: true });
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    // PDF 변환기는 이 테스트의 관심사가 아니다(LibreOffice 등 외부 도구 없이도 돌게) — findConverter가
    // null을 돌려주게 해 flushToStore가 변환 단계를 건너뛰게 한다.
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('autosave does not perpetually re-fire after its own save-status write', async () => {
    vi.useFakeTimers();

    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');

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

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'docx', [], false);

    const saveSpy = vi.spyOn(store, 'saveOriginal');

    // 실제 원격 편집 한 번을 시뮬레이션한다 — 별도의 독립된 Y.Doc에서 만든 업데이트를 서버
    // 소켓의 'message' 이벤트로 그대로 보낸다(CRDT라 서로 다른 Y.Doc 간 업데이트도 그대로 병합됨).
    const scratch = new Y.Doc();
    const fragment = scratch.getXmlFragment(XML_FRAGMENT_NAME);
    const paragraph = new Y.XmlElement('paragraph');
    const text = new Y.XmlText();
    text.insert(0, 'edited concurrently');
    paragraph.insert(0, [text]);
    fragment.insert(0, [paragraph]);
    const update = Y.encodeStateAsUpdate(scratch);

    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, update);
    socket.emit('message', encoding.toUint8Array(encoder));

    // 8초 유휴 자동저장이 한 번 돈다.
    await vi.advanceTimersByTimeAsync(8_100);
    expect(saveSpy).toHaveBeenCalledTimes(1);

    // 그 저장 자신이 기록한 저장 시각(docMeta) 쓰기가 또 다른 자동저장을 예약했다면, 추가
    // 편집이 전혀 없어도 이 시점에 saveOriginal이 다시 불렸을 것이다 — 그러면 버그(무한 루프)다.
    await vi.advanceTimersByTimeAsync(30_000);
    expect(saveSpy).toHaveBeenCalledTimes(1);
  });
});
