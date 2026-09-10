import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSamplePptx } from './helpers.js';

const MESSAGE_SYNC = 0;

/**
 * joinEditSession이 필요로 하는 최소한의 WebSocket 인터페이스만 흉내 낸다 — roomAutosaveLoop.test.ts/
 * pptxCollab.test.ts와 같은 패턴. 이 테스트는 room.ts가 직접 관리하는 ydoc을 getRoomYdocForTest로
 * 들여다보므로 FakeSocket은 서버가 메시지를 보낼 때 예외 없이 받아주기만 하면 된다.
 */
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
}

/** 페이지 채우기/재계산은 청크마다 실제 setImmediate로 양보하므로 그 턴을 흘려보낸다(pptxCollab.test.ts와 동일). */
async function flushEventLoop(turns = 30): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

describe('room.ts: pptx flushToStore re-keys slideContents to the new post-save ids', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-pptx-content-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });
    // PDF 변환기는 이 테스트의 관심사가 아니다 — findConverter가 null을 돌려주게 해 flushToStore가
    // 변환 단계를 건너뛰게 한다(roomAutosaveLoop.test.ts와 동일한 이유).
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    // recordAudit의 실제 디스크 I/O(fs.appendFile) 완료는 libuv 폴 단계를 타서 가짜 타이머로는
    // 안정적으로 기다려지지 않는다 — 이 테스트의 관심사가 아니므로 동기 완료로 교체한다.
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

  test('content written before a save is still readable by the same SlideEntry after autosave renumbers ids', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Slide A']));
    await store.writeMeta({
      id: docId,
      name: 'x.pptx',
      ext: 'pptx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      progress: 100,
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'pptx', [], false);

    const ydoc = getRoomYdocForTest(docId)!;
    const slidesArray = ydoc.getArray<{ id: string }>(PPTX_SLIDES_ARRAY);
    const entryId = slidesArray.get(0).id;
    const slideContents = ydoc.getMap(PPTX_SLIDE_CONTENTS_MAP);
    // 이 테스트는 room.ts의 실제 시딩 로직(Task 8)을 아직 안 쓰고, "id로 찾을 수 있다"만
    // 검증하는 것이 목적이라 슬라이드 콘텐츠를 직접 흉내 낸 최소 구조로 채운다.
    ydoc.transact(() => {
      const shapeMap = new Y.Map();
      shapeMap.set('content', 'PLACEHOLDER'); // 이 태스크 시점엔 Y.XmlFragment 배선이 아직 없음 — 값 종류는 중요하지 않다, 키가 살아남는지만 본다
      const perSlide = new Y.Map();
      perSlide.set('shape-1', shapeMap);
      slideContents.set(entryId, perSlide);
    });

    // 자동저장을 강제로 한 번 돌린다(8초 유휴 타이머를 건너뛰고 직접 flushToStore 경로를 태운다).
    await vi.advanceTimersByTimeAsync(9000);
    await flushEventLoop();

    const newEntryId = slidesArray.get(0).id;
    expect(newEntryId).not.toBe(entryId); // id는 정말 바뀌었다(기존 동작 확인)
    expect(slideContents.has(entryId)).toBe(false); // 옛 키는 사라졌고
    expect(slideContents.has(newEntryId)).toBe(true); // 새 키로 옮겨졌다
    expect((slideContents.get(newEntryId) as Y.Map<unknown>).get('shape-1')).toBeDefined();
  });
});
