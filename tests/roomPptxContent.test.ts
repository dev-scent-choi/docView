import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
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

/** slideContents 안, 텍스트박스 하나를 흉내 낸 최소 구조. 이 태스크 시점엔 아직 진짜
 * Y.XmlFragment 배선이 없으므로 값 종류는 중요하지 않다 — 키가 살아남는지만 본다. */
function makePlaceholderPerSlide(): Y.Map<unknown> {
  const shapeMap = new Y.Map();
  shapeMap.set('content', 'PLACEHOLDER');
  const perSlide = new Y.Map<unknown>();
  perSlide.set('shape-1', shapeMap);
  return perSlide;
}

describe('room.ts: pptx flushToStore preserves slide ids and slideContents across autosave', () => {
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

  test('a slide\'s id — and its slideContents — stay identical across autosave (no re-keying)', async () => {
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
      slideContents.set(entryId, makePlaceholderPerSlide());
    });

    // 자동저장을 강제로 한 번 돌린다(8초 유휴 타이머를 건너뛰고 직접 flushToStore 경로를 태운다).
    await vi.advanceTimersByTimeAsync(9000);
    await flushEventLoop();

    // 핵심 보장: 살아남은 슬라이드는 저장 전후로 id가 절대 바뀌지 않는다(더 이상 randomUUID를
    // 재발급하지 않으므로). id가 안정적이어야 slideContents가 같은 키 아래 아무 마이그레이션 없이
    // 그대로 유효하다 — 살아있는 에디터 바인딩이 존재해도(y-prosemirror ySyncPlugin 등, 아직 이
    // 태스크엔 없지만 다음 단계에서 반드시 생긴다) 서버가 그 바인딩이 쓰고 있는 서브트리를 지워버릴
    // 일이 없다는 뜻이다.
    const newEntryId = slidesArray.get(0).id;
    expect(newEntryId).toBe(entryId);
    expect(slideContents.has(entryId)).toBe(true);
    expect((slideContents.get(entryId) as Y.Map<unknown>).get('shape-1')).toBeDefined();
  });

  test('deleting a slide before the next autosave prunes its orphaned slideContents entry, ' +
    'but keeps the surviving slide\'s content and id', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Slide A', 'Slide B']));
    await store.writeMeta({
      id: docId,
      name: 'x.pptx',
      ext: 'pptx',
      status: 'ready',
      pageCount: 2,
      createdAt: new Date().toISOString(),
      progress: 100,
    });

    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'pptx', [], false);

    const ydoc = getRoomYdocForTest(docId)!;
    const slidesArray = ydoc.getArray<{ id: string }>(PPTX_SLIDES_ARRAY);
    expect(slidesArray.length).toBe(2);
    const keptId = slidesArray.get(0).id;
    const deletedId = slidesArray.get(1).id;
    const slideContents = ydoc.getMap(PPTX_SLIDE_CONTENTS_MAP);

    ydoc.transact(() => {
      slideContents.set(keptId, makePlaceholderPerSlide());
      slideContents.set(deletedId, makePlaceholderPerSlide());
    });

    // 사용자가 두 번째 슬라이드를 지운다 — 다음 자동저장이 돌기 전에.
    ydoc.transact(() => {
      slidesArray.delete(1, 1);
    });

    await vi.advanceTimersByTimeAsync(9000);
    await flushEventLoop();

    expect(slidesArray.length).toBe(1);
    // 살아남은 슬라이드는 이번 태스크의 다른 보장(id 안정성)대로 id도, 콘텐츠도 그대로다.
    expect(slidesArray.get(0).id).toBe(keptId);
    expect(slideContents.has(keptId)).toBe(true);
    // 삭제된 슬라이드가 가리키던 slideContents 항목은 더 이상 어떤 SlideEntry도 가리키지 않는
    // 도달 불가능한 쓰레기였다 — 자동저장이 이를 정리(prune)했어야 한다.
    expect(slideContents.has(deletedId)).toBe(false);
  });

  test('a malformed slideContents value (not a Y.Map) is pruned instead of aborting the autosave', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');
    const { DOC_META_MAP, DOC_META_SAVED_AT_KEY, DOC_META_SAVE_ERROR_AT_KEY } = await import(
      '../src/collab/gridShared.js'
    );

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
    const slideContents = ydoc.getMap<unknown>(PPTX_SLIDE_CONTENTS_MAP);
    // handleMessage는 클라이언트가 보낸 원시 Yjs 업데이트를 루트 타입 검증 없이 그대로 적용한다 —
    // 슬라이드 id는 동기화된 slidesArray로 누구나 볼 수 있으므로, 악의적이거나 버그가 있는
    // 클라이언트가 Y.Map이 아닌 값을 써 넣을 수 있는 상황을 그대로 흉내 낸다.
    ydoc.transact(() => {
      slideContents.set(entryId, 42);
    });

    await vi.advanceTimersByTimeAsync(9000);
    await flushEventLoop();

    // 잘못된 값은 조용히 정리됐고, 무엇보다 자동저장 자체는 죽지 않았다(=throw로 트랜잭션이
    // 취소되어 이 문서의 자동저장이 8초마다 영원히 실패하는 일이 없다).
    expect(slideContents.has(entryId)).toBe(false);
    const docMeta = ydoc.getMap<number>(DOC_META_MAP);
    expect(docMeta.get(DOC_META_SAVED_AT_KEY)).toEqual(expect.any(Number));
    expect(docMeta.get(DOC_META_SAVE_ERROR_AT_KEY)).toBeUndefined();
  });
});
