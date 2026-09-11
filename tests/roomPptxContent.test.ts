import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate } from 'y-protocols/sync';
import { prosemirrorToYXmlFragment } from 'y-prosemirror';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSamplePptx } from './helpers.js';
import type { SlideEntry } from '../src/collab/pptxShared.js';
import type { PerSlideContentsMap } from '../src/pptx/collab.js';

const MESSAGE_SYNC = 0;
/** room.ts의 내부 상수와 같은 값 — room.ts는 외부에 노출하지 않으므로 프로토콜 계약으로 여기서도
 * 그대로 리터럴을 쓴다(Task 6 브리프와 동일한 관례). */
const MESSAGE_SLIDE_CONTENT_SUBSCRIBE = 5;

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

describe('room.ts: pptx slide-content lazy seeding (MESSAGE_SLIDE_CONTENT_SUBSCRIBE) and export merge', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-pptx-seed-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
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

  /** MESSAGE_SLIDE_CONTENT_SUBSCRIBE 메시지를 인코딩해 FakeSocket에 흘려보낸다(브리프의 프로토콜 그대로,
   * room.ts의 handleMessage가 decoding.readVarString으로 읽는 것과 짝을 맞춘다). */
  function subscribeSlideContent(socket: FakeSocket, slideEntryId: string, duplicateOfId?: string): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SLIDE_CONTENT_SUBSCRIBE);
    encoding.writeVarString(
      encoder,
      JSON.stringify(duplicateOfId ? { slideEntryId, duplicateOfId } : { slideEntryId }),
    );
    socket.emit('message', encoding.toUint8Array(encoder));
  }

  test('opening a slide for the first time seeds its text boxes from the original XML', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Hello World']));
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
    const entryId = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY).get(0).id;
    const slideContents = ydoc.getMap<PerSlideContentsMap>(PPTX_SLIDE_CONTENTS_MAP);
    expect(slideContents.has(entryId)).toBe(false); // 열기 전엔 비어 있어야 한다(지연 로딩)

    subscribeSlideContent(socket, entryId);
    await flushEventLoop();

    const perSlide = slideContents.get(entryId);
    expect(perSlide).toBeDefined();
    expect(perSlide!.size).toBeGreaterThan(0); // "Hello World" 텍스트박스 하나가 채워졌다
  });

  test('opening the same slide twice does not overwrite already-edited content', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Hello World']));
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
    const entryId = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY).get(0).id;
    const slideContents = ydoc.getMap<PerSlideContentsMap>(PPTX_SLIDE_CONTENTS_MAP);

    subscribeSlideContent(socket, entryId);
    await flushEventLoop();

    // 사용자가 뭔가 편집했다고 흉내낸다 — 표식용 필드를 하나 심는다(shapeId -> ShapeContentMap이라는
    // PerSlideContentsMap의 정식 값 타입과는 다르지만, 이 테스트는 "키가 살아남는지"만 보면 되므로
    // 느슨한 Y.Map<unknown>으로 다뤄도 된다).
    const perSlideLoose = slideContents.get(entryId) as unknown as Y.Map<unknown>;
    ydoc.transact(() => {
      perSlideLoose.set('__edited_marker', true);
    });

    subscribeSlideContent(socket, entryId); // 같은 슬라이드를 다시 "연다"
    await flushEventLoop();

    expect(perSlideLoose.get('__edited_marker')).toBe(true); // 안 덮어써짐
  });

  test('a slide that was never opened is exported byte-for-byte identical to its original XML', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');
    const { exportPptxFromSlides } = await import('../src/pptx/collab.js');
    const { unzipSync, strFromU8 } = await import('fflate');

    const docId = randomUUID();
    const original = makeSamplePptx(['Slide One', 'Slide Two']);
    await store.saveOriginal(docId, 'pptx', original);
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
    const slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
    const firstEntryId = slidesArray.get(0).id;

    // 첫 번째 슬라이드만 연다(시딩) — 두 번째 슬라이드는 한 번도 구독하지 않는다.
    subscribeSlideContent(socket, firstEntryId);
    await flushEventLoop();

    const slideContents = ydoc.getMap<PerSlideContentsMap>(PPTX_SLIDE_CONTENTS_MAP);
    const result = exportPptxFromSlides(original, slidesArray, slideContents);

    const originalEntries = unzipSync(original);
    const secondSourcePart = slidesArray.get(1).sourceSlidePart!;
    const exportedEntries = unzipSync(result.bytes);
    const secondFinalPart = result.finalParts[1];

    expect(strFromU8(exportedEntries[`ppt/slides/${secondFinalPart}`])).toBe(
      strFromU8(originalEntries[`ppt/slides/${secondSourcePart}`]),
    );
  });

  test("editing a seeded shape's content merges into the exported slide XML (existing-shape path)", async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');
    const { exportPptxFromSlides } = await import('../src/pptx/collab.js');
    const { pptxTextSchema } = await import('../src/pptx/textSchema.js');
    const { strFromU8, unzipSync } = await import('fflate');

    const docId = randomUUID();
    const original = makeSamplePptx(['Hello World']);
    await store.saveOriginal(docId, 'pptx', original);
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
    const slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
    const entryId = slidesArray.get(0).id;

    subscribeSlideContent(socket, entryId);
    await flushEventLoop();

    const slideContents = ydoc.getMap<PerSlideContentsMap>(PPTX_SLIDE_CONTENTS_MAP);
    const perSlide = slideContents.get(entryId)!;
    // 샘플 pptx의 유일한 도형은 <p:cNvPr id="2" .../>다(helpers.ts의 pptxSlideXml 참고).
    const shapeMap = perSlide.get('2')!;
    expect(shapeMap).toBeDefined();
    const fragment = shapeMap.get('content') as Y.XmlFragment;

    // 실제 클라이언트가 y-prosemirror의 ySyncPlugin을 통해 하는 것과 같은 동작을 직접 흉내낸다 —
    // 이미 채워진(=이미 Y.Doc에 붙어 있는) fragment에 새 ProseMirror 문서를 다시 반영한다.
    const editedDoc = pptxTextSchema.node('doc', null, [
      pptxTextSchema.node('paragraph', null, [pptxTextSchema.text('Edited By Test')]),
    ]);
    ydoc.transact(() => {
      prosemirrorToYXmlFragment(editedDoc, fragment);
    });

    const result = exportPptxFromSlides(original, slidesArray, slideContents);
    const exportedEntries = unzipSync(result.bytes);
    const exportedSlideXml = strFromU8(exportedEntries[`ppt/slides/${result.finalParts[0]}`]);
    expect(exportedSlideXml).toContain('Edited By Test');
    expect(exportedSlideXml).not.toContain('Hello World');
  });

  test('a new-<uuid> shape id is routed to appendNewShape, not spliceShapeTxBody', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');
    const { exportPptxFromSlides } = await import('../src/pptx/collab.js');
    const { pptxTextSchema } = await import('../src/pptx/textSchema.js');
    const { strFromU8, unzipSync } = await import('fflate');

    const docId = randomUUID();
    const original = makeSamplePptx(['Hello World']);
    await store.saveOriginal(docId, 'pptx', original);
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
    const slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
    const entryId = slidesArray.get(0).id;

    subscribeSlideContent(socket, entryId);
    await flushEventLoop();

    const slideContents = ydoc.getMap<PerSlideContentsMap>(PPTX_SLIDE_CONTENTS_MAP);
    const perSlide = slideContents.get(entryId)!;

    // 클라이언트가 새 텍스트박스를 추가한 상황을 흉내낸다 — shapeId는 'new-' 접두어 관례를 쓴다
    // (pptxShared.ts/브리프 참고). seedSlideContentFromXml과 같은 순서로 먼저 부모에 붙이고 그
    // 다음 채운다(붙이기 전에 채우면 Yjs가 "premature access" 경고를 남긴다 — collab.ts 주석 참고).
    const newShapeId = `new-${randomUUID()}`;
    const newDoc = pptxTextSchema.node('doc', null, [
      pptxTextSchema.node('paragraph', null, [pptxTextSchema.text('Brand New Box')]),
    ]);
    ydoc.transact(() => {
      const shapeMap = new Y.Map<unknown>();
      perSlide.set(newShapeId, shapeMap); // 먼저 attach
      shapeMap.set('meta', { shapeId: newShapeId, x: 100, y: 200, w: 300, h: 400 });
      const fragment = new Y.XmlFragment();
      shapeMap.set('content', fragment); // 먼저 attach
      prosemirrorToYXmlFragment(newDoc, fragment); // 그 다음 채우기
    });

    const result = exportPptxFromSlides(original, slidesArray, slideContents);
    const exportedEntries = unzipSync(result.bytes);
    const exportedSlideXml = strFromU8(exportedEntries[`ppt/slides/${result.finalParts[0]}`]);
    expect(exportedSlideXml).toContain('Brand New Box');
    expect(exportedSlideXml).toContain('Hello World'); // 기존 도형은 그대로 남아 있어야 한다(appendNewShape는 덧붙이기만 한다)
    expect(exportedSlideXml).toContain(`id="${newShapeId}"`);
  });

  test('duplicateOfId seeds the source slide first, then deep-copies its content independently', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Source Text', 'Other Text']));
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
    const slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
    const sourceEntry = slidesArray.get(0);
    const dupEntryId = randomUUID();

    // 클라이언트가 슬라이드를 복제하면(Task 8) 먼저 새 SlideEntry를 배열에 넣고, 그 다음 이 WS
    // 메시지를 보낸다 — 그 순서를 그대로 흉내낸다.
    ydoc.transact(() => {
      slidesArray.insert(slidesArray.length, [
        { id: dupEntryId, kind: 'duplicate', sourceSlidePart: sourceEntry.sourceSlidePart, sourcePageNumber: null },
      ]);
    });

    const slideContents = ydoc.getMap<PerSlideContentsMap>(PPTX_SLIDE_CONTENTS_MAP);
    expect(slideContents.has(sourceEntry.id)).toBe(false); // 소스도 아직 안 열려 있었다

    subscribeSlideContent(socket, dupEntryId, sourceEntry.id);
    await flushEventLoop();

    // 소스가 먼저 시딩됐다.
    expect(slideContents.has(sourceEntry.id)).toBe(true);
    // 복제본도 시딩됐고, 소스와 같은 도형 수를 가진다.
    const dupPerSlide = slideContents.get(dupEntryId);
    expect(dupPerSlide).toBeDefined();
    const sourcePerSlide = slideContents.get(sourceEntry.id)!;
    expect(dupPerSlide!.size).toBe(sourcePerSlide.size);
    expect(dupPerSlide!.size).toBeGreaterThan(0);

    // 독립적인 깊은 복사인지 확인한다 — 소스를 수정해도 복제본은 영향을 받지 않는다(참조 공유가 아님).
    ydoc.transact(() => {
      sourcePerSlide.get('2')!.set('__source_only_marker', true);
    });
    expect(dupPerSlide!.get('2')!.has('__source_only_marker')).toBe(false);
  });

  test('subscribing to a blank slide seeds an empty per-slide map (no shapes to parse)', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP } = await import('../src/collab/pptxShared.js');

    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Only Slide']));
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
    const slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
    const blankId = randomUUID();
    ydoc.transact(() => {
      slidesArray.insert(slidesArray.length, [{ id: blankId, kind: 'blank', sourceSlidePart: null, sourcePageNumber: null }]);
    });

    subscribeSlideContent(socket, blankId);
    await flushEventLoop();

    const slideContents = ydoc.getMap<PerSlideContentsMap>(PPTX_SLIDE_CONTENTS_MAP);
    const perSlide = slideContents.get(blankId);
    expect(perSlide).toBeDefined();
    expect(perSlide!.size).toBe(0); // 빈 슬라이드는 파싱할 도형이 없다
  });
});
