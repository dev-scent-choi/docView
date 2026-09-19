import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import { prosemirrorToYXmlFragment } from 'y-prosemirror';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSamplePptx } from './helpers.js';
import type { SlideEntry } from '../src/collab/pptxShared.js';
import type { PerSlideContentsMap } from '../src/pptx/collab.js';

/**
 * getPptxChangeLog(room.ts / pptxChangeLog.ts) — pptx 동시편집 룸이 "어느 슬라이드의 어느 텍스트박스를
 * 누가 언제 무엇으로 바꿨는지, 슬라이드가 추가/삭제됐는지"를 누적해 돌려주는지 검증한다. xlsx/docx와
 * 같은 계약: 룸이 열려 있는 동안만 조회 가능, 룸이 열린 시점 대비 누적, 편집자 교차 오염 없음.
 */
const MESSAGE_SYNC = 0;
/** room.ts의 내부 상수와 같은 값(roomPptxContent.test.ts와 같은 관례로 프로토콜 계약을 리터럴로 쓴다). */
const MESSAGE_SLIDE_CONTENT_SUBSCRIBE = 5;
const SLIDES = 'slides';
const CONTENTS = 'slideContents';

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

  /** 클라이언트가 슬라이드를 열 때 보내는 지연 시딩 요청. 서버가 시딩한 결과는 broadcast로 mirror에 온다. */
  subscribe(slideEntryId: string, duplicateOfId?: string): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SLIDE_CONTENT_SUBSCRIBE);
    encoding.writeVarString(encoder, JSON.stringify(duplicateOfId ? { slideEntryId, duplicateOfId } : { slideEntryId }));
    this.emit('message', encoding.toUint8Array(encoder));
  }

  /** mirror에서 편집하고 그 델타만 서버로 보낸다(동기화된 mirror에서 만든 델타여야 인과적으로 이어진다). */
  edit(action: (doc: Y.Doc) => void): void {
    let captured: Uint8Array | null = null;
    const onUpdate = (update: Uint8Array) => {
      captured = captured === null ? update : Y.mergeUpdates([captured, update]);
    };
    this.mirror.on('update', onUpdate);
    this.mirror.transact(() => action(this.mirror));
    this.mirror.off('update', onUpdate);
    if (!captured) throw new Error('edit produced no update — was the slide seeded and synced first?');
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, captured);
    this.emit('message', encoding.toUint8Array(encoder));
  }
}

async function flushEventLoop(turns = 30): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

/** mirror 안의 텍스트박스 하나(슬라이드 id + shapeId)의 첫 글자 런. */
function textNodeOf(doc: Y.Doc, slideId: string, shapeId: string): Y.XmlText {
  const perSlide = doc.getMap<PerSlideContentsMap>(CONTENTS).get(slideId);
  const content = perSlide?.get(shapeId)?.get('content') as Y.XmlFragment | undefined;
  const paragraph = content?.toArray()[0];
  const text = paragraph instanceof Y.XmlElement ? paragraph.toArray().find((c): c is Y.XmlText => c instanceof Y.XmlText) : null;
  if (!text) throw new Error(`슬라이드 ${slideId}의 텍스트박스 ${shapeId}에 글자 런이 없습니다 — 시딩/동기화가 끝났는지 확인하세요.`);
  return text;
}

describe('collab/room.ts: getPptxChangeLog (pptx)', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-pptx-changelog-'));
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
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  async function openDeck(slideTexts: readonly string[]) {
    const { store } = await import('../src/storage.js');
    const room = await import('../src/collab/room.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(slideTexts));
    await store.writeMeta({
      id: docId,
      name: 'x.pptx',
      ext: 'pptx',
      status: 'ready',
      pageCount: slideTexts.length,
      createdAt: new Date().toISOString(),
      progress: 100,
    });
    return { docId, room };
  }

  async function join(room: typeof import('../src/collab/room.js'), docId: string, viewer: string, readOnly = false) {
    const socket = new FakeSocket();
    await room.joinEditSession(socket as never, {} as never, docId, 'pptx', [], readOnly, viewer);
    socket.requestSync();
    return socket;
  }

  /** 서버가 mirror로 흘려보낸 슬라이드 id 목록(현재 덱 순서). */
  function slideIds(socket: FakeSocket): string[] {
    return socket.mirror.getArray<SlideEntry>(SLIDES).toArray().map((e) => e.id);
  }

  /** 슬라이드를 열어(구독) 시딩된 내용이 mirror에 도착할 때까지 기다린다. */
  async function openSlide(socket: FakeSocket, slideId: string, duplicateOfId?: string): Promise<void> {
    socket.subscribe(slideId, duplicateOfId);
    await flushEventLoop();
  }

  test('returns null when no room is open, and [] for an open room with nothing edited', async () => {
    const { docId, room } = await openDeck(['Slide A', 'Slide B']);
    expect(room.getPptxChangeLog(randomUUID())).toBeNull();
    expect(room.getPptxChangeLog(docId)).toBeNull();

    const socket = await join(room, docId, 'viewer-a');
    expect(room.getPptxChangeLog(docId)).toEqual([]);

    // 슬라이드를 열어 서버가 원본에서 시딩하는 것은 사용자 편집이 아니다 — 기준선일 뿐 변경이 아니다.
    await openSlide(socket, slideIds(socket)[0]);
    expect(room.getPptxChangeLog(docId)).toEqual([]);
  });

  test('records slide number, shape, before/after, viewer and timestamp for a text edit', async () => {
    const { docId, room } = await openDeck(['Slide A', 'Slide B']);
    const socket = await join(room, docId, 'viewer-a');
    const [, second] = slideIds(socket);
    await openSlide(socket, second);

    const startedAt = Date.now();
    socket.edit((doc) => textNodeOf(doc, second, '2').insert('Slide B'.length, ' (edited)'));

    const log = room.getPptxChangeLog(docId)!;
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({
      kind: 'text',
      slideId: second,
      slide: 2,
      originalSlide: null,
      shape: '2',
      before: 'Slide B',
      after: 'Slide B (edited)',
      viewer: 'viewer-a',
    });
    expect(new Date(log[0].at).getTime()).toBeGreaterThanOrEqual(startedAt);
  });

  test('two viewers editing different slides are each attributed correctly', async () => {
    const { docId, room } = await openDeck(['Slide A', 'Slide B']);
    const a = await join(room, docId, 'viewer-a');
    const b = await join(room, docId, 'viewer-b');
    const [first, second] = slideIds(a);
    await openSlide(a, first);
    await openSlide(b, second);

    a.edit((doc) => textNodeOf(doc, first, '2').insert(0, 'A:'));
    b.edit((doc) => textNodeOf(doc, second, '2').insert(0, 'B:'));

    const byViewer = Object.fromEntries(room.getPptxChangeLog(docId)!.map((e) => [e.viewer, e]));
    expect(byViewer['viewer-a']).toMatchObject({ slide: 1, before: 'Slide A', after: 'A:Slide A' });
    expect(byViewer['viewer-b']).toMatchObject({ slide: 2, before: 'Slide B', after: 'B:Slide B' });
  });

  test('editing the same text box twice keeps the ORIGINAL before-text and reports the latest editor', async () => {
    const { docId, room } = await openDeck(['Slide A']);
    const a = await join(room, docId, 'viewer-a');
    const [first] = slideIds(a);
    await openSlide(a, first);
    a.edit((doc) => textNodeOf(doc, first, '2').insert('Slide A'.length, '!'));

    const b = await join(room, docId, 'viewer-b');
    await openSlide(b, first);
    b.edit((doc) => textNodeOf(doc, first, '2').insert('Slide A!'.length, '?'));

    const log = room.getPptxChangeLog(docId)!;
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({ before: 'Slide A', after: 'Slide A!?', viewer: 'viewer-b' });
  });

  test('a format-only change (text unchanged) is not reported', async () => {
    const { docId, room } = await openDeck(['Slide A']);
    const socket = await join(room, docId, 'viewer-a');
    const [first] = slideIds(socket);
    await openSlide(socket, first);

    socket.edit((doc) => textNodeOf(doc, first, '2').format(0, 5, { bold: {} }));

    expect(room.getPptxChangeLog(docId)).toEqual([]);
  });

  test('a new text box on an existing slide is reported with before "", and a deleted one with after ""', async () => {
    const { docId, room } = await openDeck(['Slide A']);
    const socket = await join(room, docId, 'viewer-a');
    const [first] = slideIds(socket);
    await openSlide(socket, first);

    const { pptxTextSchema } = await import('../src/pptx/textSchema.js');
    const newShapeId = `new-${randomUUID()}`;
    socket.edit((doc) => {
      const perSlide = doc.getMap<PerSlideContentsMap>(CONTENTS).get(first)!;
      const shapeMap = new Y.Map<unknown>();
      perSlide.set(newShapeId, shapeMap);
      shapeMap.set('meta', { shapeId: newShapeId, x: 10, y: 20, w: 300, h: 100 });
      const fragment = new Y.XmlFragment();
      shapeMap.set('content', fragment);
      prosemirrorToYXmlFragment(
        pptxTextSchema.node('doc', null, [pptxTextSchema.node('paragraph', null, [pptxTextSchema.text('Brand new')])]),
        fragment,
      );
    });

    let log = room.getPptxChangeLog(docId)!;
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({ kind: 'text', slide: 1, shape: newShapeId, before: '', after: 'Brand new', viewer: 'viewer-a' });

    // 이제 원본에 있던 텍스트박스("2")를 지운다.
    socket.edit((doc) => {
      doc.getMap<PerSlideContentsMap>(CONTENTS).get(first)!.delete('2');
    });
    log = room.getPptxChangeLog(docId)!;
    expect(log).toHaveLength(2);
    expect(log.find((e) => e.shape === '2')).toMatchObject({ kind: 'text', slide: 1, before: 'Slide A', after: '', viewer: 'viewer-a' });
  });

  test('a slide added after the room opened is reported as slide_added (with its position), including text typed into it', async () => {
    const { docId, room } = await openDeck(['Slide A']);
    const socket = await join(room, docId, 'viewer-a');
    const blankId = randomUUID();
    socket.edit((doc) => {
      doc.getArray<SlideEntry>(SLIDES).push([{ id: blankId, kind: 'blank', sourceSlidePart: null, sourcePageNumber: null }]);
    });
    await openSlide(socket, blankId);

    let log = room.getPptxChangeLog(docId)!;
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({ kind: 'slide_added', slideId: blankId, slide: 2, originalSlide: null, shape: null, viewer: 'viewer-a' });

    const { pptxTextSchema } = await import('../src/pptx/textSchema.js');
    const shapeId = `new-${randomUUID()}`;
    socket.edit((doc) => {
      const perSlide = doc.getMap<PerSlideContentsMap>(CONTENTS).get(blankId)!;
      const shapeMap = new Y.Map<unknown>();
      perSlide.set(shapeId, shapeMap);
      shapeMap.set('meta', { shapeId, x: 0, y: 0, w: 100, h: 50 });
      const fragment = new Y.XmlFragment();
      shapeMap.set('content', fragment);
      prosemirrorToYXmlFragment(
        pptxTextSchema.node('doc', null, [pptxTextSchema.node('paragraph', null, [pptxTextSchema.text('Typed here')])]),
        fragment,
      );
    });
    log = room.getPptxChangeLog(docId)!;
    expect(log.map((e) => e.kind)).toEqual(['slide_added', 'text']);
    expect(log[1]).toMatchObject({ slide: 2, shape: shapeId, before: '', after: 'Typed here', viewer: 'viewer-a' });
  });

  test('a duplicated slide is slide_added, and only edits made AFTER duplication are text changes (vs the copy)', async () => {
    const { docId, room } = await openDeck(['Slide A']);
    const socket = await join(room, docId, 'viewer-a');
    const [first] = slideIds(socket);
    const dupId = randomUUID();
    socket.edit((doc) => {
      doc.getArray<SlideEntry>(SLIDES).push([{ id: dupId, kind: 'duplicate', sourceSlidePart: 'slide1.xml', sourcePageNumber: null }]);
    });
    await openSlide(socket, dupId, first);

    expect(room.getPptxChangeLog(docId)!.map((e) => e.kind)).toEqual(['slide_added']);

    socket.edit((doc) => textNodeOf(doc, dupId, '2').insert('Slide A'.length, ' copy'));
    const log = room.getPptxChangeLog(docId)!;
    expect(log).toHaveLength(2);
    expect(log[0]).toMatchObject({ kind: 'slide_added', slide: 2 });
    expect(log[1]).toMatchObject({ kind: 'text', slide: 2, shape: '2', before: 'Slide A', after: 'Slide A copy', viewer: 'viewer-a' });
  });

  test('a removed original slide is slide_removed with its ORIGINAL position; remaining slides are renumbered', async () => {
    const { docId, room } = await openDeck(['Slide A', 'Slide B', 'Slide C']);
    const socket = await join(room, docId, 'viewer-a');
    const [, second, third] = slideIds(socket);
    await openSlide(socket, third);

    socket.edit((doc) => doc.getArray<SlideEntry>(SLIDES).delete(1, 1)); // 두 번째 슬라이드 삭제
    socket.edit((doc) => textNodeOf(doc, third, '2').insert(0, 'C:'));

    const log = room.getPptxChangeLog(docId)!;
    expect(log).toHaveLength(2);
    // 세 번째였던 슬라이드가 이제 2번째다 — 번호는 조회 시점의 현재 위치다.
    expect(log[0]).toMatchObject({ kind: 'text', slide: 2, before: 'Slide C', after: 'C:Slide C' });
    expect(log[1]).toMatchObject({ kind: 'slide_removed', slideId: second, slide: null, originalSlide: 2, viewer: 'viewer-a' });
    expect(slideIds(socket)).not.toContain(second);
  });

  test('reordering slides is not reported', async () => {
    const { docId, room } = await openDeck(['Slide A', 'Slide B']);
    const socket = await join(room, docId, 'viewer-a');
    socket.edit((doc) => {
      const slides = doc.getArray<SlideEntry>(SLIDES);
      const [a, b] = slides.toArray();
      slides.delete(0, 2);
      slides.insert(0, [b, a]);
    });
    expect(room.getPptxChangeLog(docId)).toEqual([]);
  });

  test('the log survives an autosave (which rewrites the slide list with the same ids)', async () => {
    const { docId, room } = await openDeck(['Slide A']);
    const socket = await join(room, docId, 'viewer-a');
    const [first] = slideIds(socket);
    await openSlide(socket, first);
    socket.edit((doc) => textNodeOf(doc, first, '2').insert('Slide A'.length, '!'));

    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });
    await vi.advanceTimersByTimeAsync(9000);
    await flushEventLoop();
    vi.useRealTimers();

    const log = room.getPptxChangeLog(docId)!;
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({ kind: 'text', slide: 1, before: 'Slide A', after: 'Slide A!', viewer: 'viewer-a' });
  });

  test('a read-only connection cannot create entries', async () => {
    const { docId, room } = await openDeck(['Slide A']);
    const readOnly = await join(room, docId, 'viewer-ro', true);
    const [first] = slideIds(readOnly);
    await openSlide(readOnly, first);

    readOnly.edit((doc) => textNodeOf(doc, first, '2').insert(0, 'HACK:'));

    expect(room.getPptxChangeLog(docId)).toEqual([]);
  });

  test('getChangeLog dispatches by document kind and is [] for an unknown document', async () => {
    const { docId, room } = await openDeck(['Slide A']);
    expect(room.getChangeLog(randomUUID())).toEqual([]);
    const socket = await join(room, docId, 'viewer-a');
    const [first] = slideIds(socket);
    await openSlide(socket, first);
    socket.edit((doc) => textNodeOf(doc, first, '2').insert(0, 'X'));
    expect(room.getChangeLog(docId)).toEqual(room.getPptxChangeLog(docId));
    expect(room.getChangeLog(docId)).toHaveLength(1);
  });
});
