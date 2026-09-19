import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSampleDocx } from './helpers.js';

/**
 * getParagraphChangeLog(room.ts / docxChangeLog.ts) — docx 동시편집 룸이 "어느 문단을 누가 언제
 * 무엇으로 바꿨는지" 누적해 돌려주는지 검증한다. xlsx의 roomCellChangeLog.test.ts와 같은 계약:
 * 룸이 열려 있는 동안만 조회 가능, 룸이 열린 시점 대비 누적, 편집자 교차 오염 없음.
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
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) {
        readUpdate(decoder, this.mirror, this);
      }
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}

  /** 연결 직후 클라이언트처럼 자기(빈) 상태 벡터를 보내 서버의 현재 내용을 mirror로 받는다. */
  requestSync(): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeSyncStep1(encoder, this.mirror);
    this.emit('message', encoding.toUint8Array(encoder));
  }

  /** mirror에서 편집을 수행하고, 그 결과로 생긴 update(전체 상태가 아니라 델타)만 서버로 보낸다 —
   * 동기화된 mirror에서 만든 델타여야 서버 Y.Doc의 실제 항목에 인과적으로 이어진다. */
  edit(action: (fragment: Y.XmlFragment) => void): void {
    let captured: Uint8Array | null = null;
    const onUpdate = (update: Uint8Array) => {
      captured = update;
    };
    this.mirror.on('update', onUpdate);
    action(this.mirror.getXmlFragment(XML_FRAGMENT_NAME));
    this.mirror.off('update', onUpdate);
    if (!captured) throw new Error('edit produced no update — was requestSync called first?');
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, captured);
    this.emit('message', encoding.toUint8Array(encoder));
  }
}

/** 문서 순서대로 모든 textblock(paragraph/heading) — 표 셀 안쪽까지. */
function textblocksOf(container: Y.XmlFragment | Y.XmlElement, out: Y.XmlElement[] = []): Y.XmlElement[] {
  for (const child of container.toArray()) {
    if (!(child instanceof Y.XmlElement)) continue;
    if (child.nodeName === 'paragraph' || child.nodeName === 'heading') out.push(child);
    else textblocksOf(child, out);
  }
  return out;
}

/** n번째(0-based) 문단의 글자 런(Y.XmlText). */
function textNodeOf(fragment: Y.XmlFragment, paragraphIndex: number): Y.XmlText {
  const block = textblocksOf(fragment)[paragraphIndex];
  const text = block.toArray().find((child): child is Y.XmlText => child instanceof Y.XmlText);
  if (!text) throw new Error(`문단 ${paragraphIndex}에 글자 런이 없습니다.`);
  return text;
}

function docxWithTable(intro: string, cells: readonly string[]): Buffer {
  const base = makeSampleDocx([intro]);
  const entries = unzipSync(base);
  const tableXml =
    '<w:tbl><w:tr>' +
    cells.map((c) => `<w:tc><w:p><w:r><w:t>${c}</w:t></w:r></w:p></w:tc>`).join('') +
    '</w:tr></w:tbl>';
  const xml = strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${tableXml}<w:sectPr/>`);
  entries['word/document.xml'] = strToU8(xml);
  return Buffer.from(zipSync(entries));
}

describe('collab/room.ts: getParagraphChangeLog (docx)', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-docx-changelog-'));
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

  async function openDoc(original: Buffer) {
    const { store } = await import('../src/storage.js');
    const room = await import('../src/collab/room.js');
    const docId = randomUUID();
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
    return { docId, room };
  }

  async function join(room: typeof import('../src/collab/room.js'), docId: string, viewer: string) {
    const socket = new FakeSocket();
    await room.joinEditSession(socket as never, {} as never, docId, 'docx', [], false, viewer);
    socket.requestSync();
    return socket;
  }

  test('returns null when no room is open, and [] for an open room with no edits', async () => {
    const { docId, room } = await openDoc(makeSampleDocx(['First', 'Second']));
    expect(room.getParagraphChangeLog(randomUUID())).toBeNull();
    expect(room.getParagraphChangeLog(docId)).toBeNull(); // 아직 아무도 안 들어옴 → 룸 없음

    await join(room, docId, 'viewer-a');
    // 접속·동기화만 했을 뿐 편집이 없으면, 클라이언트 쪽 정규화 같은 텍스트 무변화 쓰기가 있더라도
    // 변경으로 잡히면 안 된다.
    expect(room.getParagraphChangeLog(docId)).toEqual([]);
  });

  test('records before/after/viewer/timestamp for a text edit, with a 1-based paragraph number', async () => {
    const { docId, room } = await openDoc(makeSampleDocx(['First', 'Second']));
    const socket = await join(room, docId, 'viewer-a');

    const startedAt = Date.now();
    socket.edit((fragment) => textNodeOf(fragment, 1).insert('Second'.length, ' (edited)'));

    const log = room.getParagraphChangeLog(docId);
    expect(log).toHaveLength(1);
    expect(log![0]).toMatchObject({ paragraph: 2, before: 'Second', after: 'Second (edited)', viewer: 'viewer-a' });
    expect(new Date(log![0].at).getTime()).toBeGreaterThanOrEqual(startedAt);
  });

  test('two viewers editing different paragraphs are each attributed correctly', async () => {
    const { docId, room } = await openDoc(makeSampleDocx(['First', 'Second']));
    const a = await join(room, docId, 'viewer-a');
    const b = await join(room, docId, 'viewer-b');

    a.edit((fragment) => textNodeOf(fragment, 0).insert(0, 'A:'));
    b.edit((fragment) => textNodeOf(fragment, 1).insert(0, 'B:'));

    const byViewer = Object.fromEntries(room.getParagraphChangeLog(docId)!.map((e) => [e.viewer, e]));
    expect(byViewer['viewer-a']).toMatchObject({ paragraph: 1, before: 'First', after: 'A:First' });
    expect(byViewer['viewer-b']).toMatchObject({ paragraph: 2, before: 'Second', after: 'B:Second' });
  });

  test('editing the same paragraph twice keeps the ORIGINAL before-text and reports the latest editor', async () => {
    const { docId, room } = await openDoc(makeSampleDocx(['First']));
    const a = await join(room, docId, 'viewer-a');
    a.edit((fragment) => textNodeOf(fragment, 0).insert(5, '!'));

    const b = await join(room, docId, 'viewer-b');
    b.edit((fragment) => textNodeOf(fragment, 0).insert(6, '?'));

    const log = room.getParagraphChangeLog(docId)!;
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({ before: 'First', after: 'First!?', viewer: 'viewer-b' });
  });

  test('an inserted paragraph is reported with an empty before; an empty new paragraph is not', async () => {
    const { docId, room } = await openDoc(makeSampleDocx(['First']));
    const socket = await join(room, docId, 'viewer-a');

    socket.edit((fragment) => {
      const empty = new Y.XmlElement('paragraph');
      fragment.insert(1, [empty]);
    });
    expect(room.getParagraphChangeLog(docId)).toEqual([]);

    socket.edit((fragment) => {
      const paragraph = new Y.XmlElement('paragraph');
      const text = new Y.XmlText();
      text.insert(0, 'Brand new');
      paragraph.insert(0, [text]);
      fragment.insert(2, [paragraph]);
    });
    const log = room.getParagraphChangeLog(docId)!;
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({ paragraph: 3, before: '', after: 'Brand new', viewer: 'viewer-a' });
  });

  test('a deleted paragraph is reported with paragraph=null, its old text, and the deleter', async () => {
    const { docId, room } = await openDoc(makeSampleDocx(['First', 'Second']));
    const socket = await join(room, docId, 'viewer-a');

    socket.edit((fragment) => fragment.delete(0, 1));

    const log = room.getParagraphChangeLog(docId)!;
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({ paragraph: null, before: 'First', after: '', viewer: 'viewer-a' });
  });

  test('paragraphs inside table cells are tracked and numbered in document order', async () => {
    const { docId, room } = await openDoc(docxWithTable('Intro', ['A1', 'B1']));
    const socket = await join(room, docId, 'viewer-a');

    // Intro=1, A1=2, B1=3
    socket.edit((fragment) => textNodeOf(fragment, 2).insert(2, '+'));

    const log = room.getParagraphChangeLog(docId)!;
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({ paragraph: 3, before: 'B1', after: 'B1+', viewer: 'viewer-a' });
  });

  test('deleting a whole table attributes every cell paragraph in it to the deleter', async () => {
    const { docId, room } = await openDoc(docxWithTable('Intro', ['A1', 'B1']));
    const socket = await join(room, docId, 'viewer-a');

    socket.edit((fragment) => fragment.delete(1, 1)); // 표 통째로

    const log = room.getParagraphChangeLog(docId)!;
    expect(log.map((e) => e.before).sort()).toEqual(['A1', 'B1']);
    for (const entry of log) {
      expect(entry).toMatchObject({ paragraph: null, after: '', viewer: 'viewer-a' });
    }
  });

  test('a formatting-only change (same text) is not reported', async () => {
    const { docId, room } = await openDoc(makeSampleDocx(['First']));
    const socket = await join(room, docId, 'viewer-a');

    socket.edit((fragment) => textNodeOf(fragment, 0).format(0, 5, { strong: true }));

    expect(room.getParagraphChangeLog(docId)).toEqual([]);
  });

  test('a read-only connection cannot create a change entry', async () => {
    const { docId, room } = await openDoc(makeSampleDocx(['First']));
    const socket = new FakeSocket();
    await room.joinEditSession(socket as never, {} as never, docId, 'docx', [], true, 'observer');
    socket.requestSync();

    socket.edit((fragment) => textNodeOf(fragment, 0).insert(0, 'HACK '));

    expect(room.getParagraphChangeLog(docId)).toEqual([]);
  });
});
