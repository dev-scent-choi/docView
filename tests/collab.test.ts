import * as Y from 'yjs';
import { prosemirrorToYXmlFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import { describe, expect, test } from 'vitest';
import { exportDocx } from '../src/docx/export.js';
import { importDocx } from '../src/docx/import.js';
import { docxSchema } from '../src/docx/schema.js';
import { makeSampleDocx } from './helpers.js';

/**
 * src/collab/room.ts의 bindState/writeState가 실제로 쓰는 것과 동일한 변환
 * (docx → ProseMirror → Y.XmlFragment → [편집] → ProseMirror → docx)을
 * 실제 WebSocket 서버 없이 재현한다. 두 클라이언트가 여기에 붙는 Y.Doc을 공유한다고 가정하고,
 * "다른 클라이언트가 만든 변경"을 Yjs 문서를 직접 조작해 시뮬레이션한다.
 */
describe('collab: docx <-> Yjs <-> ProseMirror bridge', () => {
  test('seeds a Y.Doc from docx, survives a concurrent-style edit, and flushes back to a valid docx', () => {
    const original = makeSampleDocx(['Hello DocView']);
    const { doc: seedDoc } = importDocx(original);

    // 룸 생성 시점(bindState)과 동일한 시딩
    const ydoc = new Y.Doc();
    const fragment = ydoc.getXmlFragment('prosemirror');
    prosemirrorToYXmlFragment(seedDoc, fragment);

    // "다른 브라우저 탭"이 편집한 것처럼 Y 구조를 직접 조작한다 (실서비스에서는 y-prosemirror의
    // ySyncPlugin이 ProseMirror 트랜잭션을 이 형태로 변환해준다 — 여기서는 그 산출물만 재현)
    const paragraph = fragment.get(0) as Y.XmlElement;
    const textNode = paragraph.get(0) as Y.XmlText;
    expect(textNode.toString()).toBe('Hello DocView');
    textNode.insert(textNode.length, ' — edited concurrently');

    // 저장(writeState/유휴 자동저장)과 동일한 역변환
    const pmDocAfterEdit = yXmlFragmentToProseMirrorRootNode(fragment, docxSchema);
    expect(pmDocAfterEdit.textContent).toBe('Hello DocView — edited concurrently');

    const savedDocx = exportDocx(original, pmDocAfterEdit);
    const reopened = importDocx(savedDocx);
    expect(reopened.doc.textContent).toBe('Hello DocView — edited concurrently');
  });

  test('two independent inserts into the same Y.Doc both survive (CRDT merge, no last-writer-wins loss)', () => {
    const original = makeSampleDocx(['Shared paragraph']);
    const { doc: seedDoc } = importDocx(original);

    const ydoc = new Y.Doc();
    const fragment = ydoc.getXmlFragment('prosemirror');
    prosemirrorToYXmlFragment(seedDoc, fragment);

    // 두 "클라이언트"가 같은 시작 상태에서 갈라져 각자 편집한다
    const clientA = new Y.Doc();
    Y.applyUpdate(clientA, Y.encodeStateAsUpdate(ydoc));
    const clientB = new Y.Doc();
    Y.applyUpdate(clientB, Y.encodeStateAsUpdate(ydoc));

    const editIn = (doc: Y.Doc, suffix: string) => {
      const frag = doc.getXmlFragment('prosemirror');
      const paragraph = frag.get(0) as Y.XmlElement;
      const text = paragraph.get(0) as Y.XmlText;
      text.insert(text.length, suffix);
    };
    editIn(clientA, ' [A]');
    editIn(clientB, ' [B]');

    // 서버가 두 업데이트를 모두 받아 병합한다고 가정
    Y.applyUpdate(ydoc, Y.encodeStateAsUpdate(clientA, Y.encodeStateVector(ydoc)));
    Y.applyUpdate(ydoc, Y.encodeStateAsUpdate(clientB, Y.encodeStateVector(ydoc)));

    const merged = yXmlFragmentToProseMirrorRootNode(fragment, docxSchema);
    expect(merged.textContent).toContain('[A]');
    expect(merged.textContent).toContain('[B]');
    expect(merged.textContent).toContain('Shared paragraph');
  });
});
