import * as Y from 'yjs';

/**
 * docx 본문(docxChangeLog.ts)과 pptx 텍스트박스(pptxChangeLog.ts)는 둘 다 y-prosemirror가 만든
 * Y.XmlFragment로 텍스트를 담는다 — "textblock(paragraph/heading)의 현재 텍스트"를 읽는 코드를
 * 두 변경 추적기가 그대로 공유한다. pptx 텍스트 스키마(textSchema.ts)는 doc/paragraph/text뿐이라
 * heading이 없지만, 이름 집합에 들어 있어도 해가 없다.
 */
const TEXTBLOCK_NODE_NAMES: ReadonlySet<string> = new Set(['paragraph', 'heading']);

export function isTextblock(node: unknown): node is Y.XmlElement {
  return node instanceof Y.XmlElement && TEXTBLOCK_NODE_NAMES.has(node.nodeName);
}

/** textblock 하나의 현재 텍스트. y-prosemirror는 연속된 글자 런을 Y.XmlText(서식은 delta
 * attributes)로, 그 외 inline 노드(hard_break)는 Y.XmlElement 자식으로 만든다. */
export function textOfBlock(block: Y.XmlElement): string {
  let text = '';
  for (const child of block.toArray()) {
    if (child instanceof Y.XmlText) {
      for (const op of child.toDelta() as Array<{ insert?: unknown }>) {
        if (typeof op.insert === 'string') text += op.insert;
      }
    } else if (child instanceof Y.XmlElement && child.nodeName === 'hard_break') {
      text += '\n';
    }
  }
  return text;
}

/** 문서 순서대로(표·목록 안쪽까지) 모든 textblock을 모은다. */
export function collectTextblocks(container: Y.XmlFragment | Y.XmlElement, out: Y.XmlElement[]): void {
  for (const child of container.toArray()) {
    if (!(child instanceof Y.XmlElement)) continue;
    if (isTextblock(child)) out.push(child);
    else collectTextblocks(child, out);
  }
}
