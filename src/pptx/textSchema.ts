import { Schema, type MarkSpec, type NodeSpec } from 'prosemirror-model';

/**
 * pptx 텍스트박스 전용 최소 ProseMirror 스키마. docx의 docxSchema(src/docx/schema.ts)와
 * 마크 정의(강조 종류·속성)는 동일하게 맞춘다 — import/export 시 같은 pt/색 표현을 그대로
 * 재사용하기 위해서다. 다만 텍스트박스는 표·목록·제목이 필요 없어 노드는 doc/paragraph/text만
 * 남긴다(docxSchema 그대로 재사용하면 안 쓸 커맨드·키맵까지 끌고 들어온다).
 */
const nodes: Record<string, NodeSpec> = {
  doc: { content: 'paragraph+' },
  paragraph: {
    content: 'inline*',
    group: 'block',
    parseDOM: [{ tag: 'p' }],
    toDOM: () => ['p', 0],
  },
  text: { group: 'inline' },
};

const marks: Record<string, MarkSpec> = {
  strong: {
    parseDOM: [{ tag: 'strong' }, { tag: 'b' }],
    toDOM: () => ['strong', 0],
  },
  em: {
    parseDOM: [{ tag: 'em' }, { tag: 'i' }],
    toDOM: () => ['em', 0],
  },
  underline: {
    parseDOM: [{ tag: 'u' }],
    toDOM: () => ['u', 0],
  },
  textColor: {
    attrs: { color: {} },
    parseDOM: [{ style: 'color', getAttrs: (value) => ({ color: value as string }) }],
    toDOM: (mark) => ['span', { style: `color: ${mark.attrs.color as string}` }, 0],
  },
  fontSize: {
    /** 포인트(pt) 단위 — pptx의 <a:rPr sz="..">(100분의 1 포인트)와는 import/export 시 변환한다. */
    attrs: { size: {} },
    parseDOM: [{ style: 'font-size', getAttrs: (value) => ({ size: parseFloat(value as string) }) }],
    toDOM: (mark) => ['span', { style: `font-size: ${mark.attrs.size as number}pt` }, 0],
  },
};

export const pptxTextSchema = new Schema({ nodes, marks });
