import OrderedMap from 'orderedmap';
import { Schema, type MarkSpec, type NodeSpec } from 'prosemirror-model';
import { addListNodes } from 'prosemirror-schema-list';
import { tableNodes } from 'prosemirror-tables';

/**
 * docx 실시간 동시편집 1차 스코프에 맞춘 최소 스키마.
 * 지원: 문단, 제목(1~6), 굵게/기울임/밑줄/취소선, 표, 글머리·번호 목록(표시만 — 목록은
 * word/numbering.xml 연동 없이 내보낼 때 "• "/"1. " 접두사가 붙은 일반 문단으로 평탄화된다).
 * 미지원(가져오기 시 유실): 이미지, 머리글/바닥글, 각주, 하이퍼링크, 사용자 지정 스타일.
 */
/** paragraph/heading의 공통 정렬 속성 — DOM에는 인라인 text-align으로, docx로는 w:jc로 나간다 */
function alignToDOMStyle(align: unknown): Record<string, string> {
  return align ? { style: `text-align: ${align as string}` } : {};
}

const baseNodes: Record<string, NodeSpec> = {
  doc: { content: 'block+' },

  paragraph: {
    content: 'inline*',
    group: 'block',
    attrs: { align: { default: null } },
    parseDOM: [
      {
        tag: 'p',
        getAttrs: (dom) => ({ align: (dom as HTMLElement).style.textAlign || null }),
      },
    ],
    toDOM: (node) => ['p', alignToDOMStyle(node.attrs.align), 0],
  },

  heading: {
    attrs: { level: { default: 1 }, align: { default: null } },
    content: 'inline*',
    group: 'block',
    defining: true,
    parseDOM: [1, 2, 3, 4, 5, 6].map((level) => ({
      tag: `h${level}`,
      getAttrs: (dom) => ({ level, align: (dom as HTMLElement).style.textAlign || null }),
    })),
    toDOM: (node) => [`h${node.attrs.level as number}`, alignToDOMStyle(node.attrs.align), 0],
  },

  text: { group: 'inline' },

  /** docx의 w:br(문단 안 줄바꿈) — 새 문단이 아니라 같은 문단 안에서 줄만 바꾸는 것. */
  hard_break: {
    inline: true,
    group: 'inline',
    selectable: false,
    parseDOM: [{ tag: 'br' }],
    toDOM: () => ['br'],
  },
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
  strike: {
    parseDOM: [{ tag: 's' }],
    toDOM: () => ['s', 0],
  },
  textColor: {
    attrs: { color: {} },
    parseDOM: [
      {
        style: 'color',
        getAttrs: (value) => ({ color: value as string }),
      },
    ],
    toDOM: (mark) => ['span', { style: `color: ${mark.attrs.color as string}` }, 0],
  },
  fontSize: {
    /** 포인트(pt) 단위로 저장 — docx의 w:sz(하프포인트)와는 export/import 시 ×2/÷2 변환한다 */
    attrs: { size: {} },
    parseDOM: [
      {
        style: 'font-size',
        getAttrs: (value) => ({ size: parseFloat(value as string) }),
      },
    ],
    toDOM: (mark) => ['span', { style: `font-size: ${mark.attrs.size as number}pt` }, 0],
  },
};

const withLists = addListNodes(OrderedMap.from(baseNodes), 'paragraph block*', 'block');
const withTables = withLists.append(
  tableNodes({
    tableGroup: 'block',
    cellContent: 'block+',
    cellAttributes: {
      /** 셀 음영(docx의 w:tcPr/w:shd) — "#rrggbb" 또는 없음(null) */
      background: {
        default: null,
        getFromDOM: (dom) => (dom as HTMLElement).style.backgroundColor || null,
        setDOMAttr: (value, attrs) => {
          if (value) attrs.style = `${(attrs.style as string | undefined) ?? ''}background-color: ${value as string};`;
        },
      },
    },
  }),
);

export const docxSchema = new Schema({ nodes: withTables, marks });
