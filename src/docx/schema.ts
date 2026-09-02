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

/** 셀 한 변의 테두리 — docx의 w:tcBorders/w:tblBorders 한 변(top/right/bottom/left)에 대응 */
export interface CellBorderSide {
  /** CSS border-style 값(solid/dashed/dotted/double) — docx w:val(single/dashed/dotted/double 등)에서 변환 */
  readonly style: string;
  readonly widthPx: number;
  /** "#rrggbb" */
  readonly color: string;
}

export interface CellBorders {
  readonly top: CellBorderSide | null;
  readonly right: CellBorderSide | null;
  readonly bottom: CellBorderSide | null;
  readonly left: CellBorderSide | null;
}

const BORDER_SIDES = ['top', 'right', 'bottom', 'left'] as const;

function borderSideToCSS(side: CellBorderSide | null): string {
  return side ? `${side.widthPx}px ${side.style} ${side.color}` : 'none';
}

/** "1.5px solid rgb(255, 0, 0)" 같은 계산된 CSS 값을 다시 CellBorderSide로 되돌린다(HTML 붙여넣기 등 DOM 기반 입력용). */
function borderSideFromDOM(cssValue: string): CellBorderSide | null {
  if (!cssValue || cssValue === 'none') return null;
  const match = /^([\d.]+)px\s+(\S+)\s+(.+)$/.exec(cssValue.trim());
  if (!match) return null;
  const [, widthStr, style] = match;
  return { style, widthPx: parseFloat(widthStr), color: match[3] };
}

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
      /** 셀 4변 테두리(docx의 w:tcBorders/w:tblBorders에서 계산된 최종값) — 없으면 null(에디터 기본 테두리 사용) */
      borders: {
        default: null,
        getFromDOM: (dom) => {
          const style = (dom as HTMLElement).style;
          const borders: CellBorders = {
            top: borderSideFromDOM(style.borderTop),
            right: borderSideFromDOM(style.borderRight),
            bottom: borderSideFromDOM(style.borderBottom),
            left: borderSideFromDOM(style.borderLeft),
          };
          return borders.top || borders.right || borders.bottom || borders.left ? borders : null;
        },
        setDOMAttr: (value, attrs) => {
          if (!value) return;
          const borders = value as CellBorders;
          const extra = BORDER_SIDES.map((side) => `border-${side}: ${borderSideToCSS(borders[side])};`).join(' ');
          attrs.style = `${(attrs.style as string | undefined) ?? ''}${extra}`;
        },
      },
    },
  }),
);

export const docxSchema = new Schema({ nodes: withTables, marks });
