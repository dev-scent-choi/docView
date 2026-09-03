import OrderedMap from 'orderedmap';
import { Schema, type MarkSpec, type Node as PMNode, type NodeSpec } from 'prosemirror-model';
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

/**
 * docx의 w:sectPr/w:pgSz·w:pgMar — 페이지 크기·여백. import.ts가 원본에서 그대로 읽어오고,
 * export.ts는 이 스키마와 무관하게 원본 w:sectPr 자체를 그대로 보존해 내보낸다(우리가
 * 모델링하지 않는 세부 설정까지 안전하게 되돌아가도록). 여기 있는 값은 오직 "화면에 실제
 * Word 페이지와 같은 비율로 얼마나 넓게 그릴지" 계산에만 쓰인다.
 */
export interface PageSetup {
  readonly widthPx: number;
  readonly heightPx: number;
  readonly marginTopPx: number;
  readonly marginBottomPx: number;
  readonly marginLeftPx: number;
  readonly marginRightPx: number;
  readonly orientation: 'portrait' | 'landscape';
}

/** docx의 w:trPr/w:trHeight — 행 높이. rule이 'auto'면 이 속성 자체가 없음(null)과 동일하게 다룬다. */
export interface RowHeight {
  readonly rule: 'atLeast' | 'exact';
  readonly px: number;
}

function rowHeightToDOMStyle(height: unknown): Record<string, string> {
  const h = height as RowHeight | null;
  return h ? { style: `height: ${h.px}px;` } : {};
}

function rowHeightFromDOM(dom: HTMLElement): RowHeight | null {
  const px = parseFloat(dom.style.height);
  return Number.isFinite(px) && px > 0 ? { rule: 'atLeast', px } : null;
}

/** docx의 w:tblPr/w:jc — 표 자체가 페이지(콘텐츠 영역) 안에서 왼쪽/가운데/오른쪽 중 어디에 놓이는지.
 * 셀 안 문단 정렬(paragraph의 align)과는 완전히 별개 개념이다. */
function tableAlignToDOMStyle(align: unknown): Record<string, string> {
  if (align === 'center') return { style: 'margin-left: auto; margin-right: auto;' };
  if (align === 'right') return { style: 'margin-left: auto; margin-right: 0;' };
  return {};
}

/**
 * 표 자체의 실제 너비(px) — 첫 행 각 셀의 colwidth(모두 있을 때만) 합. 하나라도 없으면
 * (예: 편집기에서 새로 만든 표) null을 돌려주고, CSS 기본값(width:100%)에 맡긴다.
 * Word 원본 표가 페이지 폭보다 좁을 때 억지로 100%까지 늘어나 셀 너비가 실제보다 커지는
 * 문제(사용자가 "Cell 너비가 Word 원본과 다르다"고 지적한 부분)를 막기 위함이다.
 */
function tableWidthPx(node: PMNode): number | null {
  const firstRow = node.firstChild;
  if (!firstRow) return null;
  let total = 0;
  let ok = true;
  firstRow.forEach((cell) => {
    const colwidth = cell.attrs.colwidth as number[] | null;
    if (!Array.isArray(colwidth) || colwidth.length === 0) {
      ok = false;
      return;
    }
    total += colwidth.reduce((sum, w) => sum + w, 0);
  });
  return ok && total > 0 ? total : null;
}

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

const VALID_VERTICAL_ALIGN = new Set(['top', 'middle', 'bottom']);

/** 셀 안쪽 여백(docx의 w:tblCellMar/w:tcMar에서 계산된 최종값, px) — 4변 모두 확정된 값이다. */
export interface CellPadding {
  readonly topPx: number;
  readonly rightPx: number;
  readonly bottomPx: number;
  readonly leftPx: number;
}

function cellPaddingToCSS(padding: CellPadding): string {
  return `${padding.topPx}px ${padding.rightPx}px ${padding.bottomPx}px ${padding.leftPx}px`;
}

/** 계산된 CSS padding 값("Tpx Rpx Bpx Lpx")을 다시 CellPadding으로 되돌린다(HTML 붙여넣기 등 DOM 기반 입력용). */
function cellPaddingFromDOM(cssValue: string): CellPadding | null {
  if (!cssValue) return null;
  const parts = cssValue.trim().split(/\s+/).map(parseFloat);
  if (parts.length !== 4 || parts.some((v) => Number.isNaN(v))) return null;
  const [topPx, rightPx, bottomPx, leftPx] = parts;
  return { topPx, rightPx, bottomPx, leftPx };
}

const tableNodeSpecs = tableNodes({
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
    /** 셀 안 내용의 세로 정렬(docx의 w:tcPr/w:vAlign) — 문단의 가로 정렬(align)과는 별개 축이다 */
    verticalAlign: {
      default: null,
      getFromDOM: (dom) => {
        const v = (dom as HTMLElement).style.verticalAlign;
        return v === 'top' || v === 'middle' || v === 'bottom' ? v : null;
      },
      setDOMAttr: (value, attrs) => {
        if (value && VALID_VERTICAL_ALIGN.has(value as string)) {
          attrs.style = `${(attrs.style as string | undefined) ?? ''}vertical-align: ${value as string};`;
        }
      },
    },
    /**
     * 셀 안쪽 여백(docx의 w:tblCellMar/w:tcMar) — CSS 기본값(0.4rem 0.6rem)은 모든 셀에
     * 똑같이 적용되는 임의값이라, 원본이 이보다 좁은 여백을 쓰면 그 차이만큼 텍스트가 덜
     * 들어가 원본에선 한 줄인 게 편집기에선 줄바꿈되는 원인이 된다. 없으면 null(CSS 기본값 사용).
     */
    cellPadding: {
      default: null,
      getFromDOM: (dom) => cellPaddingFromDOM((dom as HTMLElement).style.padding),
      setDOMAttr: (value, attrs) => {
        if (value) attrs.style = `${(attrs.style as string | undefined) ?? ''}padding: ${cellPaddingToCSS(value as CellPadding)};`;
      },
    },
  },
});

/**
 * prosemirror-tables가 만들어주는 table/table_row 스펙은 attrs가 아예 없다(셀 안쪽만 다룸) —
 * "표 자체가 페이지 어디에 놓이는지(정렬)"와 "행 높이"는 표/행 자신의 속성이라 별도로 얹는다.
 */
const patchedTableNodeSpecs: Record<string, NodeSpec> = {
  ...tableNodeSpecs,
  table: {
    ...tableNodeSpecs.table,
    attrs: { align: { default: null } },
    toDOM: (node) => {
      const widthPx = tableWidthPx(node);
      const alignStyle = tableAlignToDOMStyle(node.attrs.align).style ?? '';
      const widthStyle = widthPx ? `width: ${widthPx}px;` : '';
      const style = `${widthStyle}${alignStyle}`;
      return ['table', style ? { style } : {}, ['tbody', 0]];
    },
  },
  table_row: {
    ...tableNodeSpecs.table_row,
    attrs: { height: { default: null } },
    parseDOM: [{ tag: 'tr', getAttrs: (dom) => ({ height: rowHeightFromDOM(dom as HTMLElement) }) }],
    toDOM: (node) => ['tr', rowHeightToDOMStyle(node.attrs.height), 0],
  },
};

const withLists = addListNodes(OrderedMap.from(baseNodes), 'paragraph block*', 'block');
const withTables = withLists.append(patchedTableNodeSpecs);

export const docxSchema = new Schema({ nodes: withTables, marks });
