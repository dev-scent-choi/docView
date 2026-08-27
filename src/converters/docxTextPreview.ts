import fs from 'node:fs/promises';
import fontkit from '@pdf-lib/fontkit';
import { PDFDocument, PDFFont, StandardFonts, rgb } from 'pdf-lib';
import { importDocx } from '../docx/import.js';
import type { Node as PMNode } from 'prosemirror-model';
import type { Converter } from './types.js';

/**
 * LibreOffice가 없는 환경에서 docx 업로드·동시편집 저장 후 미리보기가 완전히 막히는 것을 막는
 * 대체 경로 — 서식 없는 순수 텍스트만으로 PDF 미리보기를 만든다(pdf-lib만 사용, 외부 도구 불필요).
 * LibreOffice가 설치돼 있으면 이 컨버터는 등록되지 않고 기존 officeConverter(정확한 렌더링)가
 * 그대로 쓰인다 — converters/index.ts에서 sofficePath 유무로 갈린다.
 */
const PAGE_WIDTH = 595.28; // A4 (pt)
const PAGE_HEIGHT = 841.89;
const MARGIN = 56;
const BODY_SIZE = 11;
const LINE_HEIGHT_RATIO = 1.45;
const HEADING_SIZES: Record<number, number> = { 1: 22, 2: 18, 3: 15, 4: 13, 5: 12, 6: 11 };

/**
 * pdf-lib 내장 Standard 폰트(WinAnsi)는 한글은 물론 ⚠ 같은 기호도 인코딩하지 못해 그대로 쓰면
 * 변환이 예외로 죽는다. 이 제품은 한국어 문서가 기본이므로, 시스템에 이미 깔려 있는 한글 폰트
 * (Windows는 기본 내장된 맑은 고딕)를 fontkit으로 그대로 임베드해 쓴다 — 폰트 파일을 리포에
 * 새로 추가하지 않고도 대부분의 실제 배포 환경에서 한글이 정상적으로 나온다.
 * 후보 경로가 전부 없는 환경(예: 한글 폰트 없는 리눅스 컨테이너)에서는 Standard 폰트로
 * 낮은 품질로라도 폴백하되, 인코딩 불가능한 문자는 죽지 않도록 안전하게 걸러낸다.
 */
const KOREAN_FONT_CANDIDATES: ReadonlyArray<{ regular: string; bold?: string }> = [
  { regular: 'C:\\Windows\\Fonts\\malgun.ttf', bold: 'C:\\Windows\\Fonts\\malgunbd.ttf' },
  { regular: '/usr/share/fonts/truetype/nanum/NanumGothic.ttf', bold: '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf' },
  { regular: '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc' },
  { regular: '/System/Library/Fonts/Supplemental/AppleGothic.ttf' },
];

interface PreviewFonts {
  readonly regular: PDFFont;
  readonly bold: PDFFont;
  readonly italic: PDFFont;
  /** false면 Standard(WinAnsi) 폴백 — 한글 등은 렌더링 전에 걸러내야 한다 */
  readonly supportsUnicode: boolean;
}

async function loadPreviewFonts(pdf: PDFDocument): Promise<PreviewFonts> {
  for (const candidate of KOREAN_FONT_CANDIDATES) {
    try {
      const regularBytes = await fs.readFile(candidate.regular);
      const boldBytes = candidate.bold ? await fs.readFile(candidate.bold).catch(() => regularBytes) : regularBytes;
      pdf.registerFontkit(fontkit);
      const regular = await pdf.embedFont(regularBytes, { subset: true });
      const bold = await pdf.embedFont(boldBytes, { subset: true });
      return { regular, bold, italic: regular, supportsUnicode: true };
    } catch {
      continue; // 이 경로엔 없음 — 다음 후보로
    }
  }
  const regular = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const italic = await pdf.embedFont(StandardFonts.HelveticaOblique);
  return { regular, bold, italic, supportsUnicode: false };
}

/** Standard 폰트 폴백 전용 — 인코딩 불가능한 문자를 '?'로 바꿔 예외 없이 그릴 수 있게 한다 */
function sanitizeForStandardFont(text: string, font: PDFFont): string {
  let out = '';
  for (const ch of text) {
    try {
      font.widthOfTextAtSize(ch, 10);
      out += ch;
    } catch {
      out += '?';
    }
  }
  return out;
}

function wrapLine(text: string, font: PDFFont, size: number, maxWidth: number): string[] {
  const words = text.split(/\s+/).filter((w) => w.length > 0);
  if (words.length === 0) return [''];
  const lines: string[] = [];
  let current = '';
  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word;
    if (font.widthOfTextAtSize(candidate, size) <= maxWidth || current === '') {
      current = candidate;
    } else {
      lines.push(current);
      current = word;
    }
  }
  if (current) lines.push(current);
  return lines;
}

function blockText(node: PMNode): string {
  if (node.type.name === 'table') {
    const rows: string[] = [];
    node.forEach((row) => {
      const cells: string[] = [];
      row.forEach((cell) => cells.push(cell.textContent.trim()));
      rows.push(cells.join('  |  '));
    });
    return rows.join('\n');
  }
  return node.textContent;
}

export const docxTextPreviewConverter: Converter = {
  name: 'docx-text-preview',
  extensions: ['docx'],
  async convert(sourcePath, targetPdfPath) {
    const bytes = await fs.readFile(sourcePath);
    const { doc } = importDocx(bytes);

    const pdf = await PDFDocument.create();
    const fonts = await loadPreviewFonts(pdf);
    const maxWidth = PAGE_WIDTH - MARGIN * 2;

    let page = pdf.addPage([PAGE_WIDTH, PAGE_HEIGHT]);
    let y = PAGE_HEIGHT - MARGIN;

    const ensureSpace = (needed: number) => {
      if (y - needed < MARGIN) {
        page = pdf.addPage([PAGE_WIDTH, PAGE_HEIGHT]);
        y = PAGE_HEIGHT - MARGIN;
      }
    };

    const drawText = (rawText: string, size: number, useFont: PDFFont, gapAfter: number) => {
      const text = fonts.supportsUnicode ? rawText : sanitizeForStandardFont(rawText, useFont);
      const lineHeight = size * LINE_HEIGHT_RATIO;
      const lines = text.split('\n').flatMap((line) => wrapLine(line, useFont, size, maxWidth));
      for (const line of lines) {
        ensureSpace(lineHeight);
        y -= lineHeight;
        page.drawText(line, { x: MARGIN, y, size, font: useFont, color: rgb(0.106, 0.102, 0.09) });
      }
      y -= gapAfter;
    };

    drawText(
      '[안내] LibreOffice가 설치되지 않아 서식 없는 텍스트 미리보기입니다 — 동시편집 내용 자체는 그대로 보존됩니다.',
      9,
      fonts.italic,
      14,
    );

    doc.forEach((node) => {
      const text = blockText(node);
      if (node.type.name === 'heading') {
        const level = (node.attrs.level as number) ?? 1;
        drawText(text || ' ', HEADING_SIZES[level] ?? BODY_SIZE, fonts.bold, 8);
      } else {
        drawText(text || ' ', BODY_SIZE, fonts.regular, 8);
      }
    });

    await fs.writeFile(targetPdfPath, await pdf.save());
  },
};
