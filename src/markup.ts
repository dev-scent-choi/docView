import { randomUUID } from 'node:crypto';
import { store } from './storage.js';
import type { DocumentMeta } from './types.js';

interface MarkupItemBase {
  readonly id: string;
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly text: string;
  readonly author: string;
  readonly createdAt: string;
}

/** 사각형 영역 + 코멘트 (기존 리뷰 마크업) */
export interface CommentMarkupItem extends MarkupItemBase {
  readonly kind: 'comment';
  readonly w: number;
  readonly h: number;
}

/** 주변 폰트에 맞춰 삽입하는 실제 텍스트 (신규) — x,y는 텍스트 베이스라인의 좌측 시작점(비율) */
export interface TextMarkupItem extends MarkupItemBase {
  readonly kind: 'text';
  readonly fontSizePt: number;
  readonly bold: boolean;
}

export type MarkupKind = MarkupItem['kind'];
export type MarkupItem = CommentMarkupItem | TextMarkupItem;

export interface AddCommentInput {
  readonly kind?: 'comment';
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly text: string;
  readonly author: string;
}

export interface AddTextInput {
  readonly kind: 'text';
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly fontSizePt: number;
  readonly bold: boolean;
  readonly text: string;
  readonly author: string;
}

export type AddMarkupInput = AddCommentInput | AddTextInput;

export interface UpdateMarkupInput {
  readonly x: number;
  readonly y: number;
}

export class MarkupValidationError extends Error {}

const MAX_TEXT_LENGTH = 2000;
const MIN_FONT_SIZE_PT = 1;
const MAX_FONT_SIZE_PT = 400;
/** 좌표 비교 시 부동소수점 오차를 흡수하기 위한 여유값 */
const BOUNDS_EPSILON = 0.0001;

function assertValidCoordinate(key: string, value: number): void {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 1) {
    throw new MarkupValidationError(`잘못된 좌표값입니다: ${key}=${value}`);
  }
}

function assertValidPage(page: number, pageCount: number): void {
  if (!Number.isInteger(page) || page < 1 || page > pageCount) {
    throw new MarkupValidationError(`잘못된 페이지 번호입니다: ${page} (전체 ${pageCount}페이지)`);
  }
}

function assertValidText(text: string): void {
  if (text.trim().length === 0) {
    throw new MarkupValidationError('내용을 입력해 주세요.');
  }
  if (text.trim().length > MAX_TEXT_LENGTH) {
    throw new MarkupValidationError(`내용은 ${MAX_TEXT_LENGTH}자를 넘을 수 없습니다.`);
  }
}

function assertValidInput(input: AddMarkupInput, pageCount: number): void {
  assertValidPage(input.page, pageCount);
  assertValidCoordinate('x', input.x);
  assertValidCoordinate('y', input.y);

  if (input.kind === 'text') {
    if (
      typeof input.fontSizePt !== 'number' ||
      !Number.isFinite(input.fontSizePt) ||
      input.fontSizePt < MIN_FONT_SIZE_PT ||
      input.fontSizePt > MAX_FONT_SIZE_PT
    ) {
      throw new MarkupValidationError(
        `잘못된 글자 크기입니다: ${input.fontSizePt} (${MIN_FONT_SIZE_PT}~${MAX_FONT_SIZE_PT}pt)`,
      );
    }
    assertValidText(input.text);
    return;
  }

  assertValidCoordinate('w', input.w);
  assertValidCoordinate('h', input.h);
  if (input.w <= 0 || input.h <= 0) {
    throw new MarkupValidationError('마크업 영역의 폭과 높이는 0보다 커야 합니다.');
  }
  if (input.x + input.w > 1 + BOUNDS_EPSILON || input.y + input.h > 1 + BOUNDS_EPSILON) {
    throw new MarkupValidationError('마크업 영역이 페이지를 벗어납니다.');
  }
  assertValidText(input.text);
}

/**
 * 문서 ID별 순차 실행 큐. 같은 문서에 대한 추가/삭제가 동시에 들어와도
 * "읽기 → 수정 → 쓰기"가 서로 겹쳐 한쪽의 변경을 덮어쓰지 않게 한다.
 * (src/converters/libreOffice.ts의 프로필 슬롯 큐와 동일한 패턴)
 */
const docLocks = new Map<string, Promise<unknown>>();

async function withDocLock<T>(docId: string, task: () => Promise<T>): Promise<T> {
  const previous = docLocks.get(docId) ?? Promise.resolve();
  let release!: () => void;
  const current = new Promise<void>((resolve) => {
    release = resolve;
  });
  docLocks.set(docId, previous.then(() => current));
  await previous.catch(() => {});
  try {
    return await task();
  } finally {
    release();
  }
}

/** kind 필드가 생기기 전(2026-09-02 이전)에 저장된 항목 — 전부 comment였다. */
type StoredMarkupItem = (Omit<CommentMarkupItem, 'kind'> & { kind?: 'comment' }) | TextMarkupItem;

function normalize(item: StoredMarkupItem): MarkupItem {
  if (item.kind === 'text') return item;
  return { ...item, kind: 'comment' };
}

async function readAll(docId: string): Promise<MarkupItem[]> {
  if (!(await store.hasDerived(docId, 'markup'))) return [];
  const buffer = await store.readDerived(docId, 'markup');
  const raw = JSON.parse(buffer.toString('utf8')) as StoredMarkupItem[];
  return raw.map(normalize);
}

async function writeAll(docId: string, items: readonly MarkupItem[]): Promise<void> {
  await store.saveDerived(docId, 'markup', Buffer.from(JSON.stringify(items), 'utf8'));
}

export async function listMarkup(docId: string): Promise<readonly MarkupItem[]> {
  return readAll(docId);
}

export async function addMarkup(
  docId: string,
  meta: DocumentMeta,
  input: AddMarkupInput,
): Promise<MarkupItem> {
  assertValidInput(input, meta.pageCount);
  return withDocLock(docId, async () => {
    const items = await readAll(docId);
    const item: MarkupItem =
      input.kind === 'text'
        ? {
            id: randomUUID(),
            kind: 'text',
            page: input.page,
            x: input.x,
            y: input.y,
            fontSizePt: input.fontSizePt,
            bold: input.bold,
            text: input.text.trim(),
            author: input.author,
            createdAt: new Date().toISOString(),
          }
        : {
            id: randomUUID(),
            kind: 'comment',
            page: input.page,
            x: input.x,
            y: input.y,
            w: input.w,
            h: input.h,
            text: input.text.trim(),
            author: input.author,
            createdAt: new Date().toISOString(),
          };
    await writeAll(docId, [...items, item]);
    return item;
  });
}

/** 드래그로 위치만 옮긴다 — 크기/글자/폰트 등 다른 속성은 바뀌지 않는다. */
export async function updateMarkup(
  docId: string,
  markupId: string,
  patch: UpdateMarkupInput,
): Promise<MarkupItem | null> {
  assertValidCoordinate('x', patch.x);
  assertValidCoordinate('y', patch.y);
  return withDocLock(docId, async () => {
    const items = await readAll(docId);
    const idx = items.findIndex((item) => item.id === markupId);
    if (idx === -1) return null;
    const existing = items[idx];
    const w = existing.kind === 'comment' ? existing.w : 0;
    const h = existing.kind === 'comment' ? existing.h : 0;
    if (patch.x + w > 1 + BOUNDS_EPSILON || patch.y + h > 1 + BOUNDS_EPSILON) {
      throw new MarkupValidationError('마크업 영역이 페이지를 벗어납니다.');
    }
    const updated: MarkupItem = { ...existing, x: patch.x, y: patch.y };
    const next = [...items];
    next[idx] = updated;
    await writeAll(docId, next);
    return updated;
  });
}

export async function deleteMarkup(docId: string, markupId: string): Promise<boolean> {
  return withDocLock(docId, async () => {
    const items = await readAll(docId);
    const next = items.filter((item) => item.id !== markupId);
    if (next.length === items.length) return false;
    await writeAll(docId, next);
    return true;
  });
}
