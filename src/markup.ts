import { randomUUID } from 'node:crypto';
import { store } from './storage.js';
import type { DocumentMeta } from './types.js';

export interface MarkupItem {
  readonly id: string;
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly text: string;
  readonly author: string;
  readonly createdAt: string;
}

export interface AddMarkupInput {
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly text: string;
  readonly author: string;
}

export class MarkupValidationError extends Error {}

const MAX_TEXT_LENGTH = 2000;
/** 좌표 비교 시 부동소수점 오차를 흡수하기 위한 여유값 */
const BOUNDS_EPSILON = 0.0001;

function assertValidInput(input: AddMarkupInput, pageCount: number): void {
  if (!Number.isInteger(input.page) || input.page < 1 || input.page > pageCount) {
    throw new MarkupValidationError(`잘못된 페이지 번호입니다: ${input.page} (전체 ${pageCount}페이지)`);
  }
  for (const [key, value] of Object.entries({ x: input.x, y: input.y, w: input.w, h: input.h })) {
    if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 1) {
      throw new MarkupValidationError(`잘못된 좌표값입니다: ${key}=${value}`);
    }
  }
  if (input.w <= 0 || input.h <= 0) {
    throw new MarkupValidationError('마크업 영역의 폭과 높이는 0보다 커야 합니다.');
  }
  if (input.x + input.w > 1 + BOUNDS_EPSILON || input.y + input.h > 1 + BOUNDS_EPSILON) {
    throw new MarkupValidationError('마크업 영역이 페이지를 벗어납니다.');
  }
  if (input.text.trim().length === 0) {
    throw new MarkupValidationError('코멘트 내용을 입력해 주세요.');
  }
  if (input.text.trim().length > MAX_TEXT_LENGTH) {
    throw new MarkupValidationError(`코멘트는 ${MAX_TEXT_LENGTH}자를 넘을 수 없습니다.`);
  }
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

async function readAll(docId: string): Promise<MarkupItem[]> {
  if (!(await store.hasDerived(docId, 'markup'))) return [];
  const buffer = await store.readDerived(docId, 'markup');
  return JSON.parse(buffer.toString('utf8')) as MarkupItem[];
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
    const item: MarkupItem = {
      id: randomUUID(),
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

export async function deleteMarkup(docId: string, markupId: string): Promise<boolean> {
  return withDocLock(docId, async () => {
    const items = await readAll(docId);
    const next = items.filter((item) => item.id !== markupId);
    if (next.length === items.length) return false;
    await writeAll(docId, next);
    return true;
  });
}
