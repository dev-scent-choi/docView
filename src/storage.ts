import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { config } from './config.js';
import type { DocumentMeta } from './types.js';

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const EXT_PATTERN = /^[a-z0-9]{1,10}$/;

/** 변환 산출물 종류: pdf(페이지 뷰·출력용) / grid(스프레드시트 그리드 뷰 JSON) */
export type DerivedKind = 'pdf' | 'grid';

/**
 * 문서 저장 계층 추상화. disk(기본)와 memory 두 구현을 DOCVIEW_STORAGE로 선택한다.
 * 외부에서 함수 호출로 재사용할 수 있도록 파일 경로가 아닌 버퍼 기반 API를 노출한다.
 */
export interface DocumentStore {
  readonly mode: 'disk' | 'memory';
  saveOriginal(id: string, ext: string, data: Buffer): Promise<void>;
  readOriginal(id: string, ext: string): Promise<Buffer>;
  saveDerived(id: string, kind: DerivedKind, data: Buffer): Promise<void>;
  readDerived(id: string, kind: DerivedKind): Promise<Buffer>;
  hasDerived(id: string, kind: DerivedKind): Promise<boolean>;
  writeMeta(meta: DocumentMeta): Promise<void>;
  readMeta(id: string): Promise<DocumentMeta | null>;
  listMeta(): Promise<readonly DocumentMeta[]>;
  deleteDocument(id: string): Promise<void>;
}

export function newDocumentId(): string {
  return randomUUID();
}

function assertValidId(id: string): void {
  if (!UUID_PATTERN.test(id)) {
    throw new Error('잘못된 문서 ID 형식입니다.');
  }
}

function assertValidExt(ext: string): void {
  if (!EXT_PATTERN.test(ext)) {
    throw new Error('잘못된 확장자 형식입니다.');
  }
}

function sortByCreatedAtDesc(metas: DocumentMeta[]): DocumentMeta[] {
  return metas.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export class DiskStore implements DocumentStore {
  readonly mode = 'disk' as const;

  private originalPath(id: string, ext: string): string {
    assertValidId(id);
    assertValidExt(ext);
    return path.join(config.originalsDir, `${id}.${ext}`);
  }

  private derivedPath(id: string, kind: DerivedKind): string {
    assertValidId(id);
    return path.join(config.convertedDir, kind === 'pdf' ? `${id}.pdf` : `${id}.grid.json`);
  }

  private metaPath(id: string): string {
    assertValidId(id);
    return path.join(config.metaDir, `${id}.json`);
  }

  async saveOriginal(id: string, ext: string, data: Buffer): Promise<void> {
    await fs.writeFile(this.originalPath(id, ext), data);
  }

  async readOriginal(id: string, ext: string): Promise<Buffer> {
    return fs.readFile(this.originalPath(id, ext));
  }

  async saveDerived(id: string, kind: DerivedKind, data: Buffer): Promise<void> {
    await fs.writeFile(this.derivedPath(id, kind), data);
  }

  async readDerived(id: string, kind: DerivedKind): Promise<Buffer> {
    return fs.readFile(this.derivedPath(id, kind));
  }

  async hasDerived(id: string, kind: DerivedKind): Promise<boolean> {
    try {
      await fs.access(this.derivedPath(id, kind));
      return true;
    } catch {
      return false;
    }
  }

  async writeMeta(meta: DocumentMeta): Promise<void> {
    await fs.writeFile(this.metaPath(meta.id), JSON.stringify(meta, null, 2), 'utf8');
  }

  async readMeta(id: string): Promise<DocumentMeta | null> {
    if (!UUID_PATTERN.test(id)) {
      return null;
    }
    try {
      const raw = await fs.readFile(this.metaPath(id), 'utf8');
      return JSON.parse(raw) as DocumentMeta;
    } catch {
      return null;
    }
  }

  async listMeta(): Promise<readonly DocumentMeta[]> {
    let entries: string[];
    try {
      entries = await fs.readdir(config.metaDir);
    } catch {
      return [];
    }
    const ids = entries
      .filter((name) => name.endsWith('.json'))
      .map((name) => name.slice(0, -'.json'.length));
    const metas = await Promise.all(ids.map((id) => this.readMeta(id)));
    return sortByCreatedAtDesc(metas.filter((meta): meta is DocumentMeta => meta !== null));
  }

  async deleteDocument(id: string): Promise<void> {
    const meta = await this.readMeta(id);
    if (!meta) return;
    await Promise.allSettled([
      fs.rm(this.originalPath(id, meta.ext), { force: true }),
      fs.rm(this.derivedPath(id, 'pdf'), { force: true }),
      fs.rm(this.derivedPath(id, 'grid'), { force: true }),
      fs.rm(this.metaPath(id), { force: true }),
    ]);
  }
}

/**
 * 메모리 저장소: 디스크에 아무것도 남기지 않는다 (변환 순간의 임시 파일 제외).
 * 서버 재시작 시 전부 소멸하며, 예산 초과 시 오래된 문서부터 제거한다.
 */
export class MemoryStore implements DocumentStore {
  readonly mode = 'memory' as const;
  private readonly originals = new Map<string, Buffer>();
  private readonly derived = new Map<string, Buffer>();
  private readonly metas = new Map<string, DocumentMeta>();
  private readonly budgetBytes: number;

  constructor(budgetBytes: number = config.memoryBudgetBytes) {
    this.budgetBytes = budgetBytes;
  }

  private totalBytes(): number {
    let total = 0;
    for (const buf of this.originals.values()) total += buf.length;
    for (const buf of this.derived.values()) total += buf.length;
    return total;
  }

  /** 예산 초과 시 보호 대상(protectedId)을 제외하고 가장 오래된 문서부터 제거한다. */
  private enforceBudget(protectedId: string): void {
    while (this.totalBytes() > this.budgetBytes) {
      const oldest = sortByCreatedAtDesc([...this.metas.values()])
        .reverse()
        .find((meta) => meta.id !== protectedId);
      if (!oldest) return;
      this.originals.delete(oldest.id);
      this.derived.delete(`${oldest.id}#pdf`);
      this.derived.delete(`${oldest.id}#grid`);
      this.metas.delete(oldest.id);
      console.warn(`메모리 예산 초과로 오래된 문서를 제거했습니다: ${oldest.name} (${oldest.id})`);
    }
  }

  async saveOriginal(id: string, ext: string, data: Buffer): Promise<void> {
    assertValidId(id);
    assertValidExt(ext);
    this.originals.set(id, data);
    this.enforceBudget(id);
  }

  async readOriginal(id: string, _ext: string): Promise<Buffer> {
    const data = this.originals.get(id);
    if (!data) throw new Error('문서가 메모리에서 만료되었습니다. 다시 업로드해 주세요.');
    return data;
  }

  async saveDerived(id: string, kind: DerivedKind, data: Buffer): Promise<void> {
    assertValidId(id);
    this.derived.set(`${id}#${kind}`, data);
    this.enforceBudget(id);
  }

  async readDerived(id: string, kind: DerivedKind): Promise<Buffer> {
    const data = this.derived.get(`${id}#${kind}`);
    if (!data) throw new Error('문서가 메모리에서 만료되었습니다. 다시 업로드해 주세요.');
    return data;
  }

  async hasDerived(id: string, kind: DerivedKind): Promise<boolean> {
    return this.derived.has(`${id}#${kind}`);
  }

  async writeMeta(meta: DocumentMeta): Promise<void> {
    assertValidId(meta.id);
    this.metas.set(meta.id, meta);
  }

  async readMeta(id: string): Promise<DocumentMeta | null> {
    return this.metas.get(id) ?? null;
  }

  async listMeta(): Promise<readonly DocumentMeta[]> {
    return sortByCreatedAtDesc([...this.metas.values()]);
  }

  async deleteDocument(id: string): Promise<void> {
    this.originals.delete(id);
    this.derived.delete(`${id}#pdf`);
    this.derived.delete(`${id}#grid`);
    this.metas.delete(id);
  }
}

export const store: DocumentStore =
  config.storageMode === 'memory' ? new MemoryStore() : new DiskStore();

export async function ensureDataDirs(): Promise<void> {
  const dirs = [config.auditDir];
  if (config.storageMode === 'disk') {
    dirs.push(config.originalsDir, config.convertedDir, config.metaDir);
  }
  await Promise.all(dirs.map((dir) => fs.mkdir(dir, { recursive: true })));
}
