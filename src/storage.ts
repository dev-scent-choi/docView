import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { config } from './config.js';
import type { DocumentMeta } from './types.js';

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const EXT_PATTERN = /^[a-z0-9]{1,10}$/;

/** 변환 산출물 종류: pdf(페이지 뷰·출력용) / grid(스프레드시트 그리드 뷰 JSON) / markup(리뷰 코멘트 JSON)
 * / pagesizes(PDF 페이지 크기 배열, 뷰어 첫 화면 속도용) / gridindex(그리드 JSONL 줄 오프셋, 동일 목적)
 * / svg(도면 dxf/dwg의 벡터 뷰 — pdf와 별개로 열람 전용 부가 산출물, 실패해도 pdf가 그대로 폴백이다) */
export type DerivedKind = 'pdf' | 'grid' | 'markup' | 'pagesizes' | 'gridindex' | 'svg';

/** kind별 파일 확장자 매핑 — 한 곳에서만 관리해, 새 kind를 추가하고 여기 추가하지 않으면
 * (derivedPath의 옛 삼항 체인처럼 조용히 엉뚱한 파일에 매핑되는 대신) 타입 오류로 바로 드러난다. */
const DERIVED_SUFFIX: Record<DerivedKind, string> = {
  pdf: 'pdf',
  grid: 'grid.json',
  markup: 'markup.json',
  pagesizes: 'pagesizes.json',
  gridindex: 'gridindex.json',
  svg: 'svg',
};
const DERIVED_KINDS = Object.keys(DERIVED_SUFFIX) as readonly DerivedKind[];

/**
 * 문서 저장 계층 추상화. disk(기본)와 memory 두 구현을 DOCVIEW_STORAGE로 선택한다.
 * 외부에서 함수 호출로 재사용할 수 있도록 파일 경로가 아닌 버퍼 기반 API를 노출한다.
 */
export interface DocumentStore {
  readonly mode: 'disk' | 'memory';
  saveOriginal(id: string, ext: string, data: Buffer): Promise<void>;
  readOriginal(id: string, ext: string): Promise<Buffer>;
  /** 원본 파일의 바이트 크기만 — 내용을 읽지 않고 알아야 할 때(예: 동시편집 룸 개수 상한 판단)
   * 쓴다. 파일이 없으면 null. */
  originalSize(id: string, ext: string): Promise<number | null>;
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

  // listMeta는 문서 목록 화면의 2초 폴링(main.js LIST_POLL_INTERVAL_MS)에 걸려 자주 불리는데,
  // 캐시가 없으면 매번 metaDir의 파일을 전부 다시 읽고 JSON.parse한다 — 문서가 많아질수록
  // 폴링 한 번의 비용이 문서 수에 비례해 커진다. writeMeta/deleteDocument가 이 클래스의 유일한
  // meta 쓰기 경로이므로 그 둘에서만 캐시를 갱신해두면, listMeta/readMeta는 이미 아는 id에 대해
  // 디스크를 다시 읽지 않아도 된다. 서버를 재시작한 직후처럼 캐시가 비어 있는 id는 readMeta가
  // 그대로 디스크에서 읽어 캐시를 채운다(cold-miss 폴백) — 이 클래스 밖에서 meta 파일을 직접
  // 건드리는 경로는 없으므로 캐시가 stale해질 일이 없다.
  private readonly metaCache = new Map<string, DocumentMeta>();

  private originalPath(id: string, ext: string): string {
    assertValidId(id);
    assertValidExt(ext);
    return path.join(config.originalsDir, `${id}.${ext}`);
  }

  private derivedPath(id: string, kind: DerivedKind): string {
    assertValidId(id);
    return path.join(config.convertedDir, `${id}.${DERIVED_SUFFIX[kind]}`);
  }

  private metaPath(id: string): string {
    assertValidId(id);
    return path.join(config.metaDir, `${id}.json`);
  }

  async saveOriginal(id: string, ext: string, data: Buffer): Promise<void> {
    // 원자적 쓰기: 같은 디렉터리에 임시 파일로 먼저 쓰고 rename으로 최종 경로에 옮긴다. rename은
    // 같은 파일시스템 안에서 원자적이라, 이 저장과 동시에 그 파일을 읽는(예: 다운로드 요청이 이
    // 파일을 unzip) 다른 요청은 항상 "이전의 완전한 파일" 또는 "새로운 완전한 파일"만 보게 되고,
    // 절대 쓰다 만 반쪽짜리 파일을 보지 않는다.
    const targetPath = this.originalPath(id, ext);
    const tempPath = `${targetPath}.tmp-${randomUUID()}`;
    await fs.writeFile(tempPath, data);
    await fs.rename(tempPath, targetPath);
  }

  async readOriginal(id: string, ext: string): Promise<Buffer> {
    return fs.readFile(this.originalPath(id, ext));
  }

  async originalSize(id: string, ext: string): Promise<number | null> {
    try {
      const stat = await fs.stat(this.originalPath(id, ext));
      return stat.size;
    } catch {
      return null;
    }
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
    this.metaCache.set(meta.id, meta);
  }

  async readMeta(id: string): Promise<DocumentMeta | null> {
    if (!UUID_PATTERN.test(id)) {
      return null;
    }
    const cached = this.metaCache.get(id);
    if (cached) return cached;
    try {
      const raw = await fs.readFile(this.metaPath(id), 'utf8');
      const meta = JSON.parse(raw) as DocumentMeta;
      this.metaCache.set(id, meta);
      return meta;
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
      ...DERIVED_KINDS.map((kind) => fs.rm(this.derivedPath(id, kind), { force: true })),
      fs.rm(this.metaPath(id), { force: true }),
    ]);
    this.metaCache.delete(id);
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
  // saveOriginal/saveDerived 한 번은 최대 6개의 derived 산출물을 만들 수 있어(DERIVED_KINDS),
  // 매번 이 값을 originals+derived의 모든 버퍼를 다시 합산해 구하면 한 번의 변환이 여러 번의
  // O(전체 버퍼 수) 스캔이 된다. 대신 넣고 뺄 때마다 이 값을 직접 갱신해 항상 O(1)로 읽는다.
  private currentBytes = 0;

  constructor(budgetBytes: number = config.memoryBudgetBytes) {
    this.budgetBytes = budgetBytes;
  }

  /** id로 시작하는 모든 파생 산출물을 kind를 하나하나 나열하지 않고 한 번에 지운다 — 새
   * DerivedKind를 추가하고 여기 추가를 깜빡해도 orphan 데이터가 남지 않는다. */
  private deleteDerivedFor(id: string): void {
    const prefix = `${id}#`;
    for (const [key, buf] of this.derived) {
      if (key.startsWith(prefix)) {
        this.derived.delete(key);
        this.currentBytes -= buf.length;
      }
    }
  }

  /** 예산 초과 시 보호 대상(protectedId)을 제외하고 가장 오래된 문서부터 제거한다. */
  private enforceBudget(protectedId: string): void {
    while (this.currentBytes > this.budgetBytes) {
      const oldest = sortByCreatedAtDesc([...this.metas.values()])
        .reverse()
        .find((meta) => meta.id !== protectedId);
      if (!oldest) return;
      const original = this.originals.get(oldest.id);
      if (original) this.currentBytes -= original.length;
      this.originals.delete(oldest.id);
      this.deleteDerivedFor(oldest.id);
      this.metas.delete(oldest.id);
      console.warn(`메모리 예산 초과로 오래된 문서를 제거했습니다: ${oldest.name} (${oldest.id})`);
    }
  }

  async saveOriginal(id: string, ext: string, data: Buffer): Promise<void> {
    assertValidId(id);
    assertValidExt(ext);
    const previous = this.originals.get(id);
    if (previous) this.currentBytes -= previous.length;
    this.originals.set(id, data);
    this.currentBytes += data.length;
    this.enforceBudget(id);
  }

  async readOriginal(id: string, _ext: string): Promise<Buffer> {
    const data = this.originals.get(id);
    if (!data) throw new Error('문서가 메모리에서 만료되었습니다. 다시 업로드해 주세요.');
    return data;
  }

  async originalSize(id: string, _ext: string): Promise<number | null> {
    return this.originals.get(id)?.length ?? null;
  }

  async saveDerived(id: string, kind: DerivedKind, data: Buffer): Promise<void> {
    assertValidId(id);
    const key = `${id}#${kind}`;
    const previous = this.derived.get(key);
    if (previous) this.currentBytes -= previous.length;
    this.derived.set(key, data);
    this.currentBytes += data.length;
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
    const original = this.originals.get(id);
    if (original) this.currentBytes -= original.length;
    this.originals.delete(id);
    this.deleteDerivedFor(id);
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
