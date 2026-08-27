/**
 * 총 바이트 예산 기반 LRU 캐시.
 * 개수만 제한하던 기존 방식은 항목 크기가 제각각(수십 MB 그리드, 고배율 PDF 등)이라
 * 대형 문서가 겹치면 메모리가 사실상 무제한으로 쌓였다. 이 캐시는 항목별 크기를
 * 측정해 총량이 예산을 넘으면 가장 오래 쓰지 않은 항목부터 제거한다.
 *
 * Promise를 저장하되 resolve된 뒤 크기를 측정한다(로딩 중 항목은 0으로 계산).
 */
export interface SizedLruCacheOptions<V> {
  readonly maxBytes: number;
  readonly maxEntries?: number;
  readonly sizeOf: (value: V) => number;
}

export class SizedLruCache<V> {
  private readonly entries = new Map<string, Promise<V>>();
  private readonly sizes = new Map<string, number>();
  private totalBytes = 0;

  constructor(private readonly options: SizedLruCacheOptions<V>) {}

  get(key: string): Promise<V> | undefined {
    const value = this.entries.get(key);
    if (value === undefined) return undefined;
    // LRU 갱신: 최근 사용 항목을 맨 뒤로 이동한다
    this.entries.delete(key);
    this.entries.set(key, value);
    return value;
  }

  /** 없으면 loader로 만들어 저장하고, 있으면 기존 Promise를 재사용한다. */
  getOrLoad(key: string, loader: () => Promise<V>): Promise<V> {
    const existing = this.get(key);
    if (existing) return existing;

    const promise = loader();
    this.entries.set(key, promise);
    this.sizes.set(key, 0);
    promise.then(
      (value) => {
        if (this.entries.get(key) !== promise) return; // 이미 교체·제거됨
        const size = Math.max(0, this.options.sizeOf(value));
        this.totalBytes += size - (this.sizes.get(key) ?? 0);
        this.sizes.set(key, size);
        this.evict(key);
      },
      () => this.remove(key),
    );
    this.evict(key);
    return promise;
  }

  private remove(key: string): void {
    if (!this.entries.has(key)) return;
    this.entries.delete(key);
    this.totalBytes -= this.sizes.get(key) ?? 0;
    this.sizes.delete(key);
  }

  /** 특정 키를 강제로 제거한다 (예: 문서 재저장으로 캐시된 렌더링 결과가 stale해졌을 때). */
  invalidate(key: string): void {
    this.remove(key);
  }

  /** prefix로 시작하는 모든 키를 제거한다 (예: `${docId}#scale` 형태로 키가 파생되는 캐시). */
  invalidatePrefix(prefix: string): void {
    for (const key of [...this.entries.keys()]) {
      if (key.startsWith(prefix)) this.remove(key);
    }
  }

  /** 예산·개수 초과 시 protectedKey를 제외하고 오래된 항목부터 제거한다. */
  private evict(protectedKey: string): void {
    const maxEntries = this.options.maxEntries ?? Number.POSITIVE_INFINITY;
    for (const key of [...this.entries.keys()]) {
      if (this.totalBytes <= this.options.maxBytes && this.entries.size <= maxEntries) return;
      if (this.entries.size <= 1) return; // 방금 넣은 항목 하나는 보존
      if (key === protectedKey) continue;
      this.remove(key);
    }
  }
}
