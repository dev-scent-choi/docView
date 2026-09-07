import { describe, expect, test } from 'vitest';
import { pendingSerializedChainCount, runSerialized } from '../src/collab/saveSerializer.js';

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

describe('saveSerializer: runSerialized', () => {
  test('two calls with the same key never run concurrently', async () => {
    const key = `doc-${Math.random()}`;
    let inFlight = 0;
    let maxInFlight = 0;
    const order: number[] = [];

    async function task(id: number, ms: number): Promise<void> {
      inFlight += 1;
      maxInFlight = Math.max(maxInFlight, inFlight);
      await delay(ms);
      order.push(id);
      inFlight -= 1;
    }

    // 두 번째 호출이 먼저 끝날 조건(더 짧은 delay)을 줘도, 직렬화되어 있으면 첫 번째가
    // 시작해 끝난 뒤에야 두 번째가 시작하므로 완료 순서는 여전히 [1, 2]다.
    const p1 = runSerialized(key, () => task(1, 30));
    const p2 = runSerialized(key, () => task(2, 5));
    await Promise.all([p1, p2]);

    expect(maxInFlight).toBe(1);
    expect(order).toEqual([1, 2]);
  });

  test('calls with different keys run concurrently, unaffected by each other', async () => {
    let inFlightA = 0;
    let inFlightBOverlappedA = false;

    const pA = runSerialized(`doc-a-${Math.random()}`, async () => {
      inFlightA += 1;
      await delay(30);
      inFlightA -= 1;
    });
    const pB = runSerialized(`doc-b-${Math.random()}`, async () => {
      if (inFlightA > 0) inFlightBOverlappedA = true;
      await delay(5);
    });

    await Promise.all([pA, pB]);
    expect(inFlightBOverlappedA).toBe(true);
  });

  test('a failing task does not block the next task for the same key', async () => {
    const key = `doc-${Math.random()}`;
    const first = runSerialized(key, async () => {
      throw new Error('save failed');
    });
    await expect(first).rejects.toThrow('save failed');

    let secondRan = false;
    const second = runSerialized(key, async () => {
      secondRan = true;
      return 'ok';
    });
    await expect(second).resolves.toBe('ok');
    expect(secondRan).toBe(true);
  });

  test('each caller gets back its own task result, not swallowed by the chain', async () => {
    const key = `doc-${Math.random()}`;
    const results = await Promise.all([
      runSerialized(key, async () => 'first'),
      runSerialized(key, async () => 'second'),
      runSerialized(key, async () => 'third'),
    ]);
    expect(results).toEqual(['first', 'second', 'third']);
  });

  test('cleans up its internal chain entry once the last queued task settles', async () => {
    const key = `doc-${Math.random()}`;
    await runSerialized(key, async () => delay(1));
    // 마지막 대기자가 끝난 뒤 정리되므로, 이미 끝난 지 한 틱 지난 지금은 이 키의 체인이 남아있지
    // 않아야 한다(무한정 쌓이는 메모리 누수를 막기 위한 정리 로직 검증).
    await delay(1);
    const countBefore = pendingSerializedChainCount();
    await runSerialized(`other-key-${Math.random()}`, async () => undefined);
    await delay(1);
    expect(pendingSerializedChainCount()).toBeLessThanOrEqual(countBefore);
  });
});
