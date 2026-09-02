import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

async function freshPresenceModule() {
  vi.resetModules();
  return import('../src/presence.js');
}

describe('presence', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  test('listPresence on a document with no heartbeats returns an empty array', async () => {
    const { listPresence } = await freshPresenceModule();
    expect(listPresence('doc-1')).toEqual([]);
  });

  test('touchPresence adds the viewer and returns it in the list', async () => {
    const { touchPresence } = await freshPresenceModule();
    const viewers = touchPresence('doc-1', 'session-a', '검토자A');
    expect(viewers).toEqual(['검토자A']);
  });

  test('two different sessions with different viewer names both appear', async () => {
    const { touchPresence } = await freshPresenceModule();
    touchPresence('doc-1', 'session-a', '검토자A');
    const viewers = touchPresence('doc-1', 'session-b', '검토자B');
    expect([...viewers].sort()).toEqual(['검토자A', '검토자B']);
  });

  test('the same viewer name from two sessions (two tabs) appears only once', async () => {
    const { touchPresence } = await freshPresenceModule();
    touchPresence('doc-1', 'session-a', '검토자A');
    const viewers = touchPresence('doc-1', 'session-b', '검토자A');
    expect(viewers).toEqual(['검토자A']);
  });

  test('presence for one document does not leak into another', async () => {
    const { touchPresence, listPresence } = await freshPresenceModule();
    touchPresence('doc-1', 'session-a', '검토자A');
    expect(listPresence('doc-2')).toEqual([]);
  });

  test('a viewer disappears from the list after their heartbeat times out', async () => {
    const { touchPresence, listPresence } = await freshPresenceModule();
    touchPresence('doc-1', 'session-a', '검토자A');
    vi.advanceTimersByTime(30_000);
    expect(listPresence('doc-1')).toEqual([]);
  });

  test('a fresh heartbeat keeps a viewer alive past when an old one would have expired', async () => {
    const { touchPresence, listPresence } = await freshPresenceModule();
    touchPresence('doc-1', 'session-a', '검토자A');
    vi.advanceTimersByTime(15_000);
    touchPresence('doc-1', 'session-a', '검토자A');
    vi.advanceTimersByTime(15_000);
    expect(listPresence('doc-1')).toEqual(['검토자A']);
  });

  test('when one of two tabs for the same viewer times out but the other is still active, the viewer stays listed', async () => {
    const { touchPresence, listPresence } = await freshPresenceModule();
    touchPresence('doc-1', 'session-a', '검토자A');
    vi.advanceTimersByTime(15_000);
    touchPresence('doc-1', 'session-b', '검토자A'); // 같은 이름, 새 탭
    vi.advanceTimersByTime(10_000); // session-a는 25초 경과(만료), session-b는 10초 경과(생존)
    expect(listPresence('doc-1')).toEqual(['검토자A']);
  });
});
