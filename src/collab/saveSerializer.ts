/**
 * 같은 키(문서 ID)에 대한 비동기 작업이 서로 겹쳐 실행되지 않도록 한 줄로 세운다.
 *
 * room.ts는 같은 문서(docId)를 저장하는 지점이 셋 있다 — 유휴 자동저장 타이머, Ctrl+S 등
 * 즉시저장(flushNow), 마지막 참여자 퇴장 시 저장(closeConn). 이 셋은 서로 독립적으로
 * 트리거되므로 짧은 간격으로 겹쳐 호출될 수 있는데, 각각이 store.saveOriginal로 같은 파일을
 * 건드리는 read-modify-write라 겹치면 하나가 다른 하나의 결과를 덮어쓰는 lost-update가
 * 생긴다(디스크 쓰기 자체는 원자적 rename으로 파일이 깨지지는 않지만, "마지막에 rename한 쪽이
 * 이긴다"만 보장될 뿐 두 저장 중 하나의 내용이 통째로 사라질 수 있다). 이 헬퍼로 감싸면 두 번째
 * 호출은 첫 번째가 끝난 뒤에야 시작해, 항상 최신 Y.Doc 상태를 기준으로 다시 저장한다.
 */
const chains = new Map<string, Promise<void>>();

export function runSerialized<T>(key: string, task: () => Promise<T>): Promise<T> {
  const previous = chains.get(key) ?? Promise.resolve();
  // 이전 작업이 실패했어도(reject) 이번 작업은 그대로 시도한다 — 한 번의 저장 실패가 이후
  // 저장을 영원히 막아서는 안 된다.
  const result = previous.then(task, task);
  // 체인 자체(다음 호출이 기다릴 대상)는 절대 reject하지 않는 별도 promise로 추적한다 — 그래야
  // 이번 작업이 실패해도 다음 호출이 그 reject를 처리하지 않은 채(unhandledRejection) 이어받지
  // 않는다. 호출자에게 돌려주는 result는 실제 성공/실패를 그대로 담고 있다.
  const tracker: Promise<void> = result.then(
    () => undefined,
    () => undefined,
  );
  chains.set(key, tracker);
  // 마지막 대기자였다면(그 사이 새 호출이 없었다면) 정리해 키가 무한정 쌓이지 않게 한다.
  void tracker.then(() => {
    if (chains.get(key) === tracker) chains.delete(key);
  });
  return result;
}

/** 진단·테스트용 — 현재 대기/실행 중인 체인이 있는 키 개수 */
export function pendingSerializedChainCount(): number {
  return chains.size;
}
