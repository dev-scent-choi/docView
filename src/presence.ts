/**
 * "지금 이 문서를 보고 있는 사람" 표시 — docx/xlsx 동시편집(Yjs awareness)과 달리 뷰어(viewer.html)는
 * 실시간 WebSocket 연결이 없으므로, 클라이언트가 주기적으로 보내는 하트비트를 서버가 메모리에서
 * 추적하는 가벼운 폴링 방식을 쓴다. pages/grid 어느 viewKind든 viewer.html을 통하는 모든 문서에
 * 동일하게 적용된다.
 */

/** 하트비트 주기(클라이언트)의 2배 이상 — 네트워크 지연으로 한 번 놓쳐도 바로 사라지지 않게 여유를 둔다 */
const HEARTBEAT_TIMEOUT_MS = 20_000;

interface PresenceEntry {
  readonly viewer: string;
  lastSeen: number;
}

/** docId -> sessionId -> entry. sessionId는 클라이언트가 탭/세션마다 하나씩 발급한다(같은 이름이
 * 여러 탭으로 들어와도 세션별로 독립적으로 만료되도록 — 한쪽 탭을 닫아도 다른 탭이 살아있으면
 * 목록에서 사라지지 않는다). */
const byDoc = new Map<string, Map<string, PresenceEntry>>();

function pruneExpired(docId: string): Map<string, PresenceEntry> | undefined {
  const sessions = byDoc.get(docId);
  if (!sessions) return undefined;
  const now = Date.now();
  for (const [sessionId, entry] of sessions) {
    if (now - entry.lastSeen > HEARTBEAT_TIMEOUT_MS) sessions.delete(sessionId);
  }
  if (sessions.size === 0) {
    byDoc.delete(docId);
    return undefined;
  }
  return sessions;
}

/** 만료된 세션을 걷어내고, 이름 기준으로 중복 없이(같은 사람 여러 탭 = 한 번) 정렬 없는 목록을 돌려준다. */
export function listPresence(docId: string): readonly string[] {
  const sessions = pruneExpired(docId);
  if (!sessions) return [];
  return [...new Set([...sessions.values()].map((entry) => entry.viewer))];
}

/** 하트비트를 기록하고, 갱신된 현재 목록을 바로 돌려준다(클라이언트가 별도 조회 없이 응답만으로 갱신). */
export function touchPresence(docId: string, sessionId: string, viewer: string): readonly string[] {
  let sessions = byDoc.get(docId);
  if (!sessions) {
    sessions = new Map();
    byDoc.set(docId, sessions);
  }
  sessions.set(sessionId, { viewer, lastSeen: Date.now() });
  return listPresence(docId);
}

/** 문서가 삭제될 때 부른다 — pruneExpired는 그 문서를 다음에 다시 조회할 때만 정리하는데, 삭제된
 * 문서는 다시 조회될 일이 없어(뷰어를 열 수 없으므로) 만료된 세션 Map이 서버가 살아있는 동안
 * 영원히 byDoc에 남는다(항목 자체는 작지만, 오래 켜둔 서버에서 조금씩 쌓이는 누수). */
export function clearPresence(docId: string): void {
  byDoc.delete(docId);
}
