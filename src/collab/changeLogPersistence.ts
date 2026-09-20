import fs from 'node:fs/promises';
import path from 'node:path';
import { config } from '../config.js';

/**
 * getChangeLog(room.ts)이 돌려주는 "룸이 열린 시점 대비 누적된 변경"은 메모리에만 있어 룸이 닫히면
 * (모든 편집자가 나가면) 사라진다 — 연동 측이 그 사이에 GET /api/documents/:id를 한 번도 안 불렀거나
 * 호출이 실패했다면 그 세션의 이력은 영영 조회할 수 없었다. 이 모듈은 룸이 "닫히는 바로 그 순간"에
 * (room.ts의 closeConn, room.conns.size === 0 분기) 그 시점까지 쌓인 항목 전체를 audit.ts와 같은
 * 방식(날짜별 JSONL, 실패해도 기록만 하고 삼킴)으로 디스크에 남겨 안전망 역할을 한다.
 *
 * **한계**: 룸이 정상적으로 닫힐 때만 기록한다 — 룸이 열려 있는 도중 서버가 죽거나 재시작되면(문서
 * 내용 자체는 별도의 자동저장으로 안전하지만) 그 세션의 변경 이력은 이 영속화 대상이 아니다.
 */
function changeLogPathFor(date: Date): string {
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, '0');
  const d = String(date.getUTCDate()).padStart(2, '0');
  return path.join(config.changeLogDir, `changes-${y}-${m}-${d}.jsonl`);
}

/** entries가 비어 있으면 아무것도 쓰지 않는다(편집 없이 열었다 닫힌 룸까지 매번 로그를 남기지 않는다).
 *
 * 래퍼 필드 이름을 `docKind`로 쓴 이유: pptx 항목(PptxChangeSummary)은 자신의 필드로 이미 `kind`
 * ('text'|'slide_added'|'slide_removed')를 갖고 있다 — `...entry`를 뒤에 스프레드하면 그 값이
 * 문서 형식 필드를 덮어써 버려("pptx"가 아니라 "text"로 기록됨), 처음엔 실제로 이 버그가 났다. */
export async function persistChangeLog(
  docId: string,
  docKind: 'xlsx' | 'docx' | 'pptx',
  entries: readonly object[],
): Promise<void> {
  if (entries.length === 0) return;
  const now = new Date();
  const closedAt = now.toISOString();
  const lines = entries.map((entry) => `${JSON.stringify({ docId, docKind, closedAt, ...entry })}\n`).join('');
  try {
    await fs.appendFile(changeLogPathFor(now), lines, 'utf8');
  } catch (error) {
    console.error(`변경 이력 영속화 실패 (문서 ${docId}):`, error);
  }
}
