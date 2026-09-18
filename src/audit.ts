import fs from 'node:fs/promises';
import path from 'node:path';
import { config } from './config.js';

export interface AuditEvent {
  readonly type: string;
  readonly docId: string;
  readonly viewer?: string;
  readonly detail?: string;
}

/** 날짜별로 파일을 나눠 무한정 자라는 단일 로그를 만들지 않는다 — 기존엔 audit.jsonl 하나에
 * 서버 수명 내내 계속 append만 해서, 되돌아가 읽는 코드가 전혀 없는데도 디스크가 끝없이 커졌다.
 * 오래된 날짜 파일은 배포 환경의 일반적인 로그 보관 정책(logrotate 등)으로 지우면 된다 — 이
 * 앱이 직접 지우지는 않는다(감사 로그를 앱이 스스로 삭제하는 건 별도의 보관 정책 결정이라
 * 여기서 임의로 정하지 않는다). */
function auditLogPathFor(date: Date): string {
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, '0');
  const d = String(date.getUTCDate()).padStart(2, '0');
  return path.join(config.auditDir, `audit-${y}-${m}-${d}.jsonl`);
}

/** 감사 로그 기록 실패가 본 요청을 실패시키지 않도록 내부에서 오류를 삼키되 콘솔에 남긴다. */
export async function recordAudit(event: AuditEvent): Promise<void> {
  const now = new Date();
  const line = JSON.stringify({ at: now.toISOString(), ...event });
  try {
    await fs.appendFile(auditLogPathFor(now), `${line}\n`, 'utf8');
  } catch (error) {
    console.error('감사 로그 기록 실패:', error);
  }
}
