import fs from 'node:fs/promises';
import path from 'node:path';
import { config } from './config.js';

export interface AuditEvent {
  readonly type: string;
  readonly docId: string;
  readonly viewer?: string;
  readonly detail?: string;
}

/** 감사 로그 기록 실패가 본 요청을 실패시키지 않도록 내부에서 오류를 삼키되 콘솔에 남긴다. */
export async function recordAudit(event: AuditEvent): Promise<void> {
  const line = JSON.stringify({ at: new Date().toISOString(), ...event });
  try {
    await fs.appendFile(path.join(config.auditDir, 'audit.jsonl'), `${line}\n`, 'utf8');
  } catch (error) {
    console.error('감사 로그 기록 실패:', error);
  }
}
