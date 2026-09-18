#!/usr/bin/env node
/**
 * public/rhwp-studio/ 아래에 rhwp-studio(https://edwardkim.github.io/rhwp/)의 빌드 산출물을
 * 그대로 미러링한다. 이 프로젝트는 Rust/wasm-pack 툴체인이 필요해 로컬에서 다시 빌드하는 대신,
 * 이미 배포된 정적 파일을 그대로 받아 우리 서버에서 서빙한다(iframe이 우리 origin에서만
 * 로드되게 하기 위함 — 문서 내용이 외부 사이트로 나가지 않는다).
 *
 * rhwp가 새 버전을 배포할 때마다 다시 실행해서 갱신해야 한다:
 *   node scripts/mirror-rhwp-studio.mjs
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const SOURCE_BASE = 'https://edwardkim.github.io/rhwp/';
// rhwp-studio 빌드가 절대경로 `/rhwp/...`로 자기 자신을 참조하므로(--base=/rhwp/로 빌드됨),
// 로컬 폴더명도 그 경로 그대로 맞춰야 한다 — 그래야 기존 app.ts의 공용 정적 파일 서버
// (public/ 전체를 루트로 서빙)가 별도 라우트 등록 없이 /rhwp/... 그대로 서빙해 준다.
const OUT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'public', 'rhwp');

const visited = new Set();
const queue = ['/rhwp/', '/rhwp/index.html'];

function isMirrorablePath(p) {
  return p.startsWith('/rhwp/');
}

/** HTML/CSS/JS 텍스트에서 /rhwp/로 시작하는 경로 참조를 전부 찾는다 — Vite 번들은 청크·wasm을
 * 대부분 문자열 리터럴로 참조하므로 정규식 스캔으로 충분하다(완전한 JS 파서 불필요). */
function extractReferences(text) {
  const refs = new Set();
  const re = /["'`(]\s*(\/rhwp\/[A-Za-z0-9_\-./%]+)/g;
  let m;
  while ((m = re.exec(text))) {
    refs.add(m[1].split('?')[0].split('#')[0]);
  }
  // 폰트 파일은 절대경로(/rhwp/...)가 아니라 base URL 기준 상대경로("fonts/NotoSansKR-*.woff2")로
  // 런타임에 조립된다(new URL(relative, import.meta.env.BASE_URL) 패턴으로 추정) — 위 정규식이
  // 잡는 절대경로 문자열 리터럴이 아예 없어서 별도로 스캔해야 한다.
  const fontRe = /fonts\/[A-Za-z0-9_.-]+\.(?:woff2?|ttf|otf)/g;
  while ((m = fontRe.exec(text))) {
    refs.add(`/rhwp/${m[0]}`);
  }
  return refs;
}

async function fetchOne(urlPath) {
  const url = new URL(urlPath.replace(/^\/rhwp\//, ''), SOURCE_BASE).toString();
  const res = await fetch(url);
  if (!res.ok) {
    console.warn(`  스킵(${res.status}): ${urlPath}`);
    return null;
  }
  const buf = Buffer.from(await res.arrayBuffer());
  const destPath = path.join(OUT_DIR, urlPath.replace(/^\/rhwp\//, ''));
  await fs.mkdir(path.dirname(destPath), { recursive: true });
  await fs.writeFile(destPath, buf);
  return buf;
}

async function main() {
  await fs.rm(OUT_DIR, { recursive: true, force: true });
  let count = 0;
  while (queue.length > 0) {
    const p = queue.shift().split('?')[0].split('#')[0];
    if (visited.has(p) || !isMirrorablePath(p)) continue;
    visited.add(p);
    const normalized = p.endsWith('/') ? `${p}index.html` : p;
    const buf = await fetchOne(normalized);
    if (!buf) continue;
    count += 1;
    console.log(`받음: ${normalized} (${buf.length} bytes)`);
    if (/\.(html|js|css|webmanifest|json)$/.test(normalized)) {
      const text = buf.toString('utf-8');
      for (const ref of extractReferences(text)) {
        if (!visited.has(ref)) queue.push(ref);
      }
    }
  }
  console.log(`\n총 ${count}개 파일을 ${OUT_DIR}에 미러링했습니다.`);
}

main().catch((error) => {
  console.error('미러링 실패:', error);
  process.exit(1);
});
