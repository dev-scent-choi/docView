// 실행 중인 서버(포트 8098, DOCVIEW_DATA_DIR=%TEMP%\docview-test-data2, DOCVIEW_SECRET 고정값)를
// 대상으로 실제 WebsocketProvider(브라우저 번들과 동일 코드)로 두 "브라우저" 접속을 흉내 내
// 동시편집 → 저장까지 왕복시키는 수동 e2e 확인 스크립트. LibreOffice가 없는 이 환경에서는
// HTTP 업로드/토큰발급 API가 docx의 canEdit을 항상 꺼버리므로, 문서를 디스크에 직접 시딩하고
// JWT도 직접 서명해 WS 레이어만 검증한다.
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import jwt from 'jsonwebtoken';
import WebSocket from 'ws';
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';

const DATA_DIR = process.env.DOCVIEW_DATA_DIR;
const SECRET = process.env.DOCVIEW_SECRET;
const PORT = process.env.DOCVIEW_PORT ?? '8098';

function makeSampleDocx(paragraphs) {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const body = paragraphs
    .map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`)
    .join('');
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}<w:sectPr/></w:body></w:document>`;
  return Buffer.from(
    zipSync({
      '[Content_Types].xml': strToU8(contentTypes),
      '_rels/.rels': strToU8(rootRels),
      'word/document.xml': strToU8(documentXml),
    }),
  );
}

async function main() {
  const id = randomUUID();
  const now = new Date().toISOString();
  await fs.mkdir(path.join(DATA_DIR, 'originals'), { recursive: true });
  await fs.mkdir(path.join(DATA_DIR, 'meta'), { recursive: true });
  await fs.writeFile(path.join(DATA_DIR, 'originals', `${id}.docx`), makeSampleDocx(['Shared paragraph']));
  await fs.writeFile(
    path.join(DATA_DIR, 'meta', `${id}.json`),
    JSON.stringify({
      id,
      name: 'e2e-check.docx',
      ext: 'docx',
      status: 'ready',
      pageCount: 1,
      createdAt: now,
      viewKind: 'pages',
      progress: 100,
    }),
  );

  const sign = (viewer) =>
    jwt.sign(
      { docId: id, viewer, perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: true } },
      SECRET,
      { expiresIn: 60 },
    );

  const wsBase = `ws://127.0.0.1:${PORT}/ws/edit`;
  const ydocA = new Y.Doc();
  const ydocB = new Y.Doc();
  const providerA = new WebsocketProvider(wsBase, id, ydocA, {
    params: { token: sign('클라이언트A') },
    WebSocketPolyfill: WebSocket,
  });
  const providerB = new WebsocketProvider(wsBase, id, ydocB, {
    params: { token: sign('클라이언트B') },
    WebSocketPolyfill: WebSocket,
  });

  for (const [label, provider] of [['A', providerA], ['B', providerB]]) {
    provider.on('status', (e) => console.log(`[${label}] status:`, e.status));
    provider.on('connection-error', (e) => console.log(`[${label}] connection-error:`, e?.message ?? e));
    provider.on('connection-close', (e) => console.log(`[${label}] connection-close code:`, e?.code, e?.reason));
  }

  const waitSynced = (label, provider) =>
    new Promise((resolve, reject) => {
      const timeout = setTimeout(() => reject(new Error(`sync timeout (${label})`)), 15000);
      provider.on('synced', () => {
        clearTimeout(timeout);
        resolve();
      });
    });

  await Promise.all([waitSynced('A', providerA), waitSynced('B', providerB)]);
  console.log('both clients synced');

  const fragA = ydocA.getXmlFragment('prosemirror');
  const fragB = ydocB.getXmlFragment('prosemirror');
  const textA = fragA.get(0).get(0);
  const textB = fragB.get(0).get(0);
  console.log('seeded text (A view):', textA.toString());

  textA.insert(textA.length, ' [from A]');
  textB.insert(textB.length, ' [from B]');

  // 두 변경이 서로에게 전파될 시간을 준다
  await new Promise((resolve) => setTimeout(resolve, 1500));
  console.log('A sees:', fragA.get(0).get(0).toString());
  console.log('B sees:', fragB.get(0).get(0).toString());

  providerA.destroy();
  providerB.destroy();

  // 마지막 연결 종료 시 즉시 flush되므로 잠깐 기다렸다가 저장된 docx를 확인한다
  await new Promise((resolve) => setTimeout(resolve, 1500));
  const saved = await fs.readFile(path.join(DATA_DIR, 'originals', `${id}.docx`));
  const savedXmlEntry = zipEntryText(saved, 'word/document.xml');
  const hasA = savedXmlEntry.includes('[from A]');
  const hasB = savedXmlEntry.includes('[from B]');
  console.log('saved docx contains [from A]:', hasA);
  console.log('saved docx contains [from B]:', hasB);

  const metaAfter = JSON.parse(await fs.readFile(path.join(DATA_DIR, 'meta', `${id}.json`), 'utf8'));
  console.log('meta after save (LibreOffice absent so PDF regen is skipped, status/pageCount unchanged):', metaAfter.status, metaAfter.pageCount);

  if (!hasA || !hasB) {
    console.error('FAIL: concurrent edits were not both persisted');
    process.exit(1);
  }
  console.log('PASS: concurrent edits from two independent WebsocketProvider clients merged and persisted to docx');
}

function zipEntryText(zipBuffer, entryName) {
  return strFromU8(unzipSync(zipBuffer)[entryName]);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
