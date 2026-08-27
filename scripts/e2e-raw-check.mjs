// 프로토콜 레벨 원시 확인: y-websocket의 WebsocketProvider를 거치지 않고 직접 메시지를
// 주고받아 서버가 무엇을 보내는지 정확히 관찰한다.
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { strToU8, zipSync } from 'fflate';
import jwt from 'jsonwebtoken';
import WebSocket from 'ws';
import * as encoding from 'lib0/encoding';
import * as decoding from 'lib0/decoding';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';

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
  const body = paragraphs.map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join('');
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
  await fs.mkdir(path.join(DATA_DIR, 'originals'), { recursive: true });
  await fs.mkdir(path.join(DATA_DIR, 'meta'), { recursive: true });
  await fs.writeFile(path.join(DATA_DIR, 'originals', `${id}.docx`), makeSampleDocx(['Hello']));
  await fs.writeFile(
    path.join(DATA_DIR, 'meta', `${id}.json`),
    JSON.stringify({
      id,
      name: 'raw-check.docx',
      ext: 'docx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      viewKind: 'pages',
      progress: 100,
    }),
  );
  const token = jwt.sign(
    { docId: id, viewer: 'raw', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: true } },
    SECRET,
    { expiresIn: 60 },
  );

  const url = `ws://127.0.0.1:${PORT}/ws/edit/${id}?token=${encodeURIComponent(token)}`;
  console.log('connecting to', url);
  const ws = new WebSocket(url);
  ws.binaryType = 'arraybuffer';

  ws.on('open', () => {
    console.log('raw ws open, readyState=', ws.readyState);
    const ydoc = new Y.Doc();
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, 0); // messageSync
    syncProtocol.writeSyncStep1(encoder, ydoc);
    const bytes = encoding.toUint8Array(encoder);
    console.log('sending syncStep1, bytes:', bytes.length, [...bytes]);
    ws.send(bytes);
  });

  ws.on('message', (data, isBinary) => {
    const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data);
    console.log('RECEIVED message, isBinary=', isBinary, 'len=', bytes.length, 'first bytes:', [...bytes.slice(0, 10)]);
    const decoder = decoding.createDecoder(bytes);
    const messageType = decoding.readVarUint(decoder);
    console.log('  messageType:', messageType);
    if (messageType === 0) {
      const syncType = decoding.readVarUint(decoder);
      console.log('  syncMessageType:', syncType, '(0=step1,1=step2,2=update)');
    }
  });

  ws.on('error', (err) => console.log('raw ws error:', err.message));
  ws.on('close', (code, reason) => console.log('raw ws closed:', code, reason?.toString()));

  await new Promise((resolve) => setTimeout(resolve, 5000));
  ws.close();
  await new Promise((resolve) => setTimeout(resolve, 500));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
