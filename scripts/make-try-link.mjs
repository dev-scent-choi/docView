// 사용자가 직접 브라우저로 열어볼 수 있는 동시편집 링크를 발급한다 (docx를 디스크에 직접
// 시딩 — 이 환경엔 LibreOffice가 없어 업로드 UI로는 docx를 넣을 수 없어서).
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { strToU8, zipSync } from 'fflate';
import jwt from 'jsonwebtoken';

const DATA_DIR = process.env.DOCVIEW_DATA_DIR;
const SECRET = process.env.DOCVIEW_SECRET;
const PORT = process.env.DOCVIEW_PORT ?? '8090';

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

const id = randomUUID();
await fs.mkdir(path.join(DATA_DIR, 'originals'), { recursive: true });
await fs.mkdir(path.join(DATA_DIR, 'meta'), { recursive: true });
await fs.writeFile(
  path.join(DATA_DIR, 'originals', `${id}.docx`),
  makeSampleDocx(['이 문서는 실시간 동시편집을 시험해보기 위한 샘플입니다.', '이 링크를 두 개의 탭(또는 창)에서 열어 보세요.']),
);
await fs.writeFile(
  path.join(DATA_DIR, 'meta', `${id}.json`),
  JSON.stringify({
    id,
    name: '동시편집-체험.docx',
    ext: 'docx',
    status: 'ready',
    pageCount: 1,
    createdAt: new Date().toISOString(),
    viewKind: 'pages',
    progress: 100,
  }),
);

for (const viewer of ['사용자A', '사용자B']) {
  const token = jwt.sign(
    { docId: id, viewer, perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: true } },
    SECRET,
    { expiresIn: 1800 },
  );
  console.log(`${viewer}: http://127.0.0.1:${PORT}/editor.html?token=${encodeURIComponent(token)}`);
}
