// fast-formula-parser로 교체 + formulaOverrides.ts로 패치한 수식 엔진의 실제 동작을 확인한다.
// VLOOKUP/MATCH/INDEX/MAX/MIN/CHOOSE/SWITCH/SUMIFS/AVERAGEIFS + 중첩 IF+MAX까지, "자주 쓰는" 수식이
// 실시간으로 올바른 계산 결과를 보여주는지 하나씩 검증한다.
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { chromium } from 'playwright';
import jwt from 'jsonwebtoken';
import ExcelJS from 'exceljs';

const BASE = `http://127.0.0.1:${process.env.DOCVIEW_PORT}`;
const DATA_DIR = process.env.DOCVIEW_DATA_DIR;
const SECRET = process.env.DOCVIEW_SECRET;

async function seedEditableXlsx() {
  const id = randomUUID();
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'apple';
  sheet.getCell('B1').value = 10;
  sheet.getCell('A2').value = 'banana';
  sheet.getCell('B2').value = 20;
  sheet.getCell('A3').value = 'cherry';
  sheet.getCell('B3').value = 30;
  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());

  await fs.mkdir(path.join(DATA_DIR, 'originals'), { recursive: true });
  await fs.mkdir(path.join(DATA_DIR, 'meta'), { recursive: true });
  await fs.writeFile(path.join(DATA_DIR, 'originals', `${id}.xlsx`), buffer);
  await fs.writeFile(
    path.join(DATA_DIR, 'meta', `${id}.json`),
    JSON.stringify({
      id,
      name: 'formula-check.xlsx',
      ext: 'xlsx',
      status: 'ready',
      pageCount: 1,
      createdAt: new Date().toISOString(),
      viewKind: 'grid',
      progress: 100,
    }),
  );
  const token = jwt.sign(
    { docId: id, viewer: '수식확인', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: true } },
    SECRET,
    { expiresIn: 300 },
  );
  return token;
}

const CASES = [
  { row: 0, col: 3, formula: '=VLOOKUP("banana",A1:B3,2,FALSE)', expected: '20' },
  { row: 1, col: 3, formula: '=MATCH("cherry",A1:A3,0)', expected: '3' },
  { row: 2, col: 3, formula: '=INDEX(A1:B3,2,2)', expected: '20' },
  { row: 3, col: 3, formula: '=MAX(B1:B3)', expected: '30' },
  { row: 4, col: 3, formula: '=MIN(B1:B3)', expected: '10' },
  { row: 5, col: 3, formula: '=CHOOSE(2,"x","y","z")', expected: 'y' },
  { row: 6, col: 3, formula: '=SWITCH(2,1,"one",2,"two",3,"three")', expected: 'two' },
  { row: 7, col: 3, formula: '=SUMIFS(B1:B3,A1:A3,"banana")', expected: '20' },
  { row: 8, col: 3, formula: '=AVERAGEIFS(B1:B3,A1:A3,"banana")', expected: '20' },
  { row: 9, col: 3, formula: '=IF(MAX(B1:B3)>25,"big","small")', expected: 'big' },
];

async function setCellAndCheck(page, { row, col, formula, expected }) {
  const cell = page.locator(`#grid-edit-table td[data-row="${row}"][data-col="${col}"]`);
  await cell.click();
  await page.keyboard.press('Control+A');
  await page.keyboard.type(formula);
  await cell.blur();
  try {
    await page.waitForFunction(
      ({ row, col, expected }) =>
        document.querySelector(`#grid-edit-table td[data-row="${row}"][data-col="${col}"]`)?.textContent === expected,
      { row, col, expected },
      { timeout: 5000 },
    );
    console.log(`OK: ${formula} -> ${expected}`);
    return true;
  } catch {
    const actual = await cell.textContent();
    console.log(`FAIL: ${formula} -> expected "${expected}", got "${actual}"`);
    return false;
  }
}

async function main() {
  const browser = await chromium.launch();
  const context = await browser.newContext();
  const page = await context.newPage();
  page.on('pageerror', (err) => console.log('PAGEERROR:', err.message));

  const token = await seedEditableXlsx();
  await page.goto(`${BASE}/editor-grid.html?token=${encodeURIComponent(token)}`, { waitUntil: 'domcontentloaded' });
  await page.waitForSelector('#grid-edit-table td[contenteditable]', { timeout: 20000 });

  let passed = 0;
  for (const testCase of CASES) {
    const ok = await setCellAndCheck(page, testCase);
    if (ok) passed += 1;
  }

  console.log(`\n${passed}/${CASES.length} 수식 통과`);
  await browser.close();
  if (passed !== CASES.length) process.exit(1);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
