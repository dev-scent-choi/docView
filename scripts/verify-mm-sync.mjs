// maintenance-management 연동 검증: webviewer 임베드 URL을 직접 열어 셀 하나를 편집하고,
// blur(자동저장 디바운스) 후 maintenance-management의 GET /connection-status/{year}가 그 값을
// Postgres로 동기화해오는지 확인한다.
import { chromium } from 'playwright';

const embedUrl = process.argv[2];
if (!embedUrl) {
  console.error('usage: node verify-mm-sync.mjs <full-editor-url>');
  process.exit(1);
}

const browser = await chromium.launch();
const page = await browser.newPage();
page.on('pageerror', (err) => console.log('PAGEERROR:', err.message));

await page.goto(embedUrl, { waitUntil: 'domcontentloaded' });
await page.waitForSelector('#grid-edit-table td[contenteditable]', { timeout: 20000 });

// row=2(AlphaTest Corp), col=2(1월) — POI 0-based로 두 번째 데이터 행, 첫 번째 월 컬럼
const cell = page.locator('#grid-edit-table td[data-row="2"][data-col="2"]');
await cell.click();
await page.keyboard.press('Control+A');
await page.keyboard.type('7');
await cell.blur();
await page.waitForFunction(
  () => document.querySelector('#grid-edit-table td[data-row="2"][data-col="2"]')?.textContent === '7',
  { timeout: 5000 },
);
console.log('OK: 셀에 7 입력 확인');

// 유휴 디바운스 저장을 기다린다(room.ts의 AUTOSAVE_IDLE_MS=8000 + 여유)
await page.waitForTimeout(11000);

await browser.close();
console.log('DONE');
