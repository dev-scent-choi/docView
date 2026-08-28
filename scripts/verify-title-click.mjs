// 제목(병합/잠금) 셀 클릭 시 postMessage에 실제 화면 좌표(rect)가 같이 오는지, 그리고
// 1~12월 칸에 숫자 아닌 값을 넣으면 저장이 거부되고 안내 말풍선이 뜨는지 확인한다.
import { chromium } from 'playwright';

const embedUrl = process.argv[2];
if (!embedUrl) {
  console.error('usage: node verify-title-click.mjs <full-editor-url>');
  process.exit(1);
}
const origin = new URL(embedUrl).origin;

const browser = await chromium.launch();
const page = await browser.newPage();

const hostHtml = `<!doctype html><html><body>
<div style="margin:40px">
<iframe id="f" src="${embedUrl}" style="width:1000px;height:700px;border:0"></iframe>
</div>
<script>
  window.__received = null;
  window.addEventListener('message', (e) => {
    if (e.origin !== ${JSON.stringify(origin)}) return;
    window.__received = e.data;
  });
</script>
</body></html>`;

await page.setContent(hostHtml);
await page.waitForSelector('#f');
const frame = page.frameLocator('#f');
await frame.locator('#grid-edit-table td[data-row="0"][data-col="1"]').waitFor({ timeout: 20000 });

// 1) 제목 클릭 -> rect 포함 postMessage 확인
await frame.locator('#grid-edit-table td[data-row="0"][data-col="1"]').click();
await page.waitForTimeout(300);
const received = await page.evaluate(() => window.__received);
console.log('received:', JSON.stringify(received));
const hasRect = received?.type === 'docview:title-click' && received.rect && typeof received.rect.bottom === 'number';
console.log(hasRect ? 'OK: postMessage에 rect(화면 좌표) 포함됨' : 'FAIL: rect 없음');

// 2) 1월 칸(row=2,col=2)에 문자 입력 -> 거부되고 안내 말풍선 표시
const monthCell = frame.locator('#grid-edit-table td[data-row="2"][data-col="2"]');
await monthCell.click();
await page.keyboard.press('Control+A');
await page.keyboard.type('abc');
await monthCell.blur();
await page.waitForTimeout(300);

const tooltipVisible = await frame.locator('.g-cell-tooltip').count();
console.log(tooltipVisible > 0 ? 'OK: 숫자 아닌 입력 시 안내 말풍선 표시됨' : 'FAIL: 안내 말풍선 없음');

const cellText = await monthCell.textContent();
console.log('month cell after invalid input:', JSON.stringify(cellText));
console.log(cellText !== 'abc' ? 'OK: 잘못된 값이 저장되지 않고 원래 값으로 복귀함' : 'FAIL: 잘못된 값이 저장됨');

// 3) 이번엔 숫자를 넣으면 정상 저장되는지 확인
await monthCell.click();
await page.keyboard.press('Control+A');
await page.keyboard.type('42');
await monthCell.blur();
await page.waitForFunction(
  () => document.querySelector('#f').contentDocument.querySelector('#grid-edit-table td[data-row="2"][data-col="2"]')?.textContent === '42',
  { timeout: 5000 },
).catch(() => {});
const afterValid = await monthCell.textContent();
console.log(afterValid === '42' ? 'OK: 숫자 입력은 정상 저장됨 (42)' : `FAIL: 숫자 입력 저장 실패 (got ${JSON.stringify(afterValid)})`);

await browser.close();
