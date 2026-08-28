import { chromium } from 'playwright';

const embedUrl = process.argv[2];
const origin = new URL(embedUrl).origin;
const browser = await chromium.launch();
const page = await browser.newPage();

const hostHtml = `<!doctype html><html><body>
<div style="margin:40px"><iframe id="f" src="${embedUrl}" style="width:1000px;height:700px;border:0"></iframe></div>
<script>
  window.__received = null;
  window.addEventListener('message', (e) => { if (e.origin === ${JSON.stringify(origin)}) window.__received = e.data; });
</script>
</body></html>`;
await page.setContent(hostHtml);
const frame = page.frameLocator('#f');
await frame.locator('#grid-edit-table td[data-row="0"][data-col="0"]').waitFor({ timeout: 20000 });

// lock check with new column positions: header row=1, name/header col=0, month cols=1..12, total col=13
const lockInfo = await frame.locator('body').evaluate(() => {
  const get = (r, c) => document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`);
  return {
    header: get(1, 0)?.getAttribute('contenteditable'),
    name: get(2, 0)?.getAttribute('contenteditable'),
    month: get(2, 1)?.getAttribute('contenteditable'),
    total: get(2, 13)?.getAttribute('contenteditable'),
  };
});
console.log('lock:', JSON.stringify(lockInfo));
console.log(
  lockInfo.header === 'false' && lockInfo.name === 'false' && lockInfo.month === 'true' && lockInfo.total === 'false'
    ? 'OK: 잠금 상태 정상(헤더/이름/합계 잠김, 월 값만 편집 가능) - 새 열 위치 기준'
    : 'FAIL: 잠금 상태 문제',
);

// numeric-only check
const monthCell = frame.locator('#grid-edit-table td[data-row="2"][data-col="1"]');
await monthCell.click();
await page.keyboard.press('Control+A');
await page.keyboard.type('xyz');
await monthCell.blur();
await page.waitForTimeout(300);
const tipCount = await frame.locator('.g-cell-tooltip').count();
const text1 = await monthCell.textContent();
console.log(tipCount > 0 && text1 !== 'xyz' ? 'OK: 숫자 아닌 입력 거부 + 안내' : 'FAIL: 숫자 검증 문제');

await monthCell.click();
await page.keyboard.press('Control+A');
await page.keyboard.type('7');
await monthCell.blur();
await page.waitForTimeout(300);
const text2 = await monthCell.textContent();
console.log(text2 === '7' ? 'OK: 숫자 입력 정상 저장' : `FAIL: 숫자 저장 실패 (${text2})`);

// title click
await frame.locator('#grid-edit-table td[data-row="0"][data-col="0"]').click();
await page.waitForTimeout(300);
const received = await page.evaluate(() => window.__received);
console.log(received?.type === 'docview:title-click' && received.rect ? 'OK: 제목 클릭 postMessage+rect 정상' : 'FAIL: 제목 클릭 문제');

await browser.close();
