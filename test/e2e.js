const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args: ['--no-sandbox'] });
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  const errors = [];
  page.on('pageerror', e => errors.push('PAGEERROR: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push('CONSOLE: ' + m.text()); });

  // главная
  await page.goto('http://localhost:8901/index.html', { waitUntil: 'networkidle' });
  await page.waitForTimeout(2500);
  await page.screenshot({ path: 'test/shot_index.png', fullPage: true });

  // ЛК
  await page.goto('http://localhost:8901/app.html', { waitUntil: 'networkidle' });
  await page.screenshot({ path: 'test/shot_app_empty.png' });

  // загрузка реального файла пользователя (с опечатками в заголовках)
  const userFile = process.argv[2];
  await page.setInputFiles('#fileInput', userFile);
  await page.waitForSelector('#previewPanel:not([hidden])', { timeout: 10000 });
  await page.waitForTimeout(400);
  const hint = await page.textContent('#previewHint');
  console.log('HINT:', hint.trim());
  const rows = await page.$$eval('.preview-table tbody tr', trs => trs.length);
  console.log('PREVIEW ROWS:', rows);
  await page.screenshot({ path: 'test/shot_app_preview.png', fullPage: true });

  // генерация zip
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 30000 }),
    page.click('#generateBtn')
  ]);
  const zipPath = 'test/out.zip';
  await download.saveAs(zipPath);
  console.log('ZIP saved:', fs.statSync(zipPath).size, 'bytes');
  const status = await page.textContent('#statusLine');
  console.log('STATUS:', status.trim());
  await page.screenshot({ path: 'test/shot_app_done.png' });

  console.log('JS ERRORS:', errors.length ? errors.join('\n') : 'none');
  await browser.close();
})().catch(e => { console.error('FAIL:', e.message); process.exit(1); });
