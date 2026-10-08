// Harness for the UI that content.js injects into Teams / SharePoint Stream.
// Loads the REAL content.js, ui/tceUI.js, transcriptStyles.css,
// videoDownload/coordinator.js and videoDownload/manifestDownload.js into mock
// Teams pages (page-world capture/batch scripts are replaced by mocks in
// boot.js), drives them with Playwright and saves screenshots.
//
// Usage (from the repo root):
//   NODE_PATH=/opt/node22/lib/node_modules node scripts/teams-injected/run.mjs <outDir> [zipOutDir]
import { createRequire } from 'module';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const require = createRequire(import.meta.url);
const { chromium } = require(process.env.NODE_PATH ? path.join(process.env.NODE_PATH.split(':')[0], 'playwright') : 'playwright');

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, '../..');
const OUT = path.resolve(process.argv[2] || path.join(HERE, 'out'));
const ZIP_OUT = path.resolve(process.argv[3] || OUT);
fs.mkdirSync(OUT, { recursive: true });
fs.mkdirSync(ZIP_OUT, { recursive: true });

// Page-world scripts replaced by mocks in boot.js.
const MOCKED = new Set([
  'chatFetchOverride.js', 'contextBridge.js', 'transcriptAPIFetcher.js', 'transcriptFetchOverride.js',
  'videoDownloadOverride.js', 'batchTranscriptDownload.js', 'videoDownload/directDownload.js',
  'videoDownload/mseCaptureDownload.js', 'videoDownload/captureStreamDownload.js', 'videoDownload/fmp4ToMp4.js'
]);
const TYPES = { '.js': 'application/javascript', '.mjs': 'application/javascript', '.css': 'text/css', '.html': 'text/html; charset=utf-8', '.json': 'application/json' };

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok: !!ok, detail });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  -- ' + detail : ''}`);
};

async function setupRoutes(context) {
  await context.route('https://ext.test/**', (route) => {
    const rel = decodeURIComponent(new URL(route.request().url()).pathname.slice(1));
    const headers = { 'access-control-allow-origin': '*', 'content-type': TYPES[path.extname(rel)] || 'text/plain' };
    if (MOCKED.has(rel)) return route.fulfill({ status: 200, headers, body: '/* mocked in harness */' });
    const file = path.join(REPO, rel);
    if (!file.startsWith(REPO) || !fs.existsSync(file)) return route.fulfill({ status: 404, headers, body: '' });
    return route.fulfill({ status: 200, headers, body: fs.readFileSync(file) });
  });
  const serveHarness = (route) => {
    const url = new URL(route.request().url());
    const name = url.searchParams.get('harness') || path.basename(url.pathname);
    const file = path.join(HERE, name);
    if (!fs.existsSync(file)) return route.fulfill({ status: 404, body: 'not found' });
    return route.fulfill({ status: 200, headers: { 'content-type': TYPES[path.extname(file)] || 'text/plain' }, body: fs.readFileSync(file) });
  };
  for (const host of ['https://teams.cloud.microsoft/**', 'https://teams.microsoft.com/**', 'https://contoso.sharepoint.com/**']) {
    await context.route(host, serveHarness);
  }
}

const shot = async (page, name, opts = {}) => {
  const file = path.join(OUT, `teams-injected-${name}.png`);
  await page.screenshot({ path: file, ...opts });
  console.log('  shot', file);
};

const contrast = (a, b) => {
  const lum = (c) => {
    const [r, g, b2] = c.match(/[\d.]+/g).slice(0, 3).map(Number).map((v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; });
    return 0.2126 * r + 0.7152 * g + 0.0722 * b2;
  };
  const [l1, l2] = [lum(a), lum(b)].sort((x, y) => y - x);
  return (l1 + 0.05) / (l2 + 0.05);
};

async function openHarness(browser, { theme = 'light', width = 1280, height = 820, themeclass = false } = {}) {
  const context = await browser.newContext({ viewport: { width, height }, acceptDownloads: true });
  await setupRoutes(context);
  const page = await context.newPage();
  page.on('pageerror', (e) => console.log('  [pageerror]', e.message));
  page.on('console', (m) => { if (m.type() === 'error') console.log('  [console.error]', m.text()); });
  await page.goto(`https://teams.cloud.microsoft/harness.html?theme=${theme}${themeclass ? '&themeclass=1' : ''}`);
  await page.waitForSelector('.tce-toolbar-btn', { timeout: 8000 });
  await page.waitForSelector('.tce-recap-bar', { timeout: 8000 });
  await page.waitForSelector('.tce-cmd-btn', { timeout: 8000 });
  return { context, page };
}

async function toolbarChecks(browser, theme) {
  const { context, page } = await openHarness(browser, { theme });
  const info = await page.evaluate(() => {
    const btn = document.querySelector('.tce-toolbar-btn');
    const card = btn.closest('.card');
    return {
      color: getComputedStyle(btn).color,
      bg: getComputedStyle(card).backgroundColor,
      labels: [...document.querySelectorAll('.tce-toolbar-btn')].map((b) => b.getAttribute('aria-label')),
      recapEmoji: /[\u{1F300}-\u{1FAFF}⬇]/u.test(document.querySelector('.tce-recap-bar').textContent),
      recapInline: [...document.querySelectorAll('.tce-recap-bar button')].some((b) => b.getAttribute('style')),
      recapSvg: document.querySelectorAll('.tce-recap-bar button svg').length
    };
  });
  const ratio = contrast(info.color, info.bg);
  check(`[${theme}] toolbar icon contrast >= 4.5`, ratio >= 4.5, `${info.color} on ${info.bg} = ${ratio.toFixed(2)}:1`);
  check(`[${theme}] toolbar buttons have aria-label`, info.labels.every(Boolean), info.labels.join(', '));
  check(`[${theme}] recap bar: no emoji, no inline styles, SVG icons`, !info.recapEmoji && !info.recapInline && info.recapSvg === 4);
  await shot(page, `toolbar-${theme}`, { clip: { x: 0, y: 0, width: 1280, height: 640 } });
  await context.close();
}

async function menuChecks(browser, theme) {
  const { context, page } = await openHarness(browser, { theme });
  const trigger = page.locator('.tce-toolbar-btn[aria-label="Download transcript"]');
  await trigger.focus();
  await page.keyboard.press('ArrowDown');
  await page.waitForSelector('.tce-menu');
  await page.keyboard.press('ArrowDown');
  const st = await page.evaluate(() => {
    const m = document.querySelector('.tce-menu');
    const t = document.querySelector('.tce-toolbar-btn[aria-label="Download transcript"]');
    return {
      parentIsBody: m.parentElement === document.body,
      insideTrigger: t.contains(m),
      role: m.getAttribute('role'),
      itemRoles: [...m.children].map((c) => c.getAttribute('role')),
      expanded: t.getAttribute('aria-expanded'),
      focused: document.activeElement.textContent,
      theme: m.getAttribute('data-tce-theme')
    };
  });
  check(`[${theme}] menu appended to body, not inside trigger`, st.parentIsBody && !st.insideTrigger);
  check(`[${theme}] menu roles`, st.role === 'menu' && st.itemRoles.every((r) => r === 'menuitem') && st.expanded === 'true', JSON.stringify(st));
  check(`[${theme}] arrow-key navigation moves focus`, st.focused === 'Download TXT', st.focused);
  check(`[${theme}] menu theme matches page`, st.theme === theme, st.theme);
  await shot(page, `menu-${theme}`, { clip: { x: 0, y: 360, width: 640, height: 200 } });
  await page.keyboard.press('Escape');
  const afterEsc = await page.evaluate(() => ({ open: !!document.querySelector('.tce-menu'), focus: document.activeElement.getAttribute('aria-label') }));
  check(`[${theme}] Escape closes menu and returns focus`, !afterEsc.open && afterEsc.focus === 'Download transcript', JSON.stringify(afterEsc));
  // Mouse: click toggles open/closed (no re-open via bubbling)
  await trigger.click();
  const open1 = await page.locator('.tce-menu').count();
  await trigger.click();
  const open2 = await page.locator('.tce-menu').count();
  check(`[${theme}] click toggles menu (open then closed)`, open1 === 1 && open2 === 0, `${open1} -> ${open2}`);
  // Select an item: download with non-ASCII filename, success toast
  await trigger.click();
  const [download] = await Promise.all([
    page.waitForEvent('download'),
    page.locator('.tce-menu__item', { hasText: 'Download VTT' }).click()
  ]);
  const fname = download.suggestedFilename();
  check(`[${theme}] transcript filename keeps non-ASCII`, fname.includes('週次デザインレビュー') && fname.endsWith('.vtt'), fname);
  const menuGone = await page.locator('.tce-menu').count();
  check(`[${theme}] menu closes after selection`, menuGone === 0);
  await page.waitForSelector('.tce-toast--success');
  await context.close();
}

async function panelChecks(browser, theme) {
  const { context, page } = await openHarness(browser, { theme });
  await page.locator('.tce-toolbar-btn[aria-label="Batch download all transcripts"]').click();
  await page.waitForSelector('#tce-batch-panel');
  const focusIn = await page.evaluate(() => document.getElementById('tce-batch-panel').contains(document.activeElement) && document.activeElement.textContent);
  check(`[${theme}] focus moves into batch panel on open`, !!focusIn, String(focusIn));
  await page.locator('#tce-batch-panel .tce-btn', { hasText: 'Start' }).click();
  await page.waitForFunction(() => !document.querySelector('#tce-batch-panel .tce-btn--primary[disabled]') || document.querySelector('#tce-batch-panel').textContent.includes('Ready to download'), null, { timeout: 5000 });
  await page.waitForTimeout(200);
  const bar = await page.evaluate(() => {
    const pb = document.querySelector('#tce-batch-panel [role="progressbar"]');
    return { now: pb.getAttribute('aria-valuenow'), max: pb.getAttribute('aria-valuemax'), text: pb.getAttribute('aria-valuetext') };
  });
  check(`[${theme}] batch progressbar has aria-valuenow/max`, bar.now === '8' && bar.max === '8', JSON.stringify(bar));

  await page.locator('.tce-toolbar-btn[aria-label="Download video"]').click();
  await page.waitForSelector('#tce-video-panel');
  await page.waitForTimeout(300);
  const boxes = await page.evaluate(() => ['#tce-batch-panel', '#tce-video-panel'].map((s) => document.querySelector(s).getBoundingClientRect().toJSON()));
  const [a, b] = boxes;
  const overlap = !(a.bottom <= b.top || b.bottom <= a.top || a.right <= b.left || b.right <= a.left);
  check(`[${theme}] batch + video panels stack without overlapping`, !overlap, `batch ${Math.round(a.top)}-${Math.round(a.bottom)}, video ${Math.round(b.top)}-${Math.round(b.bottom)}`);
  const dialogs = await page.evaluate(() => [...document.querySelectorAll('.tce-panel')].map((p) => ({
    role: p.getAttribute('role'), labelled: !!document.getElementById(p.getAttribute('aria-labelledby')), close: p.querySelector('.tce-panel__close').getAttribute('aria-label')
  })));
  check(`[${theme}] panels are labelled dialogs with aria-label=Close`, dialogs.every((d) => d.role === 'dialog' && d.labelled && d.close === 'Close'), JSON.stringify(dialogs));
  const noDebug = await page.evaluate(() => !document.querySelector('#tce-video-panel').textContent.match(/🔍|analy/i));
  check(`[${theme}] no debug analyze button in video panel`, noDebug);
  const emoji = await page.evaluate(() => /[\u{1F300}-\u{1FAFF}☀-➿]/u.test(document.getElementById('tce-stack').textContent));
  check(`[${theme}] no emoji in panels`, !emoji);
  await page.setViewportSize({ width: 1280, height: 1100 });
  await shot(page, `panels-${theme}`);

  if (theme === 'light') {
    // Zip download of the batch results
    const [dl] = await Promise.all([
      page.waitForEvent('download'),
      page.locator('#tce-batch-panel .tce-btn', { hasText: 'Download .zip' }).click()
    ]);
    const zipPath = path.join(ZIP_OUT, 'batch-from-harness.zip');
    await dl.saveAs(zipPath);
    check('batch Download .zip produces one .zip download', dl.suggestedFilename().endsWith('.zip'), `${dl.suggestedFilename()} -> ${zipPath}`);

    // Keyboard: Tab within the video panel shows a focus ring; Escape closes and restores focus
    await page.locator('.tce-toolbar-btn[aria-label="Download video"]').focus();
    await page.evaluate(() => document.querySelector('#tce-video-panel .tce-btn').focus());
    await page.keyboard.press('Tab');
    const ring = await page.evaluate(() => getComputedStyle(document.activeElement).boxShadow);
    check('focused panel control shows a focus ring', ring && ring !== 'none', ring);
    await shot(page, 'keyboard-focus', { clip: { x: 1280 - 420, y: 40, width: 420, height: 1000 } });
    await page.keyboard.press('Escape');
    const afterEsc = await page.evaluate(() => ({ panel: !!document.getElementById('tce-video-panel'), focus: document.activeElement.getAttribute('aria-label') }));
    check('Escape closes video panel; focus returns to its trigger', !afterEsc.panel && afterEsc.focus === 'Download video', JSON.stringify(afterEsc));
  }
  await context.close();
}

async function progressChecks(browser) {
  const { context, page } = await openHarness(browser, { theme: 'light', width: 600, height: 760 });
  await page.evaluate(() => window.__dispatch({ action: 'downloadVideo', data: { method: 'captureStreamDownload' } }));
  await page.waitForSelector('#tce-download-progress', { timeout: 5000 });
  const m = await page.evaluate(() => ({
    w: document.getElementById('tce-download-progress').getBoundingClientRect().width,
    left: document.getElementById('tce-download-progress').getBoundingClientRect().left,
    right: document.getElementById('tce-download-progress').getBoundingClientRect().right,
    scroll: document.documentElement.scrollWidth,
    close: !!document.querySelector('#tce-download-progress .tce-panel__close'),
    font: getComputedStyle(document.getElementById('tce-download-progress')).fontFamily
  }));
  check('progress panel fits a 600px window', m.w <= 600 - 24 && m.left >= 0 && m.right <= 600 && m.scroll <= 600, `width ${m.w}, x ${m.left}-${m.right}, scrollWidth ${m.scroll}`);
  check('progress panel has a close button and no monospace', m.close && !/mono/i.test(m.font), m.font);
  await shot(page, 'progress-600');

  // Failure: readable error + close button, plus an error toast elsewhere
  await page.evaluate(() => window.__dl.resolve({ success: false, error: 'Segment 37 returned HTTP 403. The playback token expired; play the video for a few seconds and try again.' }));
  await page.waitForFunction(() => document.querySelector('#tce-download-progress .tce-status--error'));
  const fail = await page.evaluate(() => ({
    text: document.querySelector('#tce-download-progress .tce-status').textContent,
    title: document.querySelector('#tce-download-progress .tce-panel__title').textContent,
    close: !!document.querySelector('#tce-download-progress .tce-panel__close')
  }));
  check('coordinator failure shows message + close', fail.close && !/^ERROR/.test(fail.text) && fail.title.includes('failed'), JSON.stringify(fail));
  // Error toast: transcript unavailable
  await page.evaluate(() => document.getElementById('teams-chat-exporter-transcript-data').remove());
  await page.locator('.tce-toolbar-btn[aria-label="Copy transcript"]').click();
  await page.waitForSelector('.tce-toast--error');
  const toast = await page.evaluate(() => {
    const t = document.querySelector('.tce-toast--error');
    return { role: t.getAttribute('role'), close: !!t.querySelector('.tce-toast__close'), text: t.textContent };
  });
  check('error toast uses role=alert and has close', toast.role === 'alert' && toast.close, JSON.stringify(toast));
  await page.waitForTimeout(6000);
  check('error toast persists (no auto-dismiss)', await page.locator('.tce-toast--error').count() === 1);
  await shot(page, 'error-states-600');
  await context.close();
}

async function extractionChecks(browser) {
  const { context, page } = await openHarness(browser, { theme: 'dark' });
  await page.evaluate(() => window.__dispatch({ action: 'extractActiveChat' }));
  await page.waitForSelector('.tce-toast--progress');
  await page.waitForTimeout(900);
  await shot(page, 'extract-progress-dark', { clip: { x: 1280 - 420, y: 40, width: 420, height: 200 } });
  await page.waitForSelector('.tce-toast--success', { timeout: 15000 });
  await shot(page, 'extract-done-dark', { clip: { x: 1280 - 420, y: 40, width: 420, height: 200 } });
  const sent = await page.evaluate(() => window.__sent.filter((m) => m.action === 'extractionProgress').map((m) => `${m.status}:${m.count}`));
  check('extraction broadcasts progress to the extension', sent[0].startsWith('started') && sent.at(-1).startsWith('done'), sent.join(' '));
  const txt = await page.evaluate(() => document.querySelector('.tce-toast--success').textContent);
  check('extraction ends with "Opened in viewer" toast', /viewer/.test(txt), txt);
  await context.close();
}

async function savePanelCheck(browser) {
  const { context, page } = await openHarness(browser, { theme: 'light', width: 900, height: 600 });
  await page.waitForFunction(() => window.__videoDownloadModules?.manifestDownload?.showSavePanel);
  await page.evaluate(() => {
    window.__videoDownloadModules.manifestDownload.showSavePanel({
      fileName: '週次デザインレビュー 2026-10-01', combinedBlob: null,
      vBlob: new Blob([new Uint8Array(212 * 1024 * 1024)]), aBlob: new Blob([new Uint8Array(31 * 1024 * 1024)]), elapsed: 48, fixed: 965
    });
  });
  await page.locator('#tce-video-save-panel summary').click();
  await shot(page, 'save-panel-light');
  await context.close();
}

async function themeDetectionChecks(browser) {
  const { context, page } = await openHarness(browser, { theme: 'dark', themeclass: true });
  const t1 = await page.evaluate(() => window.__tceUI.detectTheme());
  check('detectTheme: Teams "theme-dark" class', t1 === 'dark', t1);
  const t2 = await page.evaluate(() => { document.body.className = 'theme-default'; return window.__tceUI.detectTheme(); });
  check('detectTheme: Teams "theme-default" class', t2 === 'light', t2);
  const t3 = await page.evaluate(() => { document.body.className = 'mock-dark'; return window.__tceUI.detectTheme(); });
  check('detectTheme: luminance fallback (dark bg, no class)', t3 === 'dark', t3);
  const san = await page.evaluate(() => [
    window.__tceUI.sanitizeFilename('会議: 設計/レビュー? "Q3" <draft>|*'),
    window.__tceUI.sanitizeFilename('   '),
    window.__tceUI.sanitizeFilename('a'.repeat(300)).length,
    window.__tceUI.sanitizeFilename('CON')
  ]);
  check('sanitizeFilename keeps non-ASCII, strips illegal chars', san[0] === '会議 設計 レビュー Q3 draft' && san[1] === 'transcript' && san[2] === 120 && san[3] === 'CON_', JSON.stringify(san));
  await context.close();
}

async function frameChecks(browser) {
  const context = await browser.newContext({ viewport: { width: 1000, height: 700 } });
  await setupRoutes(context);
  const page = await context.newPage();
  page.on('pageerror', (e) => console.log('  [pageerror]', e.message));
  await page.goto('https://teams.microsoft.com/frame-top.html');
  await page.waitForTimeout(1500);
  const frames = page.frames();
  const ready = await Promise.all(frames.map((f) => f.evaluate(() => !!window.__msgListeners?.length).catch(() => false)));
  check('frame harness: content.js running in 3 frames', ready.filter(Boolean).length === 3, `${ready}`);
  const name = (f) => (f === page.mainFrame() ? 'top' : new URL(f.url()).searchParams.get('harness'));

  const dispatchAll = (req) => Promise.all(frames.map((f) => f.evaluate((r) => window.__dispatch(r, 2500), req).then((res) => ({ frame: name(f), ...res }))));

  // UI actions: only the player frame acts.
  await dispatchAll({ action: 'startBatchTranscript' });
  await dispatchAll({ action: 'downloadVideo', data: { method: 'captureStreamDownload' } });
  await page.waitForTimeout(1200);
  const counts = await Promise.all(frames.map(async (f) => ({
    frame: name(f),
    batch: await f.evaluate(() => document.querySelectorAll('#tce-batch-panel').length),
    progress: await f.evaluate(() => document.querySelectorAll('#tce-download-progress').length)
  })));
  const player = counts.find((c) => c.frame === 'frame-player.html');
  const others = counts.filter((c) => c.frame !== 'frame-player.html');
  check('startBatchTranscript/downloadVideo handled by the player frame only', player.batch === 1 && player.progress === 1 && others.every((c) => !c.batch && !c.progress), JSON.stringify(counts));
  await shot(page, 'frames-player-only');

  // Readiness: the frame that has transcript data answers first.
  const t0 = Date.now();
  const status = await dispatchAll({ action: 'getTranscriptStatus' });
  const answered = status.filter((s) => s.response).sort((a, b) => a.at - b.at);
  check('getTranscriptStatus: player frame (has data) answers first; unrelated frame stays silent',
    answered[0]?.frame === 'frame-player.html' && answered[0].response.available === true && !status.find((s) => s.frame === 'frame-other.html').response,
    answered.map((s) => `${s.frame}@${s.at - t0}ms available=${s.response.available} source=${s.response.source}`).join('; '));

  // Fallback: no frame has a player -> the top frame handles it.
  await page.goto('https://teams.microsoft.com/frame-other.html');
  await page.waitForTimeout(800);
  await page.evaluate(() => window.__dispatch({ action: 'startBatchTranscript' }, 1500));
  await page.waitForTimeout(300);
  check('no relevant frame: top frame falls back and opens the panel', await page.locator('#tce-batch-panel').count() === 1);
  await context.close();
}

// UTF-8 locale so Chromium keeps non-ASCII download filenames (with LANG unset
// it falls back to "download").
const browser = await chromium.launch({ env: { ...process.env, LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8' } });
try {
  for (const theme of ['light', 'dark']) await toolbarChecks(browser, theme);
  for (const theme of ['light', 'dark']) await menuChecks(browser, theme);
  for (const theme of ['light', 'dark']) await panelChecks(browser, theme);
  await progressChecks(browser);
  await extractionChecks(browser);
  await savePanelCheck(browser);
  await themeDetectionChecks(browser);
  await frameChecks(browser);
} finally {
  await browser.close();
}
const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} checks passed`);
process.exitCode = failed.length ? 1 : 0;
