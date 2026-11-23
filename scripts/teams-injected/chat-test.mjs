// Run with the same Playwright installation as run.mjs.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const { chromium } = require(process.env.NODE_PATH ? path.join(process.env.NODE_PATH.split(':')[0], 'playwright') : 'playwright');
const repo = fileURLToPath(new URL('../../', import.meta.url));
const temp = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-chat-test-'));
const extension = path.join(temp, 'extension');
const domain = 'https://teams.cloud.microsoft';
const title = 'Fixture chat';
const messageId = (number) => String(Date.UTC(2026, 0, number));
const chatId = '19:fixture-history@thread.v2';

const fixture = `<!doctype html><title>${title} | Microsoft Teams</title>
  <h2>Sidebar</h2>
  <header id="chat-header-${chatId}"><h2>${title}</h2></header>
  <div data-tid="message-pane-list-viewport" style="height: 200px; overflow-y: auto; overflow-anchor: none">
    <div id="chat-pane-list" style="position: relative"></div>
  </div><script src="/fixture.js"></script>`;

const fixtureScript = `(${function renderVirtualChat() {
  const runway = document.getElementById('chat-pane-list');
  const viewport = runway.parentElement;
  const rowHeight = 50;
  let oldest = 9;
  let olderRequested = false;

  function render() {
    runway.style.height = `${(25 - oldest) * rowHeight}px`;
    runway.replaceChildren();
    const first = oldest + Math.floor(viewport.scrollTop / rowHeight);
    const end = Math.min(25, first + 6);
    for (let number = first; number < end; number++) {
      const row = document.createElement('div');
      row.dataset.tid = 'chat-pane-message';
      row.dataset.mid = String(Date.UTC(2026, 0, number));
      row.style.cssText = `position: absolute; top: ${(number - oldest) * rowHeight}px; height: ${rowHeight}px`;
      const text = document.createElement('div');
      text.id = `content-${row.dataset.mid}`;
      text.textContent = `Fixture message ${number}`;
      row.append(text);
      runway.append(row);
    }
  }

  render();
  viewport.scrollTop = viewport.scrollHeight - viewport.clientHeight - 100;
  render();
  if (location.pathname === '/api-export') {
    document.addEventListener('teams-extractor-update-bridge', () => queueMicrotask(() => {
      let bridge = document.getElementById('teams-extractor-context-bridge');
      if (!bridge) {
        bridge = document.createElement('div');
        bridge.hidden = true;
        bridge.id = 'teams-extractor-context-bridge';
        document.body.append(bridge);
      }
      bridge.setAttribute('data-token', 'skype_token fixture-only');
      bridge.setAttribute('data-conversation-id', '19:fixture-history@thread.v2');
    }));
  }
  viewport.addEventListener('scroll', () => {
    render();
    if (viewport.scrollTop !== 0 || olderRequested) return;
    olderRequested = true;
    fetch('/api/chatsvc/amer/v1/users/ME/conversations/19%3Afixture-history%40thread.v2/messages?pageSize=8');
    fetch('/api/chatsvc/amer/v1/$batch', { method: 'POST' });
    setTimeout(() => {
      oldest = 1;
      render();
      viewport.scrollTop = 8 * rowHeight;
      render();
    }, 350);
  });
}.toString()})()`;

let context;
try {
  await fs.cp(repo, extension, {
    recursive: true,
    filter: (source) => !['.git', '.github', '.vscode', 'node_modules', 'scripts'].includes(path.basename(source)) &&
      !source.endsWith('.zip'),
  });
  const manifestPath = path.join(extension, 'manifest.json');
  const manifest = JSON.parse(await fs.readFile(manifestPath, 'utf8'));
  // Simulate a tab where automatic startup was missed, without changing the
  // extension's host permissions or the popup's recovery code.
  manifest.content_scripts[0].exclude_matches = [`${domain}/reconnect`];
  await fs.writeFile(manifestPath, JSON.stringify(manifest));

  context = await chromium.launchPersistentContext(path.join(temp, 'profile'), {
    channel: 'chromium',
    args: [`--disable-extensions-except=${extension}`, `--load-extension=${extension}`],
  });
  await context.route(`${domain}/**`, async (route) => {
    const url = new URL(route.request().url());
    const { pathname } = url;
    if (pathname === '/fixture.js') return route.fulfill({ contentType: 'text/javascript', body: fixtureScript });
    if (pathname.endsWith('/messages') && url.searchParams.get('pageSize') === '200') {
      if (!pathname.startsWith('/api/chatsvc/ca/')) return route.fulfill({ status: 404 });
      const older = url.searchParams.has('syncState');
      const day = older ? 1 : 2;
      url.searchParams.set('syncState', 'fixture-older-page');
      return route.fulfill({ json: {
        messages: [{
          id: messageId(day), content: `API fixture message ${day}`, imdisplayname: 'Fixture sender',
          originalarrivaltime: new Date(Number(messageId(day))).toISOString(), messagetype: 'Text',
        }],
        _metadata: older ? {} : { backwardLink: url.href },
      } });
    }
    if (pathname.endsWith('/messages')) return route.fulfill({ json: {} });
    if (pathname.endsWith('/$batch')) return route.fulfill({ json: {} });
    if (!['/', '/reconnect', '/api-export'].includes(pathname)) return route.fulfill({ status: 404 });
    return route.fulfill({
      contentType: 'text/html',
      headers: { 'content-security-policy': "script-src 'self'; require-trusted-types-for 'script'; trusted-types 'none'" },
      body: fixture,
    });
  });
  const worker = context.serviceWorkers()[0] || await context.waitForEvent('serviceworker');
  const popupUrl = new URL('popup.html', worker.url()).href;

  for (const pathname of ['/', '/reconnect', '/api-export']) {
    const usingApi = pathname === '/api-export';
    const teamsPage = await context.newPage();
    const pageUrl = domain + pathname;
    await teamsPage.goto(pageUrl);
    const initialChat = await worker.evaluate(async ({ url, waitForStartup }) => {
      const [tab] = await chrome.tabs.query({ url });
      const attempts = waitForStartup ? 20 : 1;
      for (let attempt = 0; attempt < attempts; attempt++) {
        try {
          const state = await chrome.tabs.sendMessage(tab.id, { action: 'getState' });
          if (state?.currentChat) return state.currentChat;
        } catch {
          // The content script has not registered its listener yet.
        }
        if (attempt + 1 < attempts) await new Promise(resolve => setTimeout(resolve, 100));
      }
      return null;
    }, { url: pageUrl, waitForStartup: pathname !== '/reconnect' });
    assert.equal(initialChat, pathname === '/reconnect' ? null : title);

    if (usingApi) {
      await worker.evaluate(() => chrome.storage.local.set({ teamsChatApiPageSize: 500, teamsChatApiMaxPages: 200 }));
    }
    const popup = await context.newPage();
    await teamsPage.bringToFront();
    await popup.goto(popupUrl);
    await popup.getByText(`Teams chat: ${title}`, { exact: true }).waitFor();
    assert.equal(await popup.locator('#extractActiveChatBtn').getAttribute('aria-disabled'), null);
    if (usingApi) {
      await popup.waitForFunction(() => document.getElementById('maxMode').checked);
      assert.equal(await popup.locator('#pageSize').inputValue(), '200');
      assert.equal(await popup.locator('#maxPages').inputValue(), '200');
    }

    const nativeRequest = usingApi ? null : teamsPage.waitForEvent('console', {
      predicate: message => message.text().startsWith('[Teams native chat history] {') && message.text().includes('"matchedChat":true'),
      timeout: 10000,
    });
    const genericRequest = usingApi ? null : teamsPage.waitForEvent('console', {
      predicate: message => message.text().startsWith('[Teams native chat history] {') && message.text().includes('"matchedChat":false'),
      timeout: 10000,
    });
    const viewerOpened = context.waitForEvent('page');
    await popup.getByRole('button', { name: 'Extract this chat' }).click();
    const viewer = await viewerOpened;
    await viewer.waitForURL('**/results.html?*');
    const { teamsChatData } = await worker.evaluate(() => chrome.storage.local.get('teamsChatData'));
    assert.deepEqual(teamsChatData[title].map(message => message.id), Array.from({ length: usingApi ? 2 : 24 }, (_, i) => messageId(i + 1)));
    assert.equal(await teamsPage.locator('[data-tid="message-pane-list-viewport"]').evaluate(
      (list) => list.scrollHeight - list.clientHeight - list.scrollTop), 100);
    if (usingApi) {
      assert.ok(teamsChatData[title].every(message => message.message.startsWith('API fixture message')));
      console.log('PASS: direct export uses the observed cloud route to fetch both fixture pages');
    } else {
      const report = (await nativeRequest).text();
      assert.match(report, /"route":"\/api\/chatsvc\/amer\/v1\/users\/ME\/conversations\/\{chatId\}\/messages"/);
      assert.match(report, /"status":200/);
      assert.match(report, /"matchedChat":true/);
      assert.doesNotMatch(report, /fixture-history|pageSize=8/);
      assert.match((await genericRequest).text(), /"route":"\/api\/chatsvc\/amer\/v1\/\$batch"/);
      console.log(`PASS: ${pathname === '/' ? 'automatic startup' : 'popup recovery'} and complete virtualized fixture export`);
    }
    await viewer.close();
    await popup.close();
    await teamsPage.close();
  }
} finally {
  await context?.close();
  await fs.rm(temp, { recursive: true, force: true });
}
