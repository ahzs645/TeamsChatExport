import assert from 'node:assert/strict';
import { ApiMessageExtractor } from '../../src/modules/apiMessageExtractor.js';

const settings = new Map();
const observedResources = [];
Object.defineProperty(globalThis, 'localStorage', {
  value: { getItem: key => settings.get(key) ?? null },
  configurable: true,
});
Object.defineProperty(globalThis, 'performance', {
  value: { getEntriesByType: () => observedResources.map(name => ({ name })) },
  configurable: true,
});

async function fetchWithMock(location, matches = () => false, storageSettings = {}) {
  globalThis.window = { location: new URL(location) };
  const extractor = new ApiMessageExtractor();
  extractor.loadStorageSettings = async () => storageSettings;
  const requests = [];
  const pageLimits = [];
  const message = { id: 'fixture-message', message: 'Fixture message' };
  extractor.fetchPaged = async (url, { maxPages }) => {
    const parsed = new URL(url);
    requests.push(parsed);
    pageLimits.push(maxPages);
    return matches(parsed) ? [message] : [];
  };
  const messages = await extractor.fetchConversation('fixture-conversation');
  return { requests, messages, message, pageLimits };
}

observedResources.push(
  'https://teams.cloud.microsoft/api/chatsvc/ca/v1/users/ME/conversations/previous-fixture',
  'https://teams.cloud.microsoft/api/mt/amer/beta/me/settings/fixture',
);
const cloud = await fetchWithMock('https://teams.cloud.microsoft/', url =>
  url.hostname === 'teams.cloud.microsoft' && url.pathname.startsWith('/api/chatsvc/ca/'));
assert.deepEqual(cloud.messages, [cloud.message]);
assert.equal(cloud.requests.length, 1);
assert.equal(cloud.requests[0].pathname, '/api/chatsvc/ca/v1/users/ME/conversations/fixture-conversation/messages');
assert.deepEqual(Object.fromEntries(cloud.requests[0].searchParams), {
  view: 'msnp24Equivalent|supportsMessageProperties', pageSize: '200', startTime: '1',
});

const savedFullHistory = await fetchWithMock('https://teams.cloud.microsoft/',
  url => url.pathname.startsWith('/api/chatsvc/ca/'), { teamsChatApiPageSize: 500, teamsChatApiMaxPages: 200 });
assert.equal(savedFullHistory.requests[0].searchParams.get('pageSize'), '200');
assert.deepEqual(savedFullHistory.pageLimits, [200]);
assert.equal(new ApiMessageExtractor().readNumberSetting('missing', { max: 200, fallback: 500 }), 200);
console.log('PASS: existing full-history settings cap the page size and retain all 200 requests');

const failedCloud = await fetchWithMock('https://teams.cloud.microsoft/');
assert.deepEqual(failedCloud.requests.map(url => url.hostname), [
  'teams.cloud.microsoft', 'teams.cloud.microsoft', 'teams.cloud.microsoft', 'teams.cloud.microsoft',
  'teams.microsoft.com', 'teams.microsoft.com',
]);
assert.ok(failedCloud.requests[1].pathname.startsWith('/api/mt/amer/'));
assert.ok(failedCloud.requests[4].pathname.startsWith('/api/chatsvc/ca/'));

observedResources.length = 0;
const cloudWithoutRegion = await fetchWithMock('https://teams.cloud.microsoft/', url =>
  url.pathname.startsWith('/api/chatsvc/ca/'));
assert.deepEqual(cloudWithoutRegion.requests.map(url => url.pathname.split('/')[3]), ['amer', 'ca']);

const legacy = await fetchWithMock('https://teams.microsoft.com/');
assert.ok(legacy.requests.every(url => url.pathname.startsWith('/api/chatsvc/')));
const legacyV2 = await fetchWithMock('https://teams.microsoft.com/v2/');
assert.ok(legacyV2.requests[0].pathname.startsWith('/api/mt/'));

settings.set('teamsChatHostOverride', 'teams.example.test');
settings.set('teamsChatRegionOverride', 'emea');
const override = await fetchWithMock('https://teams.cloud.microsoft/');
assert.ok(override.requests.every(url => url.hostname === 'teams.example.test' && url.pathname.includes('/emea/')));
console.log('PASS: cloud API routing, legacy compatibility and overrides');

const extractor = new ApiMessageExtractor();
const testCredential = 'fixture-credential';
globalThis.document = { getElementById: () => ({ getAttribute: () => testCredential }) };
extractor.isValidTeamsToken = () => true;
const initialFetch = globalThis.fetch;
const initialInfo = console.info;
const diagnostics = [];
try {
  console.info = (...args) => diagnostics.push(args);
  const endpoint = 'https://teams.cloud.microsoft/api/mt/amer/beta/chats/fixture-chat-id/messages';
  globalThis.fetch = async () => new Response(null, { status: 404 });
  assert.deepEqual(await extractor.fetchPaged(endpoint), []);
  for (const detail of ['Invalid conversation ID fixture-chat-id', 'Unsupported client version for fixture-credential']) {
    globalThis.fetch = async () => Response.json({ error: { code: 'BadRequest', message: detail } }, { status: 400 });
    assert.deepEqual(await extractor.fetchPaged(endpoint), []);
  }
  const responses = diagnostics.map(([, details]) => JSON.parse(details));
  assert.deepEqual(responses.map(details => details.status), [404, 400, 400]);
  assert.ok(responses.every(details => details.host === 'teams.cloud.microsoft' && details.route === 'v2'));
  assert.deepEqual(responses.map(details => details.rejection), [undefined, 'conversation-id', 'client-version']);
  assert.doesNotMatch(JSON.stringify(diagnostics), /fixture-chat-id|fixture-credential/);
} finally {
  globalThis.fetch = initialFetch;
  console.info = initialInfo;
}
console.log('PASS: failed API responses record status without chat IDs or credentials');

const historyUrl = cloud.requests[0].href;
const olderPage = new URL(historyUrl);
olderPage.searchParams.set('syncState', 'older-fixture');
const newerPage = new URL(historyUrl);
newerPage.searchParams.set('syncState', 'newer-fixture');
const message = { id: 'recent-message', message: 'Recent fixture' };
const olderMessage = { id: 'older-message', message: 'Older fixture' };
const visited = [];
extractor.extractFromPayload = payload => payload.messages;
try {
  globalThis.fetch = async url => {
    visited.push(url);
    if (url === historyUrl) return Response.json({
      messages: [message],
      _metadata: { backwardLink: olderPage.href, syncState: newerPage.href },
    });
    return Response.json({
      messages: [olderMessage],
      _metadata: { backwardLink: 'https://other.example.test/fixture' },
    });
  };
  const messages = await extractor.fetchPaged(historyUrl);
  assert.deepEqual(visited, [historyUrl, olderPage.href]);
  assert.deepEqual(new Set(messages.map(item => item.id)), new Set([message.id, olderMessage.id]));
} finally {
  globalThis.fetch = initialFetch;
}
console.log('PASS: history pagination follows the same-origin backward link');
