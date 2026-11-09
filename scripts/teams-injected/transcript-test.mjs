import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const context = vm.createContext({});
vm.runInContext(fs.readFileSync(new URL('../../transcriptUtils.js', import.meta.url), 'utf8'), context);
const { enrichVtt, speakerMap, convertVttToTxt, speakerCoverage } = context.__tceTranscript;
const raw = 'WEBVTT\r\n\r\nmeeting/9-0\r\n00:00:21.507 --> 00:00:25.494\r\nHello,\r\nwelcome.\r\n\r\nmeeting/9-1\r\n00:00:25.494 --> 00:00:27.347\r\nNext line.\r\n\r\nother-meeting/9-0\r\n00:00:28.000 --> 00:00:29.000\r\nSomeone else.\r\n';
const named = 'WEBVTT\n\nmeeting/9\n00:00:21.000 --> 00:00:26.000\n<v Smith &amp; Jones>Welcome.</v>';
const enriched = enrichVtt(raw, speakerMap(named));
assert.equal(enriched.replace(/<v [^>]+>|<\/v>/g, ''), raw, 'Keep every original cue ID, timestamp, line break and word');
assert.match(enriched, /<v Smith &amp; Jones>Hello,\r\nwelcome.<\/v>/);
assert.match(enriched, /<v Smith &amp; Jones>Next line.<\/v>/, 'Split cues inherit their entry speaker');
assert.ok(enriched.endsWith('Someone else.\r\n'), 'Never reuse names from a different meeting');
assert.equal(enrichVtt(enriched, speakerMap(named)), enriched, 'Existing names remain untouched');
assert.equal(convertVttToTxt(enriched), '[00:00:21] Smith & Jones: Hello,\nwelcome.\n[00:00:25] Smith & Jones: Next line.\n[00:00:28] Someone else.', 'TXT omits GUIDs and handles multiline voice spans');
assert.equal(speakerCoverage(enriched).total, 3);
assert.equal(speakerCoverage(enriched).named, 2);
assert.equal(enrichVtt(raw, {}), raw, 'Missing names leave source unchanged');
assert.equal(enrichVtt(raw.replaceAll('meeting/9-0', 'constructor'), {}), raw.replaceAll('meeting/9-0', 'constructor'), 'Only explicit mappings count as names');
assert.equal(Object.keys(speakerMap(named.replace('Smith &amp; Jones', 'Unknown'))).length, 0);
const withNotes = 'WEBVTT\n\nNOTE metadata\nignore me\n\nSTYLE\n::cue { color: red; }\n\n00:00:00.000 --> 00:00:01.000\n<v.ann Speaker>Hi\nthere';
assert.equal(convertVttToTxt(withNotes), '[00:00:00] Speaker: Hi\nthere', 'Ignore metadata and accept an unclosed WebVTT voice span');
assert.match(enrichVtt(raw, { 'meeting/9': 'A <B> & C\nD' }), /<v A &lt;B&gt; &amp; C D>/, 'Escape voice annotations');

// Guests / people from other organisations: Teams can leave speakerDisplayName
// empty (or omit it) while still giving speakerId.
const { resolveSpeakers, rosterFrom, entryName } = context.__tceTranscript;
const guestEntries = [
  { id: 'm/1', speakerId: 'guest-1', text: 'Hi from outside', startOffset: '00:00:01' },
  { id: 'm/2', speakerId: 'staff-1', speakerDisplayName: 'Pat Staff', text: 'Hello', startOffset: '00:00:02' },
  { id: 'm/3', speakerId: 'guest-1', speakerDisplayName: 'Gail Guest (External)', text: 'Again', startOffset: '00:00:03' },
  { id: 'm/4', speakerId: 'guest-2', speakerDisplayName: null, text: 'No roster', startOffset: '00:00:04' }
];
const resolvedGuests = resolveSpeakers(guestEntries);
assert.deepEqual(resolvedGuests.map(e => e.speakerDisplayName), ['Gail Guest (External)', 'Pat Staff', 'Gail Guest (External)', null], 'Names copied only within the same speakerId');
const roster = rosterFrom({ entries: guestEntries, events: [{ eventType: 'JoinedConversation', userId: 'guest-2', userDisplayName: 'Gus Guest' }] });
assert.equal(resolveSpeakers(guestEntries, roster)[3].speakerDisplayName, 'Gus Guest', 'Transcript event roster names guests');
assert.equal(entryName({ speaker: { displayName: 'Nested Name' } }), 'Nested Name');
assert.equal(entryName({ speakerDisplayName: 'Unknown', speakerId: 'x' }), '');

// Recap page React state where the first entry is a guest without speakerDisplayName.
{
  const reactEntries = [
    { entryId: 'meeting/1', speakerId: 'guest-1', text: 'Guest speaks first', timestamp: 'PT1S' },
    { entryId: 'meeting/2', speakerId: 'staff-1', speakerDisplayName: 'Pat Staff', text: 'Staff reply', timestamp: 'PT5S' },
    { entryId: 'meeting/3', speakerId: 'guest-1', text: 'Guest again', timestamp: 'PT9S' }
  ];
  const listEl = { __reactFiber$x: { memoizedState: { memoizedState: reactEntries, next: null }, return: null } };
  const row = (label, text) => ({
    getAttribute: () => label,
    querySelector: () => ({ textContent: text })
  });
  let rows = [];
  const divs = {};
  const doc = {
    body: { appendChild(el) { divs[el.id] = el; } },
    documentElement: {},
    querySelector: (sel) => sel === '.ms-List' ? listEl : null,
    querySelectorAll: (sel) => sel === '[role="group"]' ? rows : [],
    getElementById: (id) => divs[id] || null,
    addEventListener() {},
    createElement: () => {
      const attrs = {};
      return { style: {}, setAttribute(k, v) { attrs[k] = v; }, getAttribute: k => attrs[k], set id(v) { this._id = v; }, get id() { return this._id; } };
    }
  };
  const page = vm.createContext({ console: { log() {}, warn() {} }, document: doc, setInterval() {}, MutationObserver: class { observe() {} } });
  page.window = page;
  page.fetch = async () => ({});
  vm.runInContext(fs.readFileSync(new URL('../../transcriptUtils.js', import.meta.url), 'utf8'), page);
  vm.runInContext(fs.readFileSync(new URL('../../transcriptFetchOverride.js', import.meta.url), 'utf8'), page);
  const utils = page.__teamsTranscriptUtils;
  let entries = utils.extractFromDOM();
  assert.equal(entries.length, 3, 'React state found even when the first speaker has no display name');
  assert.equal(entries[1].speakerDisplayName, 'Pat Staff', 'Organisation staff keep their names');
  // Teams renders the guest's name on screen; learn it for that speakerId.
  rows = [row('Gail Guest (External) 1 second', 'Guest speaks first')];
  entries = utils.extractFromDOM();
  assert.deepEqual(entries.map(e => e.speakerDisplayName), ['Gail Guest (External)', 'Pat Staff', 'Gail Guest (External)']);
  rows = [];
  entries = utils.extractFromDOM();
  assert.equal(entries[2].speakerDisplayName, 'Gail Guest (External)', 'Learned names persist after scrolling away');
  const names = speakerMap(utils.buildVttTranscript(entries));
  assert.equal(names['meeting/1'], 'Gail Guest (External)');
  assert.equal(names['meeting/2'], 'Pat Staff');
  const rawVtt = 'WEBVTT\n\nmeeting/1-0\n00:00:01.000 --> 00:00:04.000\nGuest speaks first\n\nmeeting/2-0\n00:00:05.000 --> 00:00:08.000\nStaff reply';
  assert.equal(convertVttToTxt(enrichVtt(rawVtt, names)), '[00:00:01] Gail Guest (External): Guest speaks first\n[00:00:05] Pat Staff: Staff reply');
  // Unresolved speakers never leak a raw speakerId into the name.
  assert.ok(!utils.buildVttTranscript([{ id: 'x/1', speakerId: 'guest-9', text: 'hi', startOffset: '00:00:01' }]).includes('guest-9'));
}

// Exercise the actual background message relay with a transcript frame reply.
const listeners = [];
const background = vm.createContext({
  chrome: {
    webRequest: { onBeforeSendHeaders: { addListener() {} } },
    runtime: { onMessage: { addListener(fn) { listeners.push(fn); } } },
    tabs: { sendMessage(tabId, request, callback) {
      assert.equal(tabId, 42);
      assert.equal(request.action, 'getLocalTranscriptSpeakers');
      callback({ speakers: speakerMap(named) });
    } },
    action: { onClicked: { addListener() {} } }
  }, console
});
vm.runInContext(fs.readFileSync(new URL('../../background.js', import.meta.url), 'utf8'), background);
let response;
const handled = listeners[0]({ action: 'getTranscriptSpeakers' }, { tab: { id: 42 } }, data => { response = data; });
assert.equal(handled, true);
assert.equal(response.speakers['meeting/9'], 'Smith & Jones', 'Names reach the requesting Teams frame');
let noTab;
listeners[0]({ action: 'getTranscriptSpeakers' }, {}, data => { noTab = data; });
assert.equal(Object.keys(noTab.speakers).length, 0, 'No cross-tab lookup without a sender tab');
console.log('PASS: speaker matching, source preservation, TXT conversion and cross-frame relay');
