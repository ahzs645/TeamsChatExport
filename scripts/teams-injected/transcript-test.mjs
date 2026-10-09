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
