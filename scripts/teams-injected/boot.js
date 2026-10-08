// Harness boot: stubs the `chrome` global that content.js expects and mocks the
// page-world scripts (batch transcript / video capture) so the REAL content.js,
// ui/tceUI.js, transcriptStyles.css, videoDownload/coordinator.js and
// videoDownload/manifestDownload.js can be exercised without Teams.
// Extension files are served by run.mjs at https://ext.test/<path>.
(() => {
  const listeners = [];
  window.__sent = [];
  window.__msgListeners = listeners;
  window.chrome = {
    runtime: {
      lastError: undefined,
      getURL: (p) => 'https://ext.test/' + p,
      sendMessage: (msg, cb) => {
        window.__sent.push(msg);
        let resp = {};
        if (msg.action === 'openResults') resp = { success: true };
        if (msg.action === 'getTranscriptAPIMeta') resp = {};
        if (cb) { setTimeout(() => cb(resp), 0); return undefined; }
        return Promise.resolve(resp);
      },
      onMessage: { addListener: (fn) => listeners.push(fn) }
    }
  };
  // Deliver a runtime message like chrome.tabs.sendMessage would to this frame.
  window.__dispatch = (req, waitMs = 3000) => new Promise((resolve) => {
    let done = false;
    const returned = listeners.map((fn) => fn(req, {}, (r) => { if (!done) { done = true; resolve({ response: r, at: Date.now() }); } }));
    setTimeout(() => { if (!done) resolve({ response: null, returned }); }, waitMs);
  });

  // ---- page-world mocks ----
  // Batch transcript script (batchTranscriptDownload.js is routed to an empty file).
  let cancelled = false;
  document.addEventListener('teamsBatchTranscriptCommand', async (e) => {
    const { command } = e.detail || {};
    const reply = (result) => document.dispatchEvent(new CustomEvent('teamsBatchTranscriptResponse', { detail: { command, result } }));
    if (command === 'cancel') { cancelled = true; reply({ cancelled: true }); return; }
    if (command === 'status') { reply({ running: false }); return; }
    if (command !== 'start') return;
    cancelled = false;
    const meetings = ['Oct 1, 2026', 'Sep 24, 2026', 'Sep 17, 2026', 'Sep 10, 2026', 'Sep 3, 2026', 'Aug 27, 2026', 'Aug 20, 2026', 'Aug 13, 2026'];
    const total = meetings.length;
    const prog = (d) => document.dispatchEvent(new CustomEvent('teamsBatchTranscriptProgress', { detail: d }));
    const stopAt = window.__batchStopAt || total;
    let found = 0;
    for (let i = 0; i < stopAt; i++) {
      await new Promise((r) => setTimeout(r, 60));
      const has = i % 4 !== 3;
      if (has) found++;
      prog({ phase: has ? 'extracted' : 'skipped', total, current: i + 1, currentMeeting: meetings[i], entryCount: 120 + i * 7, source: i % 2 ? 'api' : 'dom', transcriptsFound: found, message: `Processing ${meetings[i]} (${i + 1}/${total})` });
    }
    if (window.__batchHold) await window.__batchHold;
    prog({ phase: 'complete', total, transcriptsFound: found, apiSuccessCount: 4, domFallbackCount: found - 4, message: `Done! ${found}/${total} had transcripts` });
    reply({
      success: true,
      seriesName: '週次デザインレビュー: Product/Design',
      results: meetings.map((m, i) => ({ meetingDate: m, index: i, hasTranscript: i % 4 !== 3, vtt: `WEBVTT\n\n1\n00:00:00.000 --> 00:00:02.000\n<v 田中>会議 ${m}</v>\n`, txt: `[00:00:00] 田中: 会議 ${m}`, entryCount: 3 }))
    });
  });

  // Video capture override (videoDownloadOverride.js routed to empty).
  document.addEventListener('teamsVideoCommand', (e) => {
    const { command } = e.detail || {};
    const reply = (result) => document.dispatchEvent(new CustomEvent('teamsVideoResponse', { detail: { command, result } }));
    const results = {
      ping: { available: true },
      getStats: { videoCount: 412, audioCount: 412 },
      clear: { ok: true }, startCapture: { ok: true }, stopCapture: { ok: true },
      downloadFiles: { success: true, title: '週次デザインレビュー 2026-10-01', videoCount: 412, audioCount: 412 },
      analyzeUrls: { capturedUrls: ['a', 'b'] }
    };
    setTimeout(() => reply(results[command] || { error: 'mock: unknown ' + command }), 20);
  });

  // Fake download module for the coordinator. Controlled from the test via
  // window.__dl = { resolve, reject } once download() starts.
  window.__videoDownloadModules = window.__videoDownloadModules || {};
  window.__videoDownloadModules.captureStreamDownload = {
    name: 'captureStreamDownload',
    label: 'Record Stream',
    description: 'mock',
    isAvailable: () => true,
    stop: () => { window.__dlStopped = true; },
    download: (onProgress) => new Promise((resolve, reject) => {
      onProgress({ stage: 'recording', message: 'Recording at 8x… 14:32 of 32:10 captured', percent: 45 });
      window.__dl = { resolve, reject, onProgress };
    })
  };
})();
