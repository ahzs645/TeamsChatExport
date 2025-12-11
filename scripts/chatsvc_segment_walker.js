// Safer segment walker: fetches a few segments sequentially with delays, following _metadata.syncState.
// Configure via localStorage before running in the Teams console (top frame):
//   teamsChatConversationIdOverride (required)
//   teamsChatHostOverride (default teams.microsoft.com)
//   teamsChatRegionOverride (default ca)
//   teamsChatSegmentMaxSteps (default 5)
//   teamsChatSegmentPageSize (default 30)
//   teamsChatSegmentDelayMs (default 1200)
//   teamsChatSegmentStartTime (default 1, or paste lastCompleteSegmentStartTime)
//   teamsChatSegmentInitialSyncState (optional: resume from a known syncState)
(() => {
  const log = (...args) => console.log('[SegmentWalker]', ...args);
  const warn = (...args) => console.warn('[SegmentWalker]', ...args);

  const readNum = (key, fallback, { min = 1, max = 1e9 } = {}) => {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return fallback;
      const n = parseInt(raw, 10);
      if (Number.isNaN(n)) return fallback;
      return Math.min(Math.max(n, min), max);
    } catch (_e) {
      return fallback;
    }
  };

  const convId =
    localStorage.getItem('teamsChatConversationIdOverride') ||
    localStorage.getItem('teamsChatApiConversationIdOverride') ||
    localStorage.getItem('teamsChatApiConversationId') ||
    '';
  if (!convId) {
    warn('No conversation override set. Set teamsChatConversationIdOverride first.');
    return;
  }

  const host = localStorage.getItem('teamsChatHostOverride') || 'teams.microsoft.com';
  const region = localStorage.getItem('teamsChatRegionOverride') || 'ca';
  const maxSteps = readNum('teamsChatSegmentMaxSteps', 5, { min: 1, max: 50 });
  const pageSize = readNum('teamsChatSegmentPageSize', 30, { min: 1, max: 200 });
  const delayMs = readNum('teamsChatSegmentDelayMs', 1200, { min: 0, max: 5000 });
  const initialStart = readNum('teamsChatSegmentStartTime', 1, { min: 1, max: Number.MAX_SAFE_INTEGER });
  const initialSync = (() => {
    try { return localStorage.getItem('teamsChatSegmentInitialSyncState') || null; } catch (_e) { return null; }
  })();

  const token = localStorage.getItem('teamsChatAuthToken') || window.__teamsAuthToken;
  const headers = {
    'Accept': 'application/json',
    'x-ms-client-type': 'web',
    'x-ms-request-priority': '0',
    'x-ms-migration': 'True',
    'behavioroverride': 'redirectAs404',
    'x-ms-client-version': '1415/25110202315',
    'clientinfo': 'os=mac; osVer=10.15.7; proc=x86; lcid=en-us; deviceType=1; country=us; clientName=skypeteams; clientVer=1415/25110202315; utcOffset=-08:00; timezone=America/Vancouver'
  };
  if (token) {
    const lower = token.toLowerCase();
    headers.Authorization = (lower.startsWith('bearer ') || lower.startsWith('skype_token')) ? token : `Bearer ${token}`;
  }

  log('convId:', convId);
  log('host/region:', host, region);
  log('maxSteps/pageSize/delay/start:', maxSteps, pageSize, delayMs, initialStart);
  if (initialSync) log('initial syncState override present (will start with it)');

  const base = `https://${host}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(convId)}/messages`;
  let currentSync = initialSync;
  let currentStart = initialStart;
  let steps = 0;
  const dump = [];
  let earliest = null;
  let latest = null;

  const parseIso = (value) => {
    if (!value) return null;
    const t = Date.parse(value);
    return Number.isNaN(t) ? null : new Date(t).toISOString();
  };

  const msgIso = (msg) => {
    const ts =
      msg?.createdDateTime ||
      msg?.originalArrivalTime ||
      msg?.composedTime ||
      msg?.sentTime ||
      msg?.created ||
      msg?.createdTime ||
      msg?.composetime ||
      msg?.originalarrivaltime ||
      msg?.timestamp;
    return parseIso(ts);
  };

  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  (async () => {
    while (steps < maxSteps) {
      steps += 1;
      const urlObj = new URL(base);
      urlObj.searchParams.set('view', 'msnp24Equivalent|supportsMessageProperties');
      urlObj.searchParams.set('pageSize', pageSize.toString());
      urlObj.searchParams.set('startTime', currentStart.toString());
      if (currentSync) {
        urlObj.searchParams.set('syncState', currentSync);
      }

      log(`step ${steps} -> ${urlObj.toString()}`);
      let data;
      try {
        const resp = await fetch(urlObj.toString(), { credentials: 'include', headers });
        const text = await resp.text();
        if (!resp.ok) {
          warn(`HTTP ${resp.status} ${resp.statusText}`, text.slice(0, 400));
          break;
        }
        try {
          data = JSON.parse(text);
        } catch (err) {
          warn('JSON parse error', err.message, text.slice(0, 300));
          break;
        }
      } catch (err) {
        warn('fetch/parse error', err.message);
        break;
      }

      const msgs = Array.isArray(data?.messages) ? data.messages : [];
      dump.push(...msgs);

      const isoList = msgs.map(msgIso).filter(Boolean).sort();
      if (isoList[0]) earliest = earliest ? new Date(Math.min(Date.parse(earliest), Date.parse(isoList[0]))).toISOString() : isoList[0];
      if (isoList[isoList.length - 1]) latest = latest ? new Date(Math.max(Date.parse(latest), Date.parse(isoList[isoList.length - 1]))).toISOString() : isoList[isoList.length - 1];

      const metaSync = data?.syncState || data?._metadata?.syncState || null;
      const metaStart = data?._metadata?.lastCompleteSegmentStartTime || currentStart;
      const metaEnd = data?._metadata?.lastCompleteSegmentEndTime || null;

      log(`step ${steps}: msgs=${msgs.length}, metaStart=${metaStart || 'n/a'}, metaEnd=${metaEnd || 'n/a'}, syncState=${metaSync ? metaSync.slice(0, 24) + '...' : '(none)'}, earliest=${isoList[0] || 'n/a'}, latest=${isoList[isoList.length - 1] || 'n/a'}`);

      if (!metaSync) {
        break;
      }

      currentSync = metaSync;
      currentStart = metaStart;

      if (delayMs > 0) {
        await delay(delayMs);
      }
    }

    try {
      window.__teamsSegmentDump = dump;
    } catch (_e) {}

    log(`done. steps=${steps}, total msgs=${dump.length}, earliest=${earliest || 'n/a'}, latest=${latest || 'n/a'}`);
  })();
})();
