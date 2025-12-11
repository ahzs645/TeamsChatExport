// Safer, throttled pager for chatsvc messages. Default: small pages & delays to avoid crashes.
// Reads overrides from localStorage:
// - teamsChatConversationIdOverride (required)
// - teamsChatHostOverride (default teams.microsoft.com)
// - teamsChatRegionOverride (default ca)
// - teamsChatPagerMaxPages (default 5)
// - teamsChatPagerPageSize (default 50)
// - teamsChatPagerStartTime (default 1)
// - teamsChatPagerDelayMs (default 1000)
// - teamsChatPagerInitialSyncState (optional: start from a known syncState)
(async () => {
  const log = (...args) => console.log('[HistoryPagerSafe]', ...args);
  const warn = (...args) => console.warn('[HistoryPagerSafe]', ...args);

  const readNumber = (key, fallback, { min = 1, max = 1e9 } = {}) => {
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
  const host = localStorage.getItem('teamsChatHostOverride') || 'teams.microsoft.com';
  const region = localStorage.getItem('teamsChatRegionOverride') || 'ca';
  const maxPages = readNumber('teamsChatPagerMaxPages', 5, { min: 1, max: 200 });
  const pageSize = readNumber('teamsChatPagerPageSize', 50, { min: 1, max: 500 });
  const startTime = readNumber('teamsChatPagerStartTime', 1, { min: 1, max: Number.MAX_SAFE_INTEGER });
  const delayMs = readNumber('teamsChatPagerDelayMs', 1000, { min: 0, max: 5000 });
  const initialSync = (() => {
    try { return localStorage.getItem('teamsChatPagerInitialSyncState') || null; } catch (_e) { return null; }
  })();

  if (!convId) {
    warn('No conversation override found. Set teamsChatConversationIdOverride first.');
    return;
  }

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
  log('maxPages/pageSize/delay/start:', maxPages, pageSize, delayMs, startTime);
  if (initialSync) log('initial syncState override present (will start with it)');

  const base = `https://${host}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(convId)}/messages`;
  let syncState = initialSync || null;
  let page = 0;
  let total = 0;
  let earliest = null;
  let latest = null;
  const dump = [];

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

  while (page < maxPages) {
    page += 1;
    const urlObj = new URL(base);
    urlObj.searchParams.set('view', 'msnp24Equivalent|supportsMessageProperties');
    urlObj.searchParams.set('pageSize', pageSize.toString());
    urlObj.searchParams.set('startTime', startTime.toString());
    if (syncState) {
      urlObj.searchParams.set('syncState', syncState);
    }

    log(`page ${page} -> ${urlObj.toString()}`);
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
    total += msgs.length;
    dump.push(...msgs);

    const isoList = msgs.map(msgIso).filter(Boolean).sort();
    if (isoList[0]) {
      earliest = earliest ? new Date(Math.min(Date.parse(earliest), Date.parse(isoList[0]))).toISOString() : isoList[0];
    }
    if (isoList[isoList.length - 1]) {
      latest = latest ? new Date(Math.max(Date.parse(latest), Date.parse(isoList[isoList.length - 1]))).toISOString() : isoList[isoList.length - 1];
    }

    const metaSync = data?._metadata?.syncState;
    const syncForLog = data?.syncState || metaSync;
    const metaStart = data?._metadata?.lastCompleteSegmentStartTime;
    const metaEnd = data?._metadata?.lastCompleteSegmentEndTime;
    log(`page ${page}: msgs=${msgs.length}, syncState=${syncForLog ? syncForLog.slice(0, 24) + '...' : '(none)'}, metaStart=${metaStart || 'n/a'}, metaEnd=${metaEnd || 'n/a'}, earliest=${isoList[0] || 'n/a'}, latest=${isoList[isoList.length - 1] || 'n/a'}`);

    syncState = data?.syncState || metaSync || null;
    if (!syncState) {
      break;
    }

    if (delayMs > 0) {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
  }

  try {
    window.__teamsChatPagerDump = dump;
  } catch (_e) {}

  log(`done. pages=${page}, total msgs=${total}, earliest=${earliest || 'n/a'}, latest=${latest || 'n/a'}`);
})();
