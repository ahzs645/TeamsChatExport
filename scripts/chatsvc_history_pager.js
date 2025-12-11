// Run this in the Teams page console (top frame).
// It paginates the chatsvc messages API using syncState to pull as far back as possible
// and reports earliest/latest timestamps and total count retrieved.
(() => {
  const log = (...args) => console.log('[HistoryPager]', ...args);
  const warn = (...args) => console.warn('[HistoryPager]', ...args);

  const MAX_PAGES = (() => {
    try {
      const val = localStorage.getItem('teamsChatPagerMaxPages');
      const n = val ? parseInt(val, 10) : NaN;
      if (!Number.isNaN(n) && n > 0 && n <= 200) return n;
    } catch (_e) {}
    return 40;
  })();
  const PAGE_SIZE = (() => {
    try {
      const val = localStorage.getItem('teamsChatPagerPageSize');
      const n = val ? parseInt(val, 10) : NaN;
      if (!Number.isNaN(n) && n > 0 && n <= 500) return n;
    } catch (_e) {}
    return 200;
  })();
  const START_TIME = (() => {
    try {
      const val = localStorage.getItem('teamsChatPagerStartTime');
      const n = val ? parseInt(val, 10) : NaN;
      if (!Number.isNaN(n) && n > 0) return n;
    } catch (_e) {}
    return 1;
  })();
  const PAGE_DELAY_MS = (() => {
    try {
      const val = localStorage.getItem('teamsChatPagerDelayMs');
      const n = val ? parseInt(val, 10) : NaN;
      if (!Number.isNaN(n) && n >= 0 && n <= 5000) return n;
    } catch (_e) {}
    return 250; // small pause to avoid thrashing the page
  })();

  const getOverrides = () => {
    const convId =
      localStorage.getItem('teamsChatConversationIdOverride') ||
      localStorage.getItem('teamsChatApiConversationIdOverride') ||
      localStorage.getItem('teamsChatApiConversationId') ||
      '';
    const host = localStorage.getItem('teamsChatHostOverride') || 'teams.microsoft.com';
    const region = localStorage.getItem('teamsChatRegionOverride') || 'ca';
    return { convId, host, region };
  };

  const decodeJwt = (token) => {
    if (!token) return null;
    try {
      const clean = token.startsWith('Bearer ') ? token.slice(7) : token;
      const payload = atob(clean.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'));
      const json = payload.substring(0, payload.lastIndexOf('}') + 1);
      return JSON.parse(json);
    } catch (_e) {
      return null;
    }
  };

  const findToken = () => {
    let token = window.__teamsAuthToken || localStorage.getItem('teamsChatAuthToken');
    if (!token) {
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key.includes('accesstoken') && key.includes('ic3.teams.office.com')) {
          try {
            const val = JSON.parse(localStorage.getItem(key));
            if (val?.secret) {
              token = val.secret;
              break;
            }
          } catch (_err) {}
        }
      }
    }
    return token;
  };

  const buildHeaders = (token) => {
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
    return headers;
  };

  const parseIso = (value) => {
    if (!value) return null;
    const t = Date.parse(value);
    return Number.isNaN(t) ? null : new Date(t).toISOString();
  };

  const getMsgIso = (msg) => {
    const ts =
      msg.createdDateTime ||
      msg.originalArrivalTime ||
      msg.composedTime ||
      msg.sentTime ||
      msg.created ||
      msg.createdTime ||
      msg.composetime ||
      msg.originalarrivaltime ||
      msg.timestamp;
    return parseIso(ts);
  };

  (async () => {
    const { convId, host, region } = getOverrides();
    if (!convId) {
      warn('No conversation override found. Set teamsChatConversationIdOverride first.');
      return;
    }

    const token = findToken();
    const payload = decodeJwt(token);
    log('convId:', convId);
    log('host/region:', host, region);
    log('token present:', !!token, payload ? `aud=${payload.aud}` : '(decode failed or Skype token)');

    const base = `https://${host}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(convId)}/messages`;
    let syncState = null;
    let page = 0;
    let total = 0;
    let earliest = null;
    let latest = null;
    const headers = buildHeaders(token);

    while (page < MAX_PAGES) {
      page += 1;
      const urlObj = new URL(base);
      urlObj.searchParams.set('view', 'msnp24Equivalent|supportsMessageProperties');
      urlObj.searchParams.set('pageSize', PAGE_SIZE.toString());
      urlObj.searchParams.set('startTime', START_TIME.toString());
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

      const pageIso = msgs
        .map((m) => getMsgIso(m))
        .filter(Boolean)
        .sort();
      if (pageIso[0]) earliest = earliest ? Math.min(Date.parse(earliest), Date.parse(pageIso[0])) && new Date(Math.min(Date.parse(earliest), Date.parse(pageIso[0]))).toISOString() : pageIso[0];
      if (pageIso[pageIso.length - 1]) latest = latest ? Math.max(Date.parse(latest), Date.parse(pageIso[pageIso.length - 1])) && new Date(Math.max(Date.parse(latest), Date.parse(pageIso[pageIso.length - 1]))).toISOString() : pageIso[pageIso.length - 1];

      const metaSync = data?._metadata?.syncState;
      const syncForLog = data?.syncState || metaSync;
      log(`page ${page}: msgs=${msgs.length}, syncState=${syncForLog ? syncForLog.slice(0, 24) + '...' : '(none)'}, page earliest=${pageIso[0] || 'n/a'}, latest=${pageIso[pageIso.length - 1] || 'n/a'}`);

      syncState = data?.syncState || metaSync || null;
      // Continue if syncState is present, even when msgs.length === 0
      if (!syncState) {
        break;
      }

      // Brief delay to avoid overloading the page
      if (PAGE_DELAY_MS > 0) {
        await new Promise((resolve) => setTimeout(resolve, PAGE_DELAY_MS));
      }
    }

    log(`done. pages=${page}, total msgs=${total}, earliest=${earliest || 'n/a'}, latest=${latest || 'n/a'}`);
  })();
})();
