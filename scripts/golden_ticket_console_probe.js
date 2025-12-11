// Run this in the Teams page console (top frame). It gathers token/host/id info
// and performs the "golden ticket" chatsvc call with the required view/startTime params.
(() => {
  const log = (...args) => console.log('[GoldenProbe]', ...args);
  const warn = (...args) => console.warn('[GoldenProbe]', ...args);

  const decodeJwt = (token) => {
    if (!token) return null;
    try {
      const clean = token.startsWith('Bearer ') ? token.slice(7) : token;
      const parts = clean.split('.');
      if (parts.length !== 3) return null;
      const payload = atob(parts[1].replace(/-/g, '+').replace(/_/g, '/'));
      const json = payload.substring(0, payload.lastIndexOf('}') + 1);
      return JSON.parse(json);
    } catch (_err) {
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
          } catch (_err) {
            // ignore parse errors
          }
        }
      }
    }
    return token;
  };

  const detectHostAndRegion = () => {
    const overrides = { host: null, region: null };
    try {
      overrides.host = localStorage.getItem('teamsChatHostOverride') || null;
      overrides.region = localStorage.getItem('teamsChatRegionOverride') || null;
    } catch (_err) {}

    const entries = performance.getEntriesByType('resource') || [];
    const urls = entries
      .map((e) => e.name)
      .filter((u) => /chatsvc|chatservice/.test(u) && u.includes('/messages'))
      .slice(-6);

    const hosts = Array.from(new Set(urls.map((u) => {
      try { return new URL(u).host; } catch (_e) { return null; }
    }).filter(Boolean)));

    const regionMatch = urls.map((u) => u.match(/chatsvc\/([^/]+)\/v1/)).find(Boolean);
    return {
      hosts,
      region: overrides.region || (regionMatch ? regionMatch[1] : null),
      samples: urls,
      overrideHost: overrides.host
    };
  };

  const findConversationId = () => {
    try {
      const override =
        localStorage.getItem('teamsChatConversationIdOverride') ||
        localStorage.getItem('teamsChatApiConversationId') ||
        localStorage.getItem('teamsChatApiConversationIdOverride');
      if (override && override.includes('19:')) return override;
    } catch (_err) {}

    // URL hint
    const hrefMatch = (location.href || '').match(/(19:[^/?#"]+)/);
    if (hrefMatch) return hrefMatch[1];

    // Active chat element
    const active = document.querySelector('[data-tid="chat-list-item"][aria-selected="true"], [aria-selected="true"][data-qa-chat-id], [role="treeitem"][aria-selected="true"]');
    if (active && active.dataset) {
      const hit = Object.values(active.dataset).find((v) => typeof v === 'string' && v.includes('19:'));
      if (hit) return hit;
    }

    // Recent resource entry
    const perf = performance.getEntriesByType('resource') || [];
    for (let i = perf.length - 1; i >= 0; i--) {
      const name = perf[i].name || '';
      const match = name.match(/(19:[^/?#"]+)/);
      if (match) return match[1];
    }

    return null;
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

  const tryFetch = async (url, headers) => {
    try {
      const resp = await fetch(url, { headers, credentials: 'include' });
      const text = await resp.text();
      log('fetch', url, '->', resp.status, resp.statusText);
      log('body (trimmed 500):', text.slice(0, 500));
      return { status: resp.status, body: text };
    } catch (err) {
      warn('fetch error', url, err.message);
      return { status: 0, body: err.message };
    }
  };

  (async () => {
    const token = findToken();
    const payload = decodeJwt(token);
    log('token present:', !!token, token ? token.slice(0, 40) + '...' : '');
    if (payload) {
      log('token aud:', payload.aud, 'exp:', payload.exp ? new Date(payload.exp * 1000).toISOString() : 'n/a');
    } else if (token) {
      warn('token present but could not decode (maybe Skype token)');
    } else {
      warn('no token found');
    }

    const convId = findConversationId();
    log('conversationId:', convId || '(not found)');

    const { hosts, region, samples, overrideHost } = detectHostAndRegion();
    log('chatsvc hosts seen:', hosts);
    if (overrideHost) log('override host:', overrideHost);
    log('sample message URLs:', samples);
    const host = overrideHost || hosts[hosts.length - 1] || 'teams.microsoft.com';
    const geo = region || 'ca';
    const headers = buildHeaders(token);

    if (!convId) {
      warn('no conversation id; open a chat and rerun');
      return;
    }

    const encodedId = encodeURIComponent(convId);
    const base = `https://${host}/api/chatsvc/${geo}/v1`;
    const urls = [
      `${base}/users/ME/conversations/${encodedId}/messages?view=msnp24Equivalent|supportsMessageProperties&startTime=1&pageSize=5`,
      `${base}/threads/${encodedId}/messages?view=msnp24Equivalent|supportsMessageProperties&startTime=1&pageSize=5`
    ];

    for (const u of urls) {
      await tryFetch(u, headers);
    }
  })();
})();
