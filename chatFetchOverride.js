(() => {
  const { fetch: originalFetch } = window;
  const cacheKey = '__teamsExtractorMessageApiCache';
  const hiddenDivId = 'teams-extractor-message-api-cache';
  const MAX_CACHE_ENTRIES = 50;
  const debugEnabled = (() => {
    try {
      const flag = window.localStorage?.getItem('teamsChatApiDebug');
      if (flag === '0') return false;
      if (flag === '1') return true;
    } catch (_err) {
      // ignore storage errors
    }
    return true; // enable lightweight debug by default
  })();
  let debugSeen = 0;
  const DEBUG_LIMIT = 40;

  const ensureCache = () => {
    if (!Array.isArray(window[cacheKey])) {
      window[cacheKey] = [];
    }
    return window[cacheKey];
  };

  const ensureHiddenDiv = () => {
    const existing = document.getElementById(hiddenDivId);
    if (existing) {
      return existing;
    }
    const div = document.createElement('div');
    div.id = hiddenDivId;
    div.style.display = 'none';
    document.body.appendChild(div);
    return div;
  };

  const isCandidateUrl = (url) => {
    const lower = (url || '').toLowerCase();
    return [
      '/conversations/',
      '/messages',
      '/api/chatsvc/',
      'chatsvc',
      'chatservice',
      'graph.microsoft.com',
      'users/me/conversations'
    ].some((needle) => lower.includes(needle));
  };

  const pushEntry = (url, payload) => {
    const cache = ensureCache();
    cache.push({
      url,
      timestamp: Date.now(),
      payload
    });

    // Trim cache to avoid unbounded growth
    if (cache.length > MAX_CACHE_ENTRIES) {
      cache.splice(0, cache.length - MAX_CACHE_ENTRIES);
    }

    // Keep a copy in a hidden div for the content script context
    try {
      const div = ensureHiddenDiv();
      div.textContent = JSON.stringify(cache);
    } catch (err) {
      console.warn('[Teams Chat API capture] Failed to serialize cache', err);
    }
  };

  const logDebug = (msg, ...rest) => {
    if (!debugEnabled) return;
    if (debugSeen > DEBUG_LIMIT) return;
    debugSeen += 1;
    console.debug('[Teams Chat API capture]', msg, ...rest);
  };

  const capturePayload = (url, payload) => {
    pushEntry(url, payload);
    logDebug('captured', url);
  };

  console.log('[Teams Chat Extractor] Override v3 loaded');

  // Clear potentially bad token on startup
  try {
    const existing = window.localStorage.getItem('teamsChatAuthToken');
    if (existing && !isValidTeamsToken(existing)) {
      console.log('[Teams Chat Extractor] Clearing invalid token on startup');
      window.localStorage.removeItem('teamsChatAuthToken');
      window.__teamsAuthToken = null;
    }
  } catch (_e) { }

  function isValidTeamsToken(token) {
    if (!token) return false;
    if (token.toLowerCase().startsWith('skype_token ')) return true;

    try {
      const cleanToken = token.startsWith('Bearer ') ? token.substring(7) : token;
      const parts = cleanToken.split('.');
      if (parts.length !== 3) return false;

      let base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
      const pad = base64.length % 4;
      if (pad) {
        if (pad === 1) throw new Error('InvalidLengthError');
        base64 += new Array(5 - pad).join('=');
      }

      const decoded = atob(base64);
      const lastBrace = decoded.lastIndexOf('}');
      if (lastBrace === -1) return false;

      const jsonStr = decoded.substring(0, lastBrace + 1);
      const payload = JSON.parse(jsonStr);

      if (payload.aud === 'https://ic3.teams.office.com') return true;

      // console.debug('Invalid audience:', payload.aud);
      return false;
    } catch (e) {
      return false;
    }
  }

  const captureToken = (headers, sourceUrl) => {
    if (!headers) return;
    let token;
    if (typeof headers.get === 'function') {
      token = headers.get('Authorization') || headers.get('authorization');
    } else if (typeof headers === 'object') {
      token = headers.Authorization || headers.authorization;
    }

    if (token) {
      const lower = token.toLowerCase();
      const isBearer = lower.startsWith('bearer ');
      const isSkype = lower.startsWith('skype_token ');

      if (isBearer || isSkype) {
        // Validate audience if it's a Bearer token
        if (isBearer) {
          if (!isValidTeamsToken(token.substring(7))) {
            // console.debug('[Teams Chat Extractor] Ignoring token from', sourceUrl, '(wrong audience)');
            return;
          }
        }

        console.log('[Teams Chat Extractor] 🟢 Captured valid token from', sourceUrl);

        // Prefer Bearer token. Only overwrite if new is Bearer, or if current is not Bearer.
        const current = window.__teamsAuthToken || '';
        const currentIsBearer = current.toLowerCase().startsWith('bearer ');

        if (isBearer || !currentIsBearer) {
          window.__teamsAuthToken = token;
          try {
            window.localStorage.setItem('teamsChatAuthToken', token);
          } catch (_e) { }
        }
      }
    }
  };

  window.fetch = async (...args) => {
    const [resource, config] = args;
    const url = typeof resource === 'string' ? resource : resource?.url || '';

    // Capture token from ALL outgoing fetches
    if (config && config.headers) {
      captureToken(config.headers, url);
    }

    const response = await originalFetch(...args);

    if (isCandidateUrl(url)) {
      const clone = response.clone();
      clone
        .json()
        .then((data) => capturePayload(url, data))
        .catch((err) => logDebug('non-JSON or parse failed for fetch', url, err));
    }

    return response;
  };

  // Hook Response.json so we capture even if the page grabbed fetch before our override
  if (window.Response && typeof window.Response.prototype.json === 'function') {
    const originalJson = window.Response.prototype.json;
    window.Response.prototype.json = function (...jsonArgs) {
      try {
        const url = this.url || '';
        if (isCandidateUrl(url)) {
          this.clone()
            .json()
            .then((data) => capturePayload(url, data))
            .catch((_err) => {
              // ignore clone parse errors
            });
        }
      } catch (_err) {
        // ignore
      }
      return originalJson.apply(this, jsonArgs);
    };
  }

  // Hook Response.text as a fallback
  if (window.Response && typeof window.Response.prototype.text === 'function') {
    const originalText = window.Response.prototype.text;
    window.Response.prototype.text = function (...textArgs) {
      try {
        const url = this.url || '';
        if (isCandidateUrl(url)) {
          this.clone()
            .text()
            .then((body) => {
              try {
                const parsed = JSON.parse(body);
                capturePayload(url, parsed);
              } catch (_err) {
                // not JSON
              }
            })
            .catch((_err) => {
              // ignore
            });
        }
      } catch (_err) {
        // ignore
      }
      return originalText.apply(this, textArgs);
    };
  }

  // Hook XMLHttpRequest as well (some Teams variants use XHR)
  if (window.XMLHttpRequest) {
    const originalOpen = window.XMLHttpRequest.prototype.open;
    const originalSend = window.XMLHttpRequest.prototype.send;
    const originalSetRequestHeader = window.XMLHttpRequest.prototype.setRequestHeader;

    window.XMLHttpRequest.prototype.open = function (...openArgs) {
      try {
        this.__teamsChatApiUrl = openArgs[1];
      } catch (_err) {
        // ignore
      }
      return originalOpen.apply(this, openArgs);
    };

    window.XMLHttpRequest.prototype.setRequestHeader = function (header, value) {
      // Capture token from ALL XHRs
      if (header && (header.toLowerCase() === 'authorization')) {
        if (value) {
          const lower = value.toLowerCase();
          const isBearer = lower.startsWith('bearer ');
          const isSkype = lower.startsWith('skype_token ');

          if (isBearer || isSkype) {
            // Validate audience if it's a Bearer token
            if (isBearer && !isValidTeamsToken(value.substring(7))) {
              // console.debug('[Teams Chat Extractor] Ignoring XHR token (wrong audience)');
            } else {
              console.log('[Teams Chat Extractor] 🟢 Captured valid XHR token');
              const current = window.__teamsAuthToken || '';
              const currentIsBearer = current.toLowerCase().startsWith('bearer ');

              if (isBearer || !currentIsBearer) {
                window.__teamsAuthToken = value;
                try {
                  window.localStorage.setItem('teamsChatAuthToken', value);
                } catch (_e) { }
              }
            }
          }
        }
      }
      return originalSetRequestHeader.apply(this, arguments);
    };

    window.XMLHttpRequest.prototype.send = function (...sendArgs) {
      if (this.__teamsChatApiUrl) {
        this.addEventListener('loadend', () => {
          const finalUrl = this.responseURL || this.__teamsChatApiUrl;
          if (!isCandidateUrl(finalUrl)) return;

          if (this.responseType === '' || this.responseType === 'text' || this.responseType === 'json') {
            const body = this.response || this.responseText;
            try {
              const parsed = typeof body === 'string' ? JSON.parse(body) : body;
              capturePayload(finalUrl, parsed);
            } catch (_err) {
              // not JSON, ignore
            }
          }
        });
      }
      return originalSend.apply(this, sendArgs);
    };
  }
})();
