/**
 * Api Message Extractor
 * Parses captured fetch responses (Teams chat APIs) into the standard message shape.
 * This is best-effort and gated via localStorage "teamsChatApiMode".
 */

export class ApiMessageExtractor {
  constructor() {
    this.cacheKey = '__teamsExtractorMessageApiCache';
    this.hiddenDivId = 'teams-extractor-message-api-cache';
    this.conversationStartTs = 0;
  }

  readNumberSetting(key, { min = 1, max = 500, fallback }) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return fallback;
      const n = parseInt(raw, 10);
      if (Number.isNaN(n)) return fallback;
      return Math.min(Math.max(n, min), max);
    } catch (_err) {
      return fallback;
    }
  }

  getHostAndRegion() {
    const defaults = { host: 'teams.microsoft.com', region: 'ca' };
    try {
      const host = localStorage.getItem('teamsChatHostOverride') || defaults.host;
      const region = localStorage.getItem('teamsChatRegionOverride') || defaults.region;
      return { host, region };
    } catch (_err) {
      return defaults;
    }
  }

  isEnabled() {
    try {
      return window.localStorage?.getItem('teamsChatApiMode') === '1';
    } catch (_err) {
      return false;
    }
  }

  startConversationWindow() {
    this.conversationStartTs = Date.now();
  }

  readCache() {
    const inMemory = window[this.cacheKey];
    if (Array.isArray(inMemory)) {
      return inMemory;
    }
    const hidden = document.getElementById(this.hiddenDivId);
    if (hidden && hidden.textContent) {
      try {
        const parsed = JSON.parse(hidden.textContent);
        if (Array.isArray(parsed)) {
          window[this.cacheKey] = parsed;
          return parsed;
        }
      } catch (_err) {
        // ignore parse errors
      }
    }
    return [];
  }

  extractMessagesSince(timestampMs) {
    const cache = this.readCache().filter((entry) => entry && entry.timestamp >= timestampMs);
    if (!cache.length) {
      return [];
    }

    const messages = [];
    cache.forEach((entry) => {
      const payloadMessages = this.extractFromPayload(entry.payload);
      payloadMessages.forEach((msg) => messages.push(msg));
    });

    return this.sortAndDedup(messages);
  }

  async fetchConversation(conversationId, { pageSize = 200, maxPages = 15, startTime = 1 } = {}) {
    if (!conversationId) return [];

    // Allow overrides via localStorage
    const effectivePageSize = this.readNumberSetting('teamsChatApiPageSize', { min: 1, max: 500, fallback: pageSize });
    const effectiveMaxPages = this.readNumberSetting('teamsChatApiMaxPages', { min: 1, max: 200, fallback: maxPages });
    const effectiveStartTime = this.readNumberSetting('teamsChatApiStartTime', { min: 1, max: Number.MAX_SAFE_INTEGER, fallback: startTime });

    const { host, region } = this.getHostAndRegion();
    const base = `https://${host}/api/chatsvc/${region}/v1`;
    // Try conversations endpoint first, then threads endpoint as fallback.
    const encodedId = encodeURIComponent(conversationId);
    const paths = [
      `${base}/users/ME/conversations/${encodedId}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=${effectivePageSize}&startTime=${effectiveStartTime}`,
      `${base}/threads/${encodedId}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=${effectivePageSize}&startTime=${effectiveStartTime}`
    ];

    for (const path of paths) {
      const msgs = await this.fetchPaged(path, { maxPages: effectiveMaxPages });
      if (msgs.length > 0) {
        return msgs;
      }
    }

    return [];
  }

  async fetchPaged(initialUrl, { maxPages = 15 } = {}) {
    const messages = [];
    const baseUrl = new URL(initialUrl);
    let nextSyncState = baseUrl.searchParams.get('syncState') || null;

    // Get token - try multiple sources
    console.log('[Teams Chat API] Looking for auth token...');

    let token = null;

    // === FIRST: Try context bridge (most reliable) ===
    const bridge = document.getElementById('teams-extractor-context-bridge');
    if (bridge) {
      const bridgeToken = bridge.getAttribute('data-token');
      if (bridgeToken) {
        token = bridgeToken;
        console.log('[Teams Chat API] ✅ Found token in context bridge, length:', token.length);
      }
    }

    // Try direct localStorage
    if (!token) {
      try {
        token = localStorage.getItem('teamsChatAuthToken');
        if (token) {
          console.log('[Teams Chat API] ✅ Found token in localStorage.teamsChatAuthToken');
        }
      } catch (e) {
        console.log('[Teams Chat API] localStorage access error:', e.message);
      }
    }

    // Try window.__teamsAuthToken (only works if we're in page context)
    if (!token && typeof window !== 'undefined' && window.__teamsAuthToken) {
      token = window.__teamsAuthToken;
      console.log('[Teams Chat API] ✅ Found token in window.__teamsAuthToken');
    }

    // Try MSAL storage scan
    if (!token) {
      console.log('[Teams Chat API] Scanning MSAL storage for token...');
      token = this.findTokenInStorage();
      if (token) {
        console.log('[Teams Chat API] ✅ Found token in MSAL storage');
      }
    }

    // Wait and retry if still missing
    if (!token) {
      console.log('[Teams Chat API] Token not found, waiting for capture...');
      for (let i = 0; i < 10; i++) {
        await new Promise(r => setTimeout(r, 200));
        // Check bridge again
        const bridgeRetry = document.getElementById('teams-extractor-context-bridge');
        if (bridgeRetry) {
          token = bridgeRetry.getAttribute('data-token');
          if (token) {
            console.log('[Teams Chat API] ✅ Found token in bridge after waiting');
            break;
          }
        }
        try {
          token = localStorage.getItem('teamsChatAuthToken');
        } catch (e) {}
        if (!token) token = this.findTokenInStorage();
        if (token) {
          console.log('[Teams Chat API] ✅ Found token after waiting');
          break;
        }
      }
    }

    if (!token) {
      console.warn('[Teams Chat API] ❌ No auth token found after all attempts. API fetch will fail.');
      console.log('[Teams Chat API] Debug: Check if contextBridge.js and chatFetchOverride.js are loaded.');
      console.log('[Teams Chat API] Debug: Bridge element exists:', !!document.getElementById('teams-extractor-context-bridge'));
    } else {
      console.log('[Teams Chat API] Token found, length:', token.length);
    }

    if (token) {
      // Validate token audience before using
      if (!this.isValidTeamsToken(token)) {
        console.warn('[Teams Chat API] Ignoring invalid token (wrong audience):', token.substring(0, 20) + '...');
        token = null;
        try {
          window.localStorage.removeItem('teamsChatAuthToken');
        } catch (e) { }

        // Try to find a valid one again from storage
        const fresh = this.findTokenInStorage();
        if (fresh && this.isValidTeamsToken(fresh)) {
          console.log('[Teams Chat API] Found valid replacement token in storage');
          token = fresh;
        }
      }
    }

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
      // Ensure Bearer prefix if missing (MSAL secrets are just the token)
      if (!token.toLowerCase().startsWith('bearer ') && !token.toLowerCase().startsWith('skype_token')) {
        headers['Authorization'] = `Bearer ${token}`;
      } else {
        headers['Authorization'] = token;
      }
    }

    let currentUrl = baseUrl.toString();

    for (let i = 0; i < maxPages; i++) {
      try {
        console.log(`[Teams Chat API] Fetching page ${i + 1}...`);
        const resp = await fetch(currentUrl, {
          credentials: 'include',
          headers
        });

        if (resp.status === 401) {
          console.error('[Teams Chat API] 401 Unauthorized. Token was:', token ? 'Present' : 'Missing');
          try { window.localStorage.removeItem('teamsChatAuthToken'); } catch (e) { }
        }

        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        const data = await resp.json();
        const payloadMessages = this.extractFromPayload(data);
        payloadMessages.forEach((msg) => messages.push(msg));

        // syncState from Teams IS the next URL to fetch directly
        const nextUrl = data?.syncState || data?._metadata?.syncState || null;

        if (!nextUrl) {
          console.log(`[Teams Chat API] No more pages (no syncState returned)`);
          break;
        }

        // Use syncState as the next URL directly
        currentUrl = nextUrl;
      } catch (err) {
        console.warn('[Teams Chat API capture] fetch error', currentUrl, err);
        break;
      }
    }

    return this.sortAndDedup(messages);
  }

  findTokenInStorage() {
    try {
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        // Look for MSAL access tokens for Teams (IC3) or Skype
        if (key.includes('accesstoken') && (key.includes('ic3.teams.office.com') || key.includes('api.spaces.skype.com'))) {
          try {
            const val = JSON.parse(localStorage.getItem(key));
            if (val && val.secret) {
              console.log('[Teams Chat API] Found token in MSAL storage:', key);
              return val.secret;
            }
          } catch (e) { }
        }
      }
    } catch (e) {
      console.error('[Teams Chat API] Error scanning storage:', e);
    }
    return null;
  }

  isValidTeamsToken(token) {
    if (!token) return false;
    if (token.toLowerCase().startsWith('skype_token ')) return true;

    try {
      // Handle "Bearer " prefix if present
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
      // Robust JSON parse: find last '}' to ignore potential garbage
      const lastBrace = decoded.lastIndexOf('}');
      if (lastBrace === -1) return false;

      const jsonStr = decoded.substring(0, lastBrace + 1);
      const payload = JSON.parse(jsonStr);
      return payload.aud === 'https://ic3.teams.office.com';
    } catch (e) {
      return false;
    }
  }

  extractFromPayload(payload) {
    const buckets = [];

    if (Array.isArray(payload)) buckets.push(payload);
    if (Array.isArray(payload?.messages)) buckets.push(payload.messages);
    if (Array.isArray(payload?.value)) buckets.push(payload.value);
    if (Array.isArray(payload?.results)) buckets.push(payload.results);

    if (Array.isArray(payload?.conversations)) {
      payload.conversations.forEach((conv) => {
        if (Array.isArray(conv?.messages)) buckets.push(conv.messages);
        if (Array.isArray(conv?.chatMessages)) buckets.push(conv.chatMessages);
      });
    }

    const flattened = buckets.flat().filter(Boolean);
    return flattened.map((msg) => this.toStandardMessage(msg)).filter(Boolean);
  }

  toStandardMessage(msg) {
    if (!msg) return null;
    const id = msg.id || msg.messageId || msg.clientmessageid || msg.clientMessageId || msg.activityId;

    const author =
      msg.from?.user?.displayName ||
      msg.from?.user?.id ||
      msg.from?.displayName ||
      msg.from?.name ||
      msg.imDisplayName ||
      msg.sender ||
      msg.creator?.displayName ||
      'Unknown';

    const tsRaw =
      msg.createdDateTime ||
      msg.originalArrivalTime ||
      msg.composedTime ||
      msg.sentTime ||
      msg.created ||
      msg.createdTime ||
      msg.composetime ||
      msg.originalarrivaltime ||
      msg.timestamp;

    const isoTimestamp = this.parseIso(tsRaw);
    const timestamp = isoTimestamp ? new Date(isoTimestamp).toLocaleString() : (tsRaw || '');

    const body =
      msg.body?.content ??
      msg.bodyContent ??
      msg.content ??
      msg.message ??
      msg.plainText ??
      '';

    const text = this.stripHtml(typeof body === 'string' ? body : JSON.stringify(body));

    const attachments = this.extractAttachments(msg);

    const messageType = (msg.messageType || msg.type || '').toLowerCase();
    const derivedType = messageType.includes('system') ? 'system' : null;

    if (!text && attachments.length === 0) {
      return null;
    }

    return {
      id,
      author,
      timestamp,
      isoTimestamp,
      message: text,
      content: text,
      attachments,
      type: derivedType
    };
  }

  parseIso(value) {
    if (!value) return null;
    const parsed = Date.parse(value);
    if (Number.isNaN(parsed)) return null;
    return new Date(parsed).toISOString();
  }

  stripHtml(value) {
    if (!value) return '';
    const div = document.createElement('div');
    div.innerHTML = value;
    return div.textContent?.trim() || '';
  }

  extractAttachments(msg) {
    const out = [];

    if (Array.isArray(msg.attachments)) {
      msg.attachments.forEach((att) => {
        out.push({
          type: att?.contentType || att?.attachmentType || 'attachment',
          href: att?.contentUrl || att?.contentUri || att?.url,
          name: att?.name || att?.originalName || att?.displayName || att?.text || att?.title || ''
        });
      });
    }

    // Some Teams payloads include serialized files in properties.files as JSON string
    const files = msg.properties?.files || msg.files;
    if (typeof files === 'string') {
      try {
        const parsed = JSON.parse(files);
        if (Array.isArray(parsed)) {
          parsed.forEach((file) => {
            out.push({
              type: file?.type || file?.fileType || 'file',
              href: file?.objectUrl || file?.fileUrl || file?.url,
              name: file?.title || file?.fileName || file?.name || ''
            });
          });
        }
      } catch (_err) {
        // ignore parsing error
      }
    }

    return out;
  }

  sortAndDedup(messages) {
    const seen = new Map();
    messages.forEach((msg) => {
      if (!msg) return;
      const key = msg.id ? `id:${msg.id}` : `${msg.isoTimestamp || msg.timestamp}::${msg.author}::${msg.message}`;
      if (!seen.has(key)) {
        seen.set(key, msg);
      }
    });

    return Array.from(seen.values()).sort((a, b) => {
      if (a.isoTimestamp && b.isoTimestamp) {
        return new Date(a.isoTimestamp) - new Date(b.isoTimestamp);
      }
      if (a.isoTimestamp && !b.isoTimestamp) return -1;
      if (!a.isoTimestamp && b.isoTimestamp) return 1;
      return (a.timestamp || '').localeCompare(b.timestamp || '');
    });
  }
}
