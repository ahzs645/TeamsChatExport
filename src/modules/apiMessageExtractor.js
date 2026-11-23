/**
 * Api Message Extractor
 * Parses captured fetch responses (Teams chat APIs) into the standard message shape.
 * This is best-effort and gated via localStorage "teamsChatApiMode".
 */

const TEAMS_CLOUD_HOST = 'teams.cloud.microsoft';
const MAX_MESSAGES_PER_REQUEST = 200;

export class ApiMessageExtractor {
  constructor() {
    this.cacheKey = '__teamsExtractorMessageApiCache';
    this.hiddenDivId = 'teams-extractor-message-api-cache';
    this.conversationStartTs = 0;
  }

  readNumberSetting(key, { min = 1, max = 500, fallback }, storageValue = null) {
    const defaultValue = Math.min(Math.max(fallback, min), max);
    try {
      // Prefer chrome.storage.local value if provided
      if (storageValue !== null && storageValue !== undefined) {
        const n = parseInt(storageValue, 10);
        if (!Number.isNaN(n)) {
          return Math.min(Math.max(n, min), max);
        }
      }
      // Fallback to localStorage for backward compatibility
      const raw = localStorage.getItem(key);
      if (!raw) return defaultValue;
      const n = parseInt(raw, 10);
      if (Number.isNaN(n)) return defaultValue;
      return Math.min(Math.max(n, min), max);
    } catch (_err) {
      return defaultValue;
    }
  }

  async loadStorageSettings() {
    return new Promise((resolve) => {
      if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
        chrome.storage.local.get(['teamsChatApiPageSize', 'teamsChatApiMaxPages', 'teamsChatApiStartTime'], (result) => {
          resolve(result || {});
        });
      } else {
        resolve({});
      }
    });
  }

  getHostAndRegion() {
    const legacyHost = 'teams.microsoft.com';
    const host = window.location.hostname === TEAMS_CLOUD_HOST ? TEAMS_CLOUD_HOST : legacyHost;
    const defaults = { host, legacyHost, region: 'amer', chatRegion: null };
    try {
      const hostOverride = localStorage.getItem('teamsChatHostOverride');
      const regionOverride = localStorage.getItem('teamsChatRegionOverride');
      let region = regionOverride;
      let chatRegion = regionOverride;
      // Teams can route its web API and its chat history API to different regions.
      const entries = globalThis.performance?.getEntriesByType?.('resource') || [];
      const origin = `https://${hostOverride || host}/`;
      for (let i = entries.length - 1; i >= 0 && (!region || !chatRegion); i--) {
        const url = entries[i].name || '';
        if (url.startsWith(origin)) {
          const resourcePath = url.slice(origin.length);
          region ||= resourcePath.match(/^api\/mt\/([a-z0-9-]+)\//i)?.[1];
          chatRegion ||= resourcePath.match(/^api\/chatsvc\/([a-z0-9-]+)\//i)?.[1];
        }
      }
      return {
        host: hostOverride || host,
        legacyHost: hostOverride || legacyHost,
        region: region || chatRegion || defaults.region,
        chatRegion,
      };
    } catch (_err) {
      return defaults;
    }
  }

  isTeamsV2() {
    const { hostname, pathname } = window.location;
    return hostname === TEAMS_CLOUD_HOST || pathname.includes('/v2/');
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

    // Load settings from chrome.storage.local (set via popup), with localStorage fallback
    const storageSettings = await this.loadStorageSettings();

    // Allow overrides via chrome.storage.local or localStorage
    const effectivePageSize = this.readNumberSetting('teamsChatApiPageSize', { min: 1, max: MAX_MESSAGES_PER_REQUEST, fallback: pageSize }, storageSettings.teamsChatApiPageSize);
    const effectiveMaxPages = this.readNumberSetting('teamsChatApiMaxPages', { min: 1, max: 200, fallback: maxPages }, storageSettings.teamsChatApiMaxPages);
    const effectiveStartTime = this.readNumberSetting('teamsChatApiStartTime', { min: 1, max: Number.MAX_SAFE_INTEGER, fallback: startTime }, storageSettings.teamsChatApiStartTime);

    const { host, legacyHost, region, chatRegion } = this.getHostAndRegion();
    const encodedId = encodeURIComponent(conversationId);
    const isV2 = this.isTeamsV2();

    console.log(`[Teams Chat API] Detected Teams ${isV2 ? 'v2' : 'v1'}, region: ${region}`);

    // Build list of API endpoints to try
    const paths = [];
    const historyQuery = new URLSearchParams({
      view: 'msnp24Equivalent|supportsMessageProperties',
      pageSize: String(effectivePageSize),
      startTime: String(effectiveStartTime),
    });
    if (window.location.hostname === TEAMS_CLOUD_HOST) {
      // Cloud Teams serves history on the same origin. Preserve explicit and
      // observed chat regions, and try the Canadian route when only MT is visible.
      const chatRegions = new Set(chatRegion ? [chatRegion] : [region, 'ca']);
      for (const chatServiceRegion of chatRegions) {
        paths.push(`https://${host}/api/chatsvc/${chatServiceRegion}/v1/users/ME/conversations/${encodedId}/messages?${historyQuery}`);
      }
    }

    if (isV2) {
      // Teams v2 API endpoints
      const v2Base = `https://${host}/api/mt/${region}/beta`;
      paths.push(
        `${v2Base}/users/ME/chats/${encodedId}/messages?$top=${effectivePageSize}`,
        `${v2Base}/chats/${encodedId}/messages?$top=${effectivePageSize}`,
        `${v2Base}/me/chats/${encodedId}/messages?$top=${effectivePageSize}`
      );
    }

    // Always try v1 endpoints as fallback (they might still work)
    const v1Base = `https://${legacyHost}/api/chatsvc/${chatRegion || region}/v1`;
    paths.push(
      `${v1Base}/users/ME/conversations/${encodedId}/messages?${historyQuery}`,
      `${v1Base}/threads/${encodedId}/messages?${historyQuery}`,
    );

    for (const path of new Set(paths)) {
      const msgs = await this.fetchPaged(path, { maxPages: effectiveMaxPages });
      if (msgs.length > 0) {
        console.log(`[Teams Chat API] Success with endpoint, got ${msgs.length} messages`);
        return msgs;
      }
    }

    console.log('[Teams Chat API] All endpoints failed');
    return [];
  }

  async fetchPaged(initialUrl, { maxPages = 15 } = {}) {
    const messages = [];
    const baseUrl = new URL(initialUrl);
    const endpoint = {
      host: baseUrl.hostname,
      route: baseUrl.pathname.includes('/api/mt/') ? 'v2' : 'v1',
      region: baseUrl.pathname.split('/')[3],
      pageSize: Number(baseUrl.searchParams.get('pageSize') || baseUrl.searchParams.get('$top')) || undefined,
    };

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
        console.warn('[Teams Chat API] Ignoring token for a different audience');
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
    const visited = new Set([currentUrl]);

    for (let i = 0; i < maxPages; i++) {
      try {
        const resp = await fetch(currentUrl, {
          credentials: 'include',
          headers
        });

        if (resp.status === 401) {
          try { window.localStorage.removeItem('teamsChatAuthToken'); } catch (e) { }
        }

        const rejection = resp.ok ? undefined : await this.describeRejection(resp);
        console.info('[Teams Chat API] HTTP response', JSON.stringify({ ...endpoint, page: i + 1, status: resp.status, rejection }));
        if (!resp.ok) break;
        const data = await resp.json();
        const payloadMessages = this.extractFromPayload(data);
        payloadMessages.forEach((msg) => messages.push(msg));

        // Stop if we got 0 messages (no more content)
        if (payloadMessages.length === 0) {
          console.log(`[Teams Chat API] No more messages, stopping pagination`);
          break;
        }

        // Fetch older history when the chat service supplies a backward link.
        const continuation = data?._metadata?.backwardLink ||
                             data?.backwardLink ||
                             data?.syncState ||
                             data?._metadata?.syncState ||
                             data?.['@odata.nextLink'] ||
                             data?.nextLink;

        if (!continuation) {
          console.log(`[Teams Chat API] No more pages`);
          break;
        }

        const next = new URL(continuation, currentUrl);
        // Do not forward the Teams session to a different service or chat.
        if (next.origin !== baseUrl.origin || decodeURIComponent(next.pathname) !== decodeURIComponent(baseUrl.pathname) ||
            visited.has(next.href)) break;
        currentUrl = next.href;
        visited.add(currentUrl);
      } catch (err) {
        const reason = err instanceof SyntaxError ? 'Invalid JSON response' : 'Request or response processing failed';
        console.warn('[Teams Chat API]', reason, JSON.stringify({ ...endpoint, page: i + 1 }));
        break;
      }
    }

    return this.sortAndDedup(messages);
  }

  async describeRejection(response) {
    if (response.status === 401 || response.status === 403) return 'authorization';
    if (response.status !== 400 || !response.headers.get('content-type')?.includes('json')) return undefined;

    const payload = await response.json().catch(() => null);
    // Map known errors to labels. Microsoft can echo conversation IDs or other
    // request values in its error text, so never write the payload to the log.
    const details = [payload?.error?.code, payload?.error?.message, payload?.errorCode, payload?.message]
      .filter(value => typeof value === 'string').join(' ').toLowerCase();
    if (/invalid.?conversation|invalid.?thread|conversation.?not.?found|thread.?not.?found/.test(details)) return 'conversation-id';
    if (/invalid.?client|client.?version|unsupported.?version/.test(details)) return 'client-version';
    if (/access.?denied|forbidden|unauthori[sz]ed|invalid.?token|authentication/.test(details)) return 'authorization';
    if (/invalid.?parameter|pagesize|starttime|syncstate|query parameter/.test(details)) return 'history-parameters';
    return 'unclassified-bad-request';
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

    // v1 format
    if (Array.isArray(payload)) buckets.push(payload);
    if (Array.isArray(payload?.messages)) buckets.push(payload.messages);
    if (Array.isArray(payload?.results)) buckets.push(payload.results);

    // v2 format (OData style)
    if (Array.isArray(payload?.value)) buckets.push(payload.value);

    // Graph API / v2 nested format
    if (Array.isArray(payload?.conversations)) {
      payload.conversations.forEach((conv) => {
        if (Array.isArray(conv?.messages)) buckets.push(conv.messages);
        if (Array.isArray(conv?.chatMessages)) buckets.push(conv.chatMessages);
      });
    }

    // v2 chat messages might be directly in payload
    if (payload?.body?.content && payload?.from) {
      // Single message object
      buckets.push([payload]);
    }

    const flattened = buckets.flat().filter(Boolean);
    console.log(`[Teams Chat API] Extracted ${flattened.length} raw messages from payload`);
    return flattened.map((msg) => this.toStandardMessage(msg)).filter(Boolean);
  }

  toStandardMessage(msg) {
    if (!msg) return null;

    const id = msg.id || msg.messageId || msg.clientmessageid || msg.clientMessageId || msg.activityId;
    const messageType = (msg.messagetype || msg.messageType || msg.type || '').toLowerCase();

    // Filter out system/event messages that aren't useful
    const skipTypes = ['event/call', 'threadactivity', 'rlc:', 'event/', 'control/'];
    if (skipTypes.some(t => messageType.includes(t))) {
      return null;
    }

    // Check for call/meeting system messages by content pattern
    const bodyContent = msg.body?.content || msg.content || msg.message || '';
    if (this.isSystemCallMessage(bodyContent)) {
      return this.parseCallMessage(msg, bodyContent);
    }

    // Teams API uses lowercase property names
    const author =
      msg.imdisplayname ||
      msg.imDisplayName ||
      msg.displayName ||
      msg.displayname ||
      msg.from?.user?.displayName ||
      msg.from?.displayName ||
      msg.from?.name ||
      msg.properties?.displayName ||
      msg.properties?.displayname ||
      msg.sender ||
      msg.creator?.displayName ||
      msg.from?.user?.id ||
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

    const bodyStr = typeof body === 'string' ? body : JSON.stringify(body);
    const text = this.stripHtml(bodyStr);

    // Extract embedded images from HTML body
    const embeddedImages = this.extractEmbeddedImages(bodyStr);

    // Extract attachments
    const attachments = this.extractAttachments(msg);

    // Extract reactions
    const reactions = this.extractReactions(msg);

    // Derive type from already-defined messageType
    const derivedType = messageType.includes('system') ? 'system' : null;

    if (!text && attachments.length === 0 && embeddedImages.length === 0) {
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
      embeddedImages,
      reactions,
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

  extractEmbeddedImages(htmlBody) {
    const images = [];
    if (!htmlBody || typeof htmlBody !== 'string') return images;

    // Parse HTML to find img tags
    const div = document.createElement('div');
    div.innerHTML = htmlBody;

    const imgElements = div.querySelectorAll('img');
    imgElements.forEach((img) => {
      const src = img.getAttribute('src') || img.getAttribute('data-src');
      if (src) {
        images.push({
          src: src,
          alt: img.getAttribute('alt') || '',
          title: img.getAttribute('title') || '',
          itemtype: img.getAttribute('itemtype') || '',
          isEmoji: src.includes('emoji') || src.includes('sticker') || img.classList.contains('emoji')
        });
      }
    });

    // Also check for image URLs in the text (sometimes memes are just URLs)
    const urlRegex = /(https?:\/\/[^\s<>"]+\.(?:png|jpg|jpeg|gif|webp))/gi;
    const matches = htmlBody.match(urlRegex) || [];
    matches.forEach((url) => {
      // Avoid duplicates
      if (!images.some(img => img.src === url)) {
        images.push({
          src: url,
          alt: '',
          title: '',
          isEmoji: false
        });
      }
    });

    return images;
  }

  extractReactions(msg) {
    const reactions = [];

    // Teams stores reactions in properties.emotions or properties.reactions
    const emotions = msg.properties?.emotions || msg.emotions || msg.reactions;

    if (emotions) {
      // Emotions can be a JSON string or object
      let emotionsData = emotions;
      if (typeof emotions === 'string') {
        try {
          emotionsData = JSON.parse(emotions);
        } catch (_e) {
          return reactions;
        }
      }

      // Format: { "like": [{"user": "name", ...}], "heart": [...] }
      // Or: [{ "key": "like", "users": [...] }]
      if (Array.isArray(emotionsData)) {
        emotionsData.forEach((reaction) => {
          const key = reaction.key || reaction.type || reaction.emoji;
          const users = reaction.users || reaction.user || [];
          const userList = Array.isArray(users) ? users : [users];
          reactions.push({
            emoji: this.mapReactionKey(key),
            key: key,
            count: userList.length,
            users: userList.map(u => u.displayName || u.name || u.mri || u).filter(Boolean)
          });
        });
      } else if (typeof emotionsData === 'object') {
        Object.entries(emotionsData).forEach(([key, users]) => {
          const userList = Array.isArray(users) ? users : [users];
          reactions.push({
            emoji: this.mapReactionKey(key),
            key: key,
            count: userList.length,
            users: userList.map(u => u?.displayName || u?.name || u?.mri || (typeof u === 'string' ? u : '')).filter(Boolean)
          });
        });
      }
    }

    return reactions;
  }

  mapReactionKey(key) {
    const emojiMap = {
      'like': '👍',
      'heart': '❤️',
      'laugh': '😂',
      'surprised': '😮',
      'sad': '😢',
      'angry': '😠',
      'thumbsup': '👍',
      'thumbsdown': '👎',
      'clap': '👏',
      'fire': '🔥',
      'celebrate': '🎉',
      'thinking': '🤔'
    };
    return emojiMap[key?.toLowerCase()] || key || '👍';
  }

  isSystemCallMessage(content) {
    if (!content || typeof content !== 'string') return false;
    // Detect call/meeting system messages by common patterns
    const patterns = [
      /8:orgid:[a-f0-9-]+/i,  // User ID pattern
      /callStarted|callEnded/i,
      /flightproxy\.teams\.microsoft\.com/i,
      /RecurringException/i,
      /040000008200E00074C5B7101A82E008/i  // Calendar event ID pattern
    ];
    return patterns.some(p => p.test(content));
  }

  parseCallMessage(msg, content) {
    const tsRaw = msg.createdDateTime || msg.originalArrivalTime || msg.composetime || msg.timestamp;
    const isoTimestamp = this.parseIso(tsRaw);
    const timestamp = isoTimestamp ? new Date(isoTimestamp).toLocaleString() : (tsRaw || '');

    // Try to extract meeting title
    let meetingTitle = null;
    // Look for text between brackets like [Meeting Name]
    const bracketMatch = content.match(/\[([^\]]+)\]/);
    if (bracketMatch) {
      meetingTitle = bracketMatch[1];
    }

    // Determine if it's a call start or end
    let callAction = null;
    if (/callStarted/i.test(content)) {
      callAction = 'started';
    } else if (/callEnded/i.test(content)) {
      callAction = 'ended';
    }

    // Extract participant names from the raw data
    const participants = [];
    const nameMatches = content.matchAll(/8:orgid:[a-f0-9-]+([A-Z][a-z]+ [A-Z][a-z]+)/g);
    for (const match of nameMatches) {
      if (match[1] && !participants.includes(match[1])) {
        participants.push(match[1]);
      }
    }

    // Build a readable message
    let readableMessage = '';
    if (callAction === 'started') {
      readableMessage = '📞 Call started';
    } else if (callAction === 'ended') {
      readableMessage = '📞 Call ended';
    } else {
      readableMessage = '📅 Meeting event';
    }

    if (meetingTitle) {
      readableMessage += `: ${meetingTitle}`;
    }

    if (participants.length > 0) {
      readableMessage += ` (${participants.join(', ')})`;
    }

    return {
      id: msg.id || msg.messageId,
      author: 'System',
      timestamp,
      isoTimestamp,
      message: readableMessage,
      content: readableMessage,
      type: 'system',
      attachments: [],
      embeddedImages: [],
      reactions: []
    };
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
