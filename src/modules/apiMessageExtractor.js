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

  readNumberSetting(key, { min = 1, max = 500, fallback }, storageValue = null) {
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
      if (!raw) return fallback;
      const n = parseInt(raw, 10);
      if (Number.isNaN(n)) return fallback;
      return Math.min(Math.max(n, min), max);
    } catch (_err) {
      return fallback;
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
    const defaults = { host: 'teams.microsoft.com', region: 'amer' };
    try {
      const host = localStorage.getItem('teamsChatHostOverride') || defaults.host;
      // Try to detect region from recent API calls
      let region = localStorage.getItem('teamsChatRegionOverride');
      if (!region) {
        const perfEntries = performance?.getEntriesByType?.('resource') || [];
        for (let i = perfEntries.length - 1; i >= 0; i--) {
          const name = perfEntries[i].name || '';
          // Check for /api/mt/{region}/ pattern (v2)
          const mtMatch = name.match(/\/api\/mt\/([a-z]+)\//);
          if (mtMatch) {
            region = mtMatch[1];
            break;
          }
          // Check for /chatsvc/{region}/ pattern (v1)
          const chatsvcMatch = name.match(/\/chatsvc\/([a-z]+)\//);
          if (chatsvcMatch) {
            region = chatsvcMatch[1];
            break;
          }
        }
      }
      return { host, region: region || defaults.region };
    } catch (_err) {
      return defaults;
    }
  }

  isTeamsV2() {
    return window.location.href.includes('/v2/');
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
    const effectivePageSize = this.readNumberSetting('teamsChatApiPageSize', { min: 1, max: 500, fallback: pageSize }, storageSettings.teamsChatApiPageSize);
    const effectiveMaxPages = this.readNumberSetting('teamsChatApiMaxPages', { min: 1, max: 200, fallback: maxPages }, storageSettings.teamsChatApiMaxPages);
    const effectiveStartTime = this.readNumberSetting('teamsChatApiStartTime', { min: 1, max: Number.MAX_SAFE_INTEGER, fallback: startTime }, storageSettings.teamsChatApiStartTime);

    const { host, region } = this.getHostAndRegion();
    const encodedId = encodeURIComponent(conversationId);
    const isV2 = this.isTeamsV2();

    console.log(`[Teams Chat API] Detected Teams ${isV2 ? 'v2' : 'v1'}, region: ${region}`);

    // Build list of API endpoints to try
    const paths = [];

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
    const v1Base = `https://${host}/api/chatsvc/${region}/v1`;
    paths.push(
      `${v1Base}/users/ME/conversations/${encodedId}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=${effectivePageSize}&startTime=${effectiveStartTime}`,
      `${v1Base}/threads/${encodedId}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=${effectivePageSize}&startTime=${effectiveStartTime}`
    );

    for (const path of paths) {
      console.log(`[Teams Chat API] Trying: ${path.substring(0, 80)}...`);
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

        // Stop if we got 0 messages (no more content)
        if (payloadMessages.length === 0) {
          console.log(`[Teams Chat API] No more messages, stopping pagination`);
          break;
        }

        // Get next page URL - v1 uses syncState, v2 uses @odata.nextLink
        const nextUrl = data?.syncState ||
                       data?._metadata?.syncState ||
                       data?.['@odata.nextLink'] ||
                       data?.nextLink ||
                       null;

        if (!nextUrl) {
          console.log(`[Teams Chat API] No more pages`);
          break;
        }

        // Use next URL directly
        currentUrl = nextUrl;
      } catch (err) {
        // Don't log 404s as warnings - they're expected when trying multiple endpoints
        if (!err.message?.includes('404')) {
          console.warn('[Teams Chat API] fetch error', currentUrl, err);
        }
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
