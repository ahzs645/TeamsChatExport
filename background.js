// === SharePoint Token Capture via webRequest ===
// Captures Authorization headers from requests to *.sharepoint.com
// and stores them for use by the batch transcript API fetcher.
const spTokens = {}; // { hostname: { token, capturedAt } }

chrome.webRequest.onBeforeSendHeaders.addListener(
  (details) => {
    if (!details.requestHeaders) return;

    const authHeader = details.requestHeaders.find(
      (h) => h.name.toLowerCase() === 'authorization'
    );
    if (!authHeader || !authHeader.value) return;

    try {
      const url = new URL(details.url);
      const host = url.hostname;

      // Only capture for sharepoint.com hosts
      if (!host.includes('sharepoint.com') && !host.includes('sharepoint.us')) return;

      spTokens[host] = {
        token: authHeader.value,
        capturedAt: Date.now(),
        host
      };

      // Store in chrome.storage.local for content scripts to read
      chrome.storage.local.set({ sharePointTokens: spTokens });

      // Also forward to the active Teams tab immediately
      if (details.tabId > 0) {
        chrome.tabs.sendMessage(details.tabId, {
          action: 'sharePointTokenCaptured',
          host,
          token: authHeader.value,
          capturedAt: Date.now()
        }, () => {
          if (chrome.runtime.lastError) {
            // Tab may not have content script yet, ignore
          }
        });
      }
    } catch (e) {
      // Invalid URL, ignore
    }
  },
  {
    urls: ['*://*.sharepoint.com/*', '*://*.sharepoint.us/*'],
    types: ['xmlhttprequest', 'other']
  },
  ['requestHeaders']
);

// Store transcript API metadata from any frame (top frame or iframe)
const transcriptAPICache = {}; // { tabId: { metadata, tokens, transcriptData } }

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'storeTranscriptAPIMeta') {
    const tabId = sender.tab?.id;
    if (tabId) {
      if (!transcriptAPICache[tabId]) transcriptAPICache[tabId] = {};
      transcriptAPICache[tabId].metadata = {
        ...(transcriptAPICache[tabId].metadata || {}),
        ...request.metadata
      };
      if (request.tokens) {
        transcriptAPICache[tabId].tokens = {
          ...(transcriptAPICache[tabId].tokens || {}),
          ...request.tokens
        };
      }
      // Broadcast to all frames in this tab so the iframe content script can use it
      chrome.tabs.sendMessage(tabId, {
        action: 'transcriptAPIMetaUpdated',
        metadata: transcriptAPICache[tabId].metadata,
        tokens: transcriptAPICache[tabId].tokens || {}
      }, () => { if (chrome.runtime.lastError) { /* ignore */ } });
    }
    sendResponse({ ok: true });
    return true;
  }
  if (request.action === 'getTranscriptAPIMeta') {
    const tabId = sender.tab?.id;
    sendResponse(transcriptAPICache[tabId] || {});
    return true;
  }
});

// Allow content scripts to request current captured tokens
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'getSharePointTokens') {
    // Also refresh from storage in case another listener updated it
    chrome.storage.local.get(['sharePointTokens'], (result) => {
      const stored = result.sharePointTokens || {};
      // Merge with in-memory (in-memory is more recent)
      const merged = { ...stored, ...spTokens };
      sendResponse({ tokens: merged });
    });
    return true; // async
  }
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === "extract") {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0] && tabs[0].id) {
        chrome.tabs.sendMessage(tabs[0].id, {action: "extract", selectAll: request.selectAll}, (response) => {
          if (chrome.runtime.lastError) {
            // Silently ignore the error
          }
        });
      }
    });
  } else if (request.action === "download") {
    // Legacy entry point: persist the extraction and open the viewer on it.
    storeExtraction(request.data, (keys) => openResultsTab(keys));
  } else if (request.action === "openResults") {
    // Persist the latest extraction, then open the viewer with the stored key selected.
    // The viewer reads everything from storage; no displayData message is sent any more,
    // so an extraction appears exactly once in the sidebar.
    storeExtraction(request.data, (keys) => {
      openResultsTab(keys);
      if (typeof sendResponse === 'function') {
        sendResponse({ success: true, keys });
      }
    });
    return true;
  }
});

// savedExtractions:     { "[<locale time>] <chat name>": messages[] }  (key format read by older exports)
// savedExtractionsMeta: { <same key>: { name, source: 'extract', savedAt } }
// teamsChatData:        the latest extraction only (fallback for the viewer)
function storeExtraction(data, callback) {
  chrome.storage.local.get(['savedExtractions', 'savedExtractionsMeta'], (result) => {
    const savedExtractions = { ...(result.savedExtractions || {}) };
    const savedExtractionsMeta = { ...(result.savedExtractionsMeta || {}) };
    const now = new Date();
    const stamp = now.toLocaleString();
    const keys = [];

    Object.keys(data || {}).forEach((conversationName) => {
      let key = `[${stamp}] ${conversationName}`;
      for (let n = 2; key in savedExtractions; n += 1) {
        key = `[${stamp} #${n}] ${conversationName}`;
      }
      savedExtractions[key] = data[conversationName];
      savedExtractionsMeta[key] = { name: conversationName, source: 'extract', savedAt: now.toISOString() };
      keys.push(key);
    });

    chrome.storage.local.set({
      teamsChatData: data || {},
      savedExtractions,
      savedExtractionsMeta
    }, () => callback(keys));
  });
}

// Opens a new viewer tab; results.html?select=<key> selects the freshly stored conversation.
function openResultsTab(keys) {
  const url = chrome.runtime.getURL('results.html') +
    (keys && keys.length ? `?select=${encodeURIComponent(keys[0])}` : '');
  chrome.tabs.create({ url });
}
