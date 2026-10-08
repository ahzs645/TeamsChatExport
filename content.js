/**
 * Microsoft Teams Chat Extractor - Content Script
 *
 * Simplified version that uses API-based extraction only.
 * Injects scripts to capture auth tokens and conversation IDs,
 * then extracts chat messages via the Teams API.
 * Also supports transcript extraction from Teams recordings and Microsoft Stream.
 */

// === HELPER FUNCTIONS ===
const isVideoPage = () => {
  return window.location.href.includes('stream.aspx') ||
         window.location.href.includes('/recordings/') ||
         window.location.href.includes('streamContent') ||
         document.querySelector('video[src*="stream"]') !== null;
};

const isTeamsPage = () => {
  return window.location.href.includes('teams.microsoft.com') ||
         window.location.href.includes('teams.cloud.microsoft');
};

// === IMMEDIATE SCRIPT INJECTION ===
// Inject into page context immediately, before waiting for anything
(() => {
  const injectScript = (url, name) => {
    const script = document.createElement('script');
    script.src = chrome.runtime.getURL(url);
    script.onload = () => {
      console.log(`[Teams Chat Extractor] ${name} injected`);
      script.remove();
    };
    script.onerror = (e) => {
      console.error(`[Teams Chat Extractor] Failed to inject ${name}:`, e);
    };
    (document.head || document.documentElement).appendChild(script);
  };

  // Shared panel/toast helpers for page-world scripts (coordinator, manifestDownload).
  // The same file is also loaded in this isolated world via the manifest.
  if (!window.__tceUIInjected) {
    window.__tceUIInjected = true;
    injectScript('ui/tceUI.js', 'UI helpers');
  }

  // Inject fetch override first (captures tokens)
  if (!window.__teamsChatOverrideInjectedEarly) {
    window.__teamsChatOverrideInjectedEarly = true;
    injectScript('chatFetchOverride.js', 'Fetch override');
  }

  // Inject context bridge (exposes data to content script)
  if (!window.__teamsContextBridgeInjected) {
    window.__teamsContextBridgeInjected = true;
    injectScript('contextBridge.js', 'Context bridge');
  }

  // Inject API fetcher (Response.prototype.json interceptor) — must be early
  if (!window.__teamsAPIFetcherInjected) {
    window.__teamsAPIFetcherInjected = true;
    injectScript('transcriptAPIFetcher.js', 'Transcript API fetcher');
  }

  // Inject transcript fetch override for video/stream pages
  if (!window.__teamsTranscriptOverrideInjected) {
    window.__teamsTranscriptOverrideInjected = true;
    injectScript('transcriptFetchOverride.js', 'Transcript fetch override');
  }

  // Inject video download override for video/stream pages
  if (!window.__teamsVideoOverrideInjected) {
    window.__teamsVideoOverrideInjected = true;
    injectScript('videoDownloadOverride.js', 'Video download override');
    // Inject modular video download modules
    injectScript('videoDownload/directDownload.js', 'Video direct download');
    injectScript('videoDownload/manifestDownload.js', 'Video manifest download');
    injectScript('videoDownload/mseCaptureDownload.js', 'Video MSE capture download');
    injectScript('videoDownload/captureStreamDownload.js', 'Video capture stream');
    injectScript('videoDownload/fmp4ToMp4.js', 'fMP4 to MP4 converter');
    injectScript('videoDownload/coordinator.js', 'Video download coordinator');
  }

  // Inject batch transcript download
  if (!window.__teamsBatchTranscriptInjected) {
    window.__teamsBatchTranscriptInjected = true;
    injectScript('batchTranscriptDownload.js', 'Batch transcript download');
  }

  console.log('[Teams Chat Extractor] Content script starting...');
})();

(async () => {
  // Ensure body is ready when running at document_start
  if (!document.body) {
    await new Promise((resolve) => {
      if (document.readyState === 'complete' || document.readyState === 'interactive') {
        resolve();
        return;
      }
      document.addEventListener('DOMContentLoaded', resolve, { once: true });
    });
  }

  const [
    teamsModule,
    extractionModule
  ] = await Promise.all([
    import(chrome.runtime.getURL('src/modules/teamsVariantDetector.js')),
    import(chrome.runtime.getURL('src/modules/extractionEngine.js'))
  ]);

  const { TeamsVariantDetector } = teamsModule;
  const { ExtractionEngine } = extractionModule;

  // Shared injected UI (panels, toasts, menu). Normally loaded by the manifest
  // before this script; import it as a fallback (it is a classic IIFE that
  // assigns globalThis.__tceUI, which also works when imported as a module).
  if (!globalThis.__tceUI) {
    try { await import(chrome.runtime.getURL('ui/tceUI.js')); } catch (e) {
      console.error('[Teams Chat Extractor] Failed to load UI helpers:', e);
    }
  }
  const UI = globalThis.__tceUI;

  console.log('Teams Chat Extractor initialized');

  const extractionEngine = new ExtractionEngine();

  // --- Forward SharePoint tokens from background to page context ---
  const forwardTokenToPage = (host, token, capturedAt) => {
    document.dispatchEvent(new CustomEvent('teamsSharePointToken', {
      detail: { host, token, capturedAt }
    }));
  };

  // Load any previously captured tokens on startup
  chrome.runtime.sendMessage({ action: 'getSharePointTokens' }, (response) => {
    if (chrome.runtime.lastError || !response) return;
    const tokens = response.tokens || {};
    for (const [host, entry] of Object.entries(tokens)) {
      if (entry.token) {
        forwardTokenToPage(host, entry.token, entry.capturedAt);
      }
    }
  });

  // --- Cross-frame transcript API metadata ---
  // Watch for API metadata changes in the page's hidden div and forward to background
  let lastAPIMetaUpdate = 0;
  const forwardAPIMetaToBackground = () => {
    const apiDiv = document.getElementById('transcript-api-data');
    if (!apiDiv) return;
    const updated = parseInt(apiDiv.getAttribute('data-updated') || '0');
    if (updated <= lastAPIMetaUpdate) return;
    lastAPIMetaUpdate = updated;
    try {
      const metadata = JSON.parse(apiDiv.getAttribute('data-transcript-api') || '{}');
      const tokens = JSON.parse(apiDiv.getAttribute('data-sp-tokens') || '{}');
      if (Object.keys(metadata).length > 0 || Object.keys(tokens).length > 0) {
        chrome.runtime.sendMessage({
          action: 'storeTranscriptAPIMeta',
          metadata,
          tokens
        }, () => { if (chrome.runtime.lastError) { /* ignore */ } });
      }
    } catch (e) {}
  };

  // Poll for API metadata changes (the page script updates the hidden div)
  setInterval(forwardAPIMetaToBackground, 2000);

  // Store cross-frame metadata received from background
  let crossFrameAPIMeta = {};
  let crossFrameTokens = {};

  // === FRAME GATING ===
  // content.js runs in every frame (all_frames: true) and the popup messages
  // the whole tab (tabs.sendMessage without a frameId), so every frame gets
  // every message. Actions that create UI must run in exactly one frame:
  //   1. a frame that actually hosts the relevant element (player <video>,
  //      recap meeting picker, chat message list) handles it immediately and
  //      posts a claim to the top frame;
  //   2. otherwise the top frame handles it after a short grace period,
  //      unless a claim arrived meanwhile;
  //   3. every other frame ignores it (returns false so it doesn't hold the
  //      response channel open).
  // "Top frame only" would be wrong for Stream-in-Teams, where the player
  // lives in a SharePoint iframe; "every frame" produced duplicate panels.
  const IS_TOP_FRAME = (() => { try { return window.top === window; } catch (e) { return false; } })();
  const CLAIM_GRACE_MS = 400;
  const recentClaims = {};
  if (IS_TOP_FRAME) {
    window.addEventListener('message', (e) => {
      const d = e.data;
      if (d && d.__tceFrameClaim === true && typeof d.action === 'string') {
        recentClaims[d.action] = Date.now();
      }
    });
  }
  const announceClaim = (action) => {
    if (IS_TOP_FRAME) { recentClaims[action] = Date.now(); return; }
    try { window.top.postMessage({ __tceFrameClaim: true, action }, '*'); } catch (e) {}
  };
  // A real player, not a thumbnail/preview: rendered at a usable size.
  const hasPlayerVideo = () => Array.from(document.querySelectorAll('video'))
    .some((v) => v.offsetWidth * v.offsetHeight >= 160 * 90);
  const FRAME_RELEVANCE = {
    downloadVideo: () => hasPlayerVideo(),
    startBatchTranscript: () => !!document.querySelector(
      '[data-testid="intelligent-recap-instance-select-dropdown"], [aria-label="Transcript actions"], [aria-label="Audio recap"]'
    ),
    extractActiveChat: () => !!document.querySelector(
      '[data-tid="message-pane-list-container"], [data-tid="chat-pane-message"], #message-list'
    )
  };
  /**
   * Run `handler` in exactly one frame of the tab (rules above).
   * Returns true if this frame took (or may take) the action.
   */
  const runInOneFrame = (action, handler) => {
    let relevant = false;
    try { relevant = !!FRAME_RELEVANCE[action]?.(); } catch (e) {}
    if (relevant) {
      announceClaim(action);
      handler();
      return true;
    }
    if (!IS_TOP_FRAME) return false;
    const receivedAt = Date.now();
    setTimeout(() => {
      // A frame that claimed this action within ~1s of us receiving it wins.
      if ((recentClaims[action] || 0) > receivedAt - 1000) return;
      handler();
    }, CLAIM_GRACE_MS);
    return true;
  };

  // --- Message Listener ---
  chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
    switch (request.action) {
      case 'transcriptAPIMetaUpdated':
        // Received API metadata from another frame (via background.js)
        crossFrameAPIMeta = request.metadata || {};
        crossFrameTokens = request.tokens || {};
        console.log('[Teams Chat Extractor] Received cross-frame API metadata:',
          Object.keys(crossFrameAPIMeta).length, 'entries');
        sendResponse({ received: true });
        return true;
      case 'sharePointTokenCaptured':
        // Forward from background script to page context
        forwardTokenToPage(request.host, request.token, request.capturedAt);
        sendResponse({ received: true });
        return true;

      case 'getState':
        // Sub-frames (e.g. a Stream player iframe) only answer if they host a
        // chat, so their <h2> video title can't win the race against the top frame.
        if (!IS_TOP_FRAME && !FRAME_RELEVANCE.extractActiveChat()) return false;
        // Try multiple methods to get the chat name
        let currentChatName = null;

        // Method 1: h2 element (common in many Teams versions)
        const h2 = document.querySelector('h2');
        if (h2 && h2.textContent) {
          const name = h2.textContent.trim();
          // Allow commas for "Last, First" format names
          if (name.length > 0 && name.length < 100 && !name.includes('Microsoft')) {
            currentChatName = name;
          }
        }

        // Method 2: Teams v2 header selectors
        if (!currentChatName) {
          const headerSelectors = [
            '[data-tid="chat-header-title"]',
            '[data-tid="thread-header-title"]',
            '[data-tid="chat-title"]',
            '.fui-ThreadHeader__title',
            'main h1',
            '[role="main"] h1',
            '[role="heading"][aria-level="1"]'
          ];
          for (const selector of headerSelectors) {
            const el = document.querySelector(selector);
            if (el && el.textContent) {
              const name = el.textContent.trim();
              if (name.length > 0 && name.length < 100) {
                currentChatName = name;
                break;
              }
            }
          }
        }

        // Method 3: Selected sidebar item
        if (!currentChatName) {
          const selectedItem = document.querySelector('[aria-selected="true"]');
          if (selectedItem) {
            const firstLine = selectedItem.textContent?.split('\n')[0]?.trim();
            if (firstLine && firstLine.length > 0 && firstLine.length < 50) {
              currentChatName = firstLine;
            }
          }
        }

        // Method 4: TeamsVariantDetector fallback
        if (!currentChatName) {
          currentChatName = TeamsVariantDetector.getCurrentChatTitle();
        }

        sendResponse({
          currentChat: currentChatName || 'No chat open'
        });
        return true;

      case 'extractActiveChat':
        // Progress is shown in-page (toast) and broadcast to the extension as
        // { action: 'extractionProgress', status, count, message }.
        return runInOneFrame('extractActiveChat', () => {
          runChatExtraction();
          sendResponse({ status: 'extracting' });
        });

      case 'getCurrentChat':
        const currentChat = TeamsVariantDetector.getCurrentChatTitle();
        sendResponse({chatTitle: currentChat});
        return true;

      case 'startBatchTranscript':
        return runInOneFrame('startBatchTranscript', () => {
          setupBatchTranscriptPanel();
          sendResponse({ status: 'panel_opened' });
        });

      case 'downloadVideo':
        // Write command to a hidden div, coordinator (page world) picks it up
        return runInOneFrame('downloadVideo', () => {
          const method = request.data?.method || '';
          let cmdDiv = document.getElementById('tce-video-cmd');
          if (!cmdDiv) {
            cmdDiv = document.createElement('div');
            cmdDiv.id = 'tce-video-cmd';
            cmdDiv.style.display = 'none';
            document.body.appendChild(cmdDiv);
          }
          cmdDiv.setAttribute('data-command', 'download');
          cmdDiv.setAttribute('data-method', method);
          cmdDiv.setAttribute('data-timestamp', Date.now().toString());
          sendResponse({ status: 'download_started' });
        });

      case 'getVideoModules':
        // Check available modules by reading hidden divs from page world
        (() => {
          const hasCrypto = document.getElementById('video-crypto-data')?.getAttribute('data-ready') === '1';
          const hasDrive = !!document.getElementById('video-drive-data');
          const hasVideo = !!document.querySelector('video');
          // If crypto key was captured, the video played and segment URLs should be available
          const hasTemplates = hasCrypto;

          const modules = [
            {
              name: 'directDownload',
              label: 'Direct Download',
              description: 'Download original MP4 (needs download permission)',
              available: hasDrive
            },
            {
              name: 'manifestDownload',
              label: 'Fast Download',
              description: 'Parallel fetch + decrypt (play video briefly first)',
              available: hasTemplates
            },
            {
              name: 'mseCaptureDownload',
              label: 'Save MSE Capture',
              description: 'Save exact decrypted bytes the browser played (no re-fetch)',
              available: hasVideo
            },
            {
              name: 'captureStreamDownload',
              label: 'Record Stream',
              description: 'Record video at high speed (keep tab in foreground)',
              available: hasVideo
            }
          ];
          sendResponse({ modules });
        })();
        return true;

      case 'stopVideoDownload':
        // Inline <script> injection is blocked by the page CSP; use the same
        // hidden command div the coordinator already polls.
        (() => {
          let cmdDiv = document.getElementById('tce-video-cmd');
          if (!cmdDiv) {
            cmdDiv = document.createElement('div');
            cmdDiv.id = 'tce-video-cmd';
            cmdDiv.style.display = 'none';
            document.body.appendChild(cmdDiv);
          }
          cmdDiv.setAttribute('data-command', 'stop');
          cmdDiv.setAttribute('data-timestamp', Date.now().toString());
          sendResponse({ stopped: true });
        })();
        return true;

      case 'cancelBatchTranscript':
        sendBatchCommand('cancel');
        sendResponse({ status: 'cancelled' });
        return true;

      case 'getBatchTranscriptStatus':
        (async () => {
          const result = await sendBatchCommand('status');
          sendResponse(result || { running: false });
        })();
        return true;

      case 'getTranscriptStatus': {
        // `available` now means "extractTranscript would succeed": true if the
        // API path is usable (metadata + fresh token, local or cross-frame) or
        // DOM/captured transcript data exists. Extra fields: source
        // ('api' | 'captured' | 'dom' | null), apiAvailable, domAvailable.
        // Frames with data answer at once; frames with a player answer after a
        // short delay; the top frame answers last, so the most useful frame's
        // answer reaches the popup first.
        const hasVideo = !!document.querySelector('video');
        const onVideoPage = isVideoPage();
        const build = (r) => ({
          available: r.ready,
          hasVideo,
          isVideoPage: onVideoPage,
          source: r.source,
          apiAvailable: r.api,
          domAvailable: r.dom
        });
        const local = getLocalTranscriptReadiness();
        if (local.ready) {
          sendResponse(build(local));
          return false;
        }
        if (!hasVideo && !onVideoPage && !IS_TOP_FRAME) return false;
        getBackgroundTranscriptReadiness().then((bg) => {
          const r = bg.ready ? bg : local;
          if (r.ready) sendResponse(build(r));
          else setTimeout(() => sendResponse(build(r)), (hasVideo || onVideoPage) ? 150 : 400);
        });
        return true;
      }

      case 'extractTranscript': {
        // Extract transcript via API first, then DOM fallback. A frame that
        // finds nothing answers late, so a frame that has the transcript wins.
        const relevantFrame = IS_TOP_FRAME || isVideoPage() || !!document.querySelector('video') ||
          !!document.getElementById('teams-chat-exporter-transcript-data');
        if (!relevantFrame) return false;
        const failLater = (payload) => setTimeout(() => sendResponse(payload), IS_TOP_FRAME ? 1500 : 2500);
        (async () => {
          try {
            const data = await getTranscriptData();
            if (!data) {
              failLater({
                success: false,
                error: 'Transcript not ready. Start playback to load the transcript.'
              });
              return;
            }
            sendResponse({
              success: true,
              vtt: data.vtt,
              txt: data.txt,
              title: getVideoTitle(),
              source: data.source || 'unknown'
            });
          } catch (err) {
            failLater({
              success: false,
              error: 'Failed to extract transcript: ' + err.message
            });
          }
        })();
        return true;
      }

      default:
        sendResponse({status: 'unknown action'});
        return true;
    }
  });

  console.log('Teams Chat Extractor ready');

  // === NATIVE TOOLBAR INTEGRATION ===

  // Outline icons (20x20, stroke = currentColor) shared with the panels.
  const ICONS = UI.ICONS;

  // Title of the current recording, safe to use in a filename (keeps
  // non-ASCII characters; only strips characters illegal in filenames).
  const getVideoTitle = () => {
    const heading = document.querySelector('h1[class*="videoTitleViewModeHeading"] label');
    const h2 = document.querySelector('h2');
    const videoTitle = heading?.innerText?.trim() || h2?.innerText?.trim() || document.title?.trim() || '';
    return UI.sanitizeFilename(videoTitle, 'transcript');
  };

  // === API-based transcript fetch (same approach as batch downloader) ===

  const getAPIMetadata = () => {
    if (window.__transcriptAPIData && Object.keys(window.__transcriptAPIData).length > 0) {
      return { ...window.__transcriptAPIData };
    }
    const div = document.getElementById('transcript-api-data');
    if (!div) return {};
    try {
      return JSON.parse(div.getAttribute('data-transcript-api') || '{}');
    } catch (e) {
      return {};
    }
  };

  const getSharePointTokens = () => {
    if (window.__sharePointTokens && Object.keys(window.__sharePointTokens).length > 0) {
      return { ...window.__sharePointTokens };
    }
    const div = document.getElementById('transcript-api-data');
    if (!div) return {};
    try {
      return JSON.parse(div.getAttribute('data-sp-tokens') || '{}');
    } catch (e) {
      return {};
    }
  };

  const getTokenForUrl = (transcriptUrl) => {
    // Check local tokens, cross-frame tokens, and background tokens
    const localTokens = getSharePointTokens();
    const allTokens = { ...crossFrameTokens, ...localTokens };
    if (!transcriptUrl || Object.keys(allTokens).length === 0) return null;
    try {
      const host = new URL(transcriptUrl).hostname;
      const tokenEntry = allTokens[host];
      if (tokenEntry && tokenEntry.token) {
        if (Date.now() - tokenEntry.capturedAt < 30 * 60 * 1000) {
          return tokenEntry.token;
        }
      }
    } catch (e) {}
    return null;
  };

  const convertVttToTxt = (vttContent) => {
    const lines = vttContent.split('\n');
    const txtLines = [];
    let currentTimestamp = '';
    for (const line of lines) {
      if (line.includes('-->')) {
        currentTimestamp = line.split('-->')[0].trim().split('.')[0] || '00:00:00';
      } else if (line.trim() && !line.startsWith('WEBVTT') && !line.startsWith('NOTE') && !line.match(/^\d+$/)) {
        const vMatch = line.match(/^<v\s+([^>]+)>(.+)<\/v>$/);
        if (vMatch) {
          txtLines.push(`[${currentTimestamp}] ${vMatch[1]}: ${vMatch[2]}`);
        } else if (line.trim()) {
          txtLines.push(`[${currentTimestamp}] ${line.trim()}`);
        }
      }
    }
    return txtLines.join('\n') || vttContent;
  };

  const fetchTranscriptViaAPI = async (apiMeta) => {
    if (!apiMeta || !apiMeta.resources) return null;

    const transcriptResource =
      apiMeta.resources['TranscriptV2'] ||
      apiMeta.resources['transcriptV2'] ||
      Object.values(apiMeta.resources).find((r) =>
        r.location && r.location.includes('transcript')
      );

    if (!transcriptResource || !transcriptResource.location) return null;

    const transcriptUrl = transcriptResource.location;
    const token = getTokenForUrl(transcriptUrl);
    if (!token) return null;

    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 15000);

      const response = await fetch(transcriptUrl, {
        headers: { 'Authorization': token },
        signal: controller.signal
      });
      clearTimeout(timeout);

      if (!response.ok) return null;

      const contentType = response.headers.get('content-type') || '';
      const text = await response.text();
      if (!text || text.length < 10) return null;

      let vtt = '';
      let txt = '';

      if (text.startsWith('WEBVTT')) {
        vtt = text;
        txt = convertVttToTxt(text);
      } else if (contentType.includes('json') || text.startsWith('{') || text.startsWith('[')) {
        try {
          const data = JSON.parse(text);
          const entries = Array.isArray(data) ? data : (data.entries || data.captions || []);
          if (entries.length > 0) {
            const utils = window.__teamsTranscriptUtils;
            if (utils) {
              vtt = utils.buildVttTranscript(entries);
              txt = utils.buildTxtTranscript(entries);
            } else {
              txt = entries.map((e) => {
                const ts = e.startOffset || e.offset || '00:00:00';
                const speaker = e.speakerDisplayName || e.speaker || 'Unknown';
                const content = e.text || e.content || '';
                return `[${ts}] ${speaker}: ${content}`;
              }).join('\n');
              vtt = 'WEBVTT\n\n' + entries.map((e, i) => {
                const ts = e.startOffset || e.offset || '00:00:00';
                const speaker = e.speakerDisplayName || e.speaker || 'Unknown';
                const content = e.text || e.content || '';
                return `${i + 1}\n${ts}.000 --> ${ts}.000\n<v ${speaker}>${content}</v>`;
              }).join('\n\n');
            }
          }
        } catch (e) {
          return null;
        }
      } else {
        txt = text;
        vtt = 'WEBVTT\n\nNOTE Raw transcript from API\n\n1\n00:00:00.000 --> 99:59:59.000\n' + text;
      }

      if (!vtt && !txt) return null;
      return { vtt, txt, source: 'api' };
    } catch (err) {
      console.error('[Teams Chat Extractor] API fetch failed:', err);
      return null;
    }
  };

  // Get captured transcript content from cdnmedia/transcripts intercept
  const getCapturedTranscriptContent = () => {
    if (window.__capturedTranscriptContent && Object.keys(window.__capturedTranscriptContent).length > 0) {
      return { ...window.__capturedTranscriptContent };
    }
    const div = document.getElementById('captured-transcript-content');
    if (!div) return {};
    try {
      return JSON.parse(div.getAttribute('data-transcripts') || '{}');
    } catch (e) {
      return {};
    }
  };

  // Priority: 1) API metadata fetch (local or cross-frame), 2) captured cdnmedia, 3) React state, 4) DOM
  const getTranscriptData = async () => {
    const utils = window.__teamsTranscriptUtils;

    // 1a. Try API approach via local readcollabobject metadata
    const allMeta = getAPIMetadata();
    let metaKeys = Object.keys(allMeta);
    if (metaKeys.length > 0) {
      const latestKey = metaKeys[metaKeys.length - 1];
      const apiResult = await fetchTranscriptViaAPI(allMeta[latestKey]);
      if (apiResult) {
        console.log('[Teams Chat Extractor] Using API-fetched transcript (complete, local)');
        return apiResult;
      }
    }

    // 1b. Try cross-frame API metadata (from another frame in same tab, via background.js)
    if (Object.keys(crossFrameAPIMeta).length > 0) {
      metaKeys = Object.keys(crossFrameAPIMeta);
      const latestKey = metaKeys[metaKeys.length - 1];
      const apiResult = await fetchTranscriptViaAPI(crossFrameAPIMeta[latestKey]);
      if (apiResult) {
        console.log('[Teams Chat Extractor] Using API-fetched transcript (complete, cross-frame)');
        return apiResult;
      }
    }

    // 1c. Ask background.js for any cached metadata (fallback)
    try {
      const bgMeta = await new Promise((resolve) => {
        chrome.runtime.sendMessage({ action: 'getTranscriptAPIMeta' }, (resp) => {
          if (chrome.runtime.lastError) resolve({});
          else resolve(resp || {});
        });
      });
      if (bgMeta.metadata) {
        metaKeys = Object.keys(bgMeta.metadata);
        if (metaKeys.length > 0) {
          const latestKey = metaKeys[metaKeys.length - 1];
          const apiResult = await fetchTranscriptViaAPI(bgMeta.metadata[latestKey]);
          if (apiResult) {
            console.log('[Teams Chat Extractor] Using API-fetched transcript (complete, from background cache)');
            return apiResult;
          }
        }
      }
    } catch (e) {}

    // 2. Try captured cdnmedia/transcripts response (intercepted at page load)
    const captured = getCapturedTranscriptContent();
    const capturedKeys = Object.keys(captured);
    if (capturedKeys.length > 0) {
      const latest = captured[capturedKeys[capturedKeys.length - 1]];

      // Case A: raw VTT captured directly
      if (latest && latest.rawVtt) {
        console.log(`[Teams Chat Extractor] Using captured cdnmedia VTT (${latest.entryCount} entries, complete)`);
        return {
          vtt: latest.rawVtt,
          txt: convertVttToTxt(latest.rawVtt),
          source: 'cdnmedia-vtt'
        };
      }

      // Case B: JSON entries captured
      if (latest && latest.entries && latest.entries.length > 0 && utils) {
        const entries = latest.entries.map((e, i) => ({
          startOffset: e.startOffset || e.offset || (utils.parsePTDuration ? utils.parsePTDuration(e.timestamp) : '00:00:00'),
          endOffset: e.endOffset || e.startOffset || '00:00:00',
          speakerDisplayName: e.speakerDisplayName || e.speaker || 'Unknown',
          text: e.text || e.content || '',
          id: e.entryId || `captured-${i + 1}`
        })).filter(e => e.text.trim());

        if (entries.length > 0) {
          console.log(`[Teams Chat Extractor] Using captured cdnmedia transcript (${entries.length} entries, complete)`);
          return {
            vtt: utils.buildVttTranscript(entries),
            txt: utils.buildTxtTranscript(entries),
            source: 'cdnmedia'
          };
        }
      }
    }

    // 3. Try React state extraction - gets all entries from virtualized list
    if (utils && utils.extractFromReactState) {
      const reactEntries = utils.extractFromReactState();
      if (reactEntries && reactEntries.length > 0) {
        console.log(`[Teams Chat Extractor] Using React state transcript (${reactEntries.length} entries, complete)`);
        return {
          vtt: utils.buildVttTranscript(reactEntries),
          txt: utils.buildTxtTranscript(reactEntries),
          source: 'react'
        };
      }
    }

    // 4. Fall back to DOM hidden div (populated by transcriptFetchOverride.js which uses React state)
    const container = document.getElementById('teams-chat-exporter-transcript-data');
    if (!container) return null;
    const entryCount = container.getAttribute('data-count') || '?';
    const dataSource = container.getAttribute('data-source') || 'unknown';
    console.log(`[Teams Chat Extractor] Using transcript from hidden div (${entryCount} entries, source: ${dataSource})`);
    return {
      vtt: container.getAttribute('data-vtt') || container.textContent,
      txt: container.getAttribute('data-txt') || container.textContent,
      source: dataSource
    };
  };

  // === TRANSCRIPT READINESS ===
  // Mirrors what getTranscriptData() can actually use, so the popup's
  // "Transcript ready" matches what extraction will do.

  const latestEntry = (map) => {
    const keys = Object.keys(map || {});
    return keys.length ? map[keys[keys.length - 1]] : null;
  };

  // Same resource selection as fetchTranscriptViaAPI.
  const apiReadyFor = (apiMeta) => {
    if (!apiMeta || !apiMeta.resources) return false;
    const res = apiMeta.resources['TranscriptV2'] ||
      apiMeta.resources['transcriptV2'] ||
      Object.values(apiMeta.resources).find((r) => r.location && r.location.includes('transcript'));
    return !!(res && res.location && getTokenForUrl(res.location));
  };

  const hasDomTranscript = () => {
    const container = document.getElementById('teams-chat-exporter-transcript-data');
    if (container && (container.textContent || container.getAttribute('data-vtt'))) return true;
    const captured = latestEntry(getCapturedTranscriptContent());
    return !!(captured && captured.rawVtt);
  };

  const getLocalTranscriptReadiness = () => {
    const api = apiReadyFor(latestEntry(getAPIMetadata())) || apiReadyFor(latestEntry(crossFrameAPIMeta));
    const dom = hasDomTranscript();
    return { ready: api || dom, api, dom, source: api ? 'api' : (dom ? 'dom' : null) };
  };

  const getBackgroundTranscriptReadiness = () => new Promise((resolve) => {
    const none = { ready: false, api: false, dom: false, source: null };
    try {
      chrome.runtime.sendMessage({ action: 'getTranscriptAPIMeta' }, (resp) => {
        if (chrome.runtime.lastError || !resp || !resp.metadata) { resolve(none); return; }
        const api = apiReadyFor(latestEntry(resp.metadata));
        resolve(api ? { ready: true, api: true, dom: false, source: 'api' } : none);
      });
    } catch (e) {
      resolve(none);
    }
  });

  // === SAVING / FEEDBACK HELPERS ===

  const saveBlob = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  };

  const downloadFile = (content, filename) => {
    saveBlob(new Blob([content], { type: 'text/plain;charset=utf-8' }), filename);
  };

  const NOT_READY_MESSAGE = 'Transcript not ready yet. Start playback (or open the Transcript tab) so it loads, then try again.';

  // Fetch transcript data; show a progress toast if it takes a moment and an
  // error toast if nothing is available.
  const getTranscriptWithFeedback = async () => {
    let busy = null;
    const timer = setTimeout(() => {
      busy = UI.toast({ kind: 'progress', message: 'Fetching transcript…' });
    }, 400);
    try {
      const data = await getTranscriptData();
      if (!data) UI.toast({ kind: 'error', title: 'Transcript not available', message: NOT_READY_MESSAGE });
      return data;
    } catch (err) {
      UI.toast({ kind: 'error', title: 'Transcript not available', message: err.message || String(err) });
      return null;
    } finally {
      clearTimeout(timer);
      if (busy) busy.close();
    }
  };

  // Brief "copied" state on a toolbar/command button (icon swaps to a check).
  const flashCopied = (btn) => {
    if (!btn || btn.classList.contains('tce-copied')) return;
    const origHTML = btn.innerHTML;
    btn.classList.add('tce-copied');
    const svg = btn.querySelector('svg');
    if (svg) svg.outerHTML = ICONS.check;
    const label = btn.querySelector('.tce-label-text');
    if (label) label.textContent = 'Copied';
    setTimeout(() => {
      btn.classList.remove('tce-copied');
      btn.innerHTML = origHTML;
    }, 1500);
  };

  // Copy transcript handler
  const handleCopy = async (btn) => {
    const data = await getTranscriptWithFeedback();
    if (!data) return;
    const ok = await UI.copyText(data.vtt);
    if (ok) {
      flashCopied(btn);
      UI.toast({ kind: 'success', message: 'Transcript copied to clipboard.' });
    } else {
      console.error('[Teams Chat Extractor] Failed to copy transcript to clipboard');
      UI.toast({
        kind: 'error',
        title: 'Couldn’t copy transcript',
        message: 'The browser blocked clipboard access. Click on the page so it has focus, then try again, or use Download instead.'
      });
    }
  };

  // Download VTT / TXT handlers
  const handleDownload = async (format) => {
    const data = await getTranscriptWithFeedback();
    if (!data) return;
    const filename = `transcript-${getVideoTitle()}.${format}`;
    downloadFile(format === 'txt' ? data.txt : data.vtt, filename);
    UI.toast({ kind: 'success', message: `Saved ${filename}` });
  };
  const handleDownloadVTT = () => handleDownload('vtt');
  const handleDownloadTXT = () => handleDownload('txt');

  const downloadMenuItems = () => ([
    { label: 'Download VTT', icon: 'download', onSelect: handleDownloadVTT },
    { label: 'Download TXT', icon: 'download', onSelect: handleDownloadTXT }
  ]);

  // Icon-only button for Teams' recap toolbar: inherits the toolbar's colour.
  const makeToolbarBtn = (icon, label, handler) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'tce-toolbar-btn';
    btn.title = label;
    btn.setAttribute('aria-label', label);
    btn.innerHTML = icon;
    if (handler) btn.addEventListener('click', handler);
    return btn;
  };

  // Labelled command button (transcript actions menubar, recap bar).
  const makeCmdBtn = (label, icon, handler, ariaLabel) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'tce-cmd-btn';
    btn.innerHTML = icon;
    const span = document.createElement('span');
    span.className = 'tce-label-text';
    span.textContent = label;
    btn.appendChild(span);
    if (ariaLabel) btn.setAttribute('aria-label', ariaLabel);
    if (handler) btn.addEventListener('click', handler);
    return btn;
  };

  // === LOCATION 1: Recap Action Toolbar ===
  const injectRecapToolbarButtons = () => {
    // Find toolbar by anchor button
    const anchorBtn = document.querySelector('[aria-label="Audio recap"]');
    const toolbar = anchorBtn?.closest('.fui-Toolbar');
    if (!toolbar || toolbar.hasAttribute('data-tce-injected')) return;
    toolbar.setAttribute('data-tce-injected', 'true');

    const copyBtn = makeToolbarBtn(ICONS.copy, 'Copy transcript', (e) => handleCopy(e.currentTarget));
    const downloadBtn = makeToolbarBtn(ICONS.download + ICONS.chevron, 'Download transcript');
    UI.bindMenuButton(downloadBtn, downloadMenuItems, { label: 'Download transcript' });
    const batchBtn = makeToolbarBtn(ICONS.batch, 'Batch download all transcripts', () => setupBatchTranscriptPanel());
    const videoBtn = makeToolbarBtn(ICONS.video, 'Download video', () => handleDirectVideoDownload());

    toolbar.appendChild(copyBtn);
    toolbar.appendChild(downloadBtn);
    toolbar.appendChild(batchBtn);
    toolbar.appendChild(videoBtn);

    console.log('[Teams Chat Extractor] Injected buttons into recap action toolbar');
  };

  // === LOCATION 2: Transcript Actions Menubar ===
  const injectTranscriptActionButtons = () => {
    const menubar = document.querySelector('[aria-label="Transcript actions"]');
    if (!menubar || menubar.hasAttribute('data-tce-injected')) return;
    menubar.setAttribute('data-tce-injected', 'true');

    // Add a separator first
    const sep = document.createElement('span');
    sep.className = 'tce-cmd-separator';
    sep.setAttribute('role', 'separator');
    menubar.appendChild(sep);

    menubar.appendChild(makeCmdBtn('Copy', ICONS.copy, (e) => handleCopy(e.currentTarget), 'Copy transcript'));
    menubar.appendChild(makeCmdBtn('Download VTT', ICONS.download, handleDownloadVTT));
    menubar.appendChild(makeCmdBtn('Download TXT', ICONS.download, handleDownloadTXT));
    menubar.appendChild(makeCmdBtn('Batch Download', ICONS.batch, () => setupBatchTranscriptPanel(), 'Batch download all transcripts'));

    console.log('[Teams Chat Extractor] Injected buttons into transcript actions menubar');
  };

  // === LOCATION 3: Teams Recap Page (teams.cloud.microsoft) ===
  const injectTeamsRecapButtons = () => {
    // Only on teams.cloud.microsoft
    if (!window.location.hostname.includes('teams.cloud.microsoft')) return;

    // Don't inject twice
    if (document.querySelector('.tce-recap-bar')) return;

    // Check if we have API metadata (means we're on a recap page)
    // Note: __transcriptAPIData is in the page world, not content script world.
    // Read from the hidden DOM element instead.
    const apiDiv = document.getElementById('transcript-api-data');
    let hasAPIMeta = false;
    if (apiDiv) {
      try {
        const meta = JSON.parse(apiDiv.getAttribute('data-transcript-api') || '{}');
        hasAPIMeta = Object.keys(meta).length > 0;
      } catch (e) {}
    }
    if (!hasAPIMeta) return;

    // Find the Notes/AI summary/Transcript tablist
    const tablists = document.querySelectorAll('[role="tablist"]');
    let targetTablist = null;
    for (const tl of tablists) {
      if (tl.textContent?.includes('Transcript')) {
        targetTablist = tl;
        break;
      }
    }
    if (!targetTablist) return;
    const recapPanel = targetTablist.parentElement;
    if (!recapPanel) return;

    // Compact toolbar using the same command-button style as the other locations
    const bar = document.createElement('div');
    bar.className = 'tce-recap-bar';
    bar.setAttribute('role', 'toolbar');
    bar.setAttribute('aria-label', 'Teams Chat Exporter');

    bar.appendChild(makeCmdBtn('Copy', ICONS.copy, (e) => handleCopy(e.currentTarget), 'Copy transcript'));
    bar.appendChild(makeCmdBtn('VTT', ICONS.download, handleDownloadVTT, 'Download transcript as VTT'));
    bar.appendChild(makeCmdBtn('TXT', ICONS.download, handleDownloadTXT, 'Download transcript as TXT'));
    bar.appendChild(makeCmdBtn('Video', ICONS.video, () => handleDirectVideoDownload(), 'Download video'));

    // Insert after the tablist's parent container (below the tabs, inside the scrollable area)
    targetTablist.parentElement.insertAdjacentElement('afterend', bar);
    console.log('[Teams Chat Extractor] Injected buttons into Teams recap page');
  };

  const tryInjectToolbarButtons = () => {
    injectRecapToolbarButtons();
    injectTranscriptActionButtons();
    injectTeamsRecapButtons();
  };

  // === CHAT EXTRACTION (popup "Extract This Chat") ===

  const broadcastExtraction = (detail) => {
    try {
      chrome.runtime.sendMessage({ action: 'extractionProgress', ...detail }, () => {
        if (chrome.runtime.lastError) { /* popup closed: ignore */ }
      });
    } catch (e) {}
  };

  let chatExtractionRunning = false;

  /**
   * Runs extractionEngine.extractActiveChat() with in-page feedback. Progress
   * comes from wrapping the engine's per-page / per-stage methods on this
   * instance only (the modules themselves are unchanged).
   */
  const runChatExtraction = async () => {
    if (chatExtractionRunning) return;
    chatExtractionRunning = true;

    const progress = UI.toast({ kind: 'progress', title: 'Extracting chat', message: 'Looking for messages…' });
    broadcastExtraction({ status: 'started', count: 0, message: 'Looking for messages…' });

    let count = 0;
    const report = (message) => {
      progress.update({ message });
      broadcastExtraction({ status: 'progress', count, message });
    };

    const restores = [];
    const wrap = (obj, name, makeWrapper) => {
      if (!obj || typeof obj[name] !== 'function') return;
      const own = Object.prototype.hasOwnProperty.call(obj, name);
      const orig = obj[name];
      obj[name] = makeWrapper(orig);
      restores.push(() => { if (own) obj[name] = orig; else delete obj[name]; });
    };

    const api = extractionEngine.apiMessageExtractor;
    wrap(api, 'extractFromPayload', (orig) => function (...args) {
      const result = orig.apply(this, args);
      if (Array.isArray(result) && result.length) {
        count += result.length;
        report(`Extracting chat… ${count.toLocaleString()} messages`);
      }
      return result;
    });
    wrap(extractionEngine, 'scrollToLoadMessages', (orig) => function (...args) {
      report('Scrolling to load older messages…');
      return orig.apply(this, args);
    });
    wrap(extractionEngine, 'extractMessagesFromDOM', (orig) => function (...args) {
      const result = orig.apply(this, args);
      if (Array.isArray(result)) {
        count = result.length;
        report(`Extracting chat… ${count.toLocaleString()} messages`);
      }
      return result;
    });
    wrap(extractionEngine, 'embedAvatars', (orig) => function (messages, ...rest) {
      report(`Preparing ${(messages?.length || count).toLocaleString()} messages…`);
      return orig.call(this, messages, ...rest);
    });

    try {
      const result = await extractionEngine.extractActiveChat();
      restores.forEach((fn) => fn());
      restores.length = 0;
      if (!result) {
        progress.update({
          kind: 'warning',
          title: 'No messages found',
          message: 'Open a chat and let it load, then try again.',
          timeout: 0
        });
        broadcastExtraction({ status: 'empty', count: 0, message: 'No messages found' });
        return;
      }
      const total = Object.values(result).reduce((n, msgs) => n + (Array.isArray(msgs) ? msgs.length : 0), 0);
      const response = await chrome.runtime.sendMessage({ action: 'openResults', data: result });
      if (response && response.success === false) {
        throw new Error(response.error || 'The viewer could not be opened.');
      }
      progress.update({
        kind: 'success',
        title: 'Chat extracted',
        message: `Opened ${total.toLocaleString()} messages in the viewer.`
      });
      broadcastExtraction({ status: 'done', count: total, message: 'Opened in viewer' });
    } catch (error) {
      console.error('Error extracting active chat:', error);
      progress.update({
        kind: 'error',
        title: 'Chat extraction failed',
        message: error?.message || String(error)
      });
      broadcastExtraction({ status: 'error', count, message: error?.message || String(error) });
    } finally {
      restores.forEach((fn) => fn());
      chatExtractionRunning = false;
    }
  };

  // === VIDEO DOWNLOAD PANEL ===
  let videoCaptureAvailable = false;

  // Helper to send commands to the injected video capture script
  const sendVideoCommand = (command, data = {}, timeoutMs = 2000) => {
    // Use longer timeout for download operations
    if (command === 'downloadFiles' || command === 'directDownload' || command === 'downloadCombined') {
      timeoutMs = 300000; // 5 minutes for downloads
    }

    return new Promise((resolve) => {
      const handler = (e) => {
        if (e.detail?.command === command) {
          document.removeEventListener('teamsVideoResponse', handler);
          resolve(e.detail.result);
        }
      };
      document.addEventListener('teamsVideoResponse', handler);
      document.dispatchEvent(new CustomEvent('teamsVideoCommand', {
        detail: { command, data }
      }));
      // Timeout
      setTimeout(() => {
        document.removeEventListener('teamsVideoResponse', handler);
        resolve(null);
      }, timeoutMs);
    });
  };

  // Helper to send commands to the injected batch transcript script
  const sendBatchCommand = (command, data = {}, timeoutMs = 600000) => {
    return new Promise((resolve) => {
      const handler = (e) => {
        if (e.detail?.command === command) {
          document.removeEventListener('teamsBatchTranscriptResponse', handler);
          resolve(e.detail.result);
        }
      };
      document.addEventListener('teamsBatchTranscriptResponse', handler);
      document.dispatchEvent(new CustomEvent('teamsBatchTranscriptCommand', {
        detail: { command, data }
      }));
      setTimeout(() => {
        document.removeEventListener('teamsBatchTranscriptResponse', handler);
        resolve(null);
      }, timeoutMs);
    });
  };

  // Check if video capture is available
  const checkVideoCaptureAvailable = async () => {
    const result = await sendVideoCommand('ping');
    videoCaptureAvailable = result?.available === true;
    return videoCaptureAvailable;
  };

  // Listen for video capture ready event
  document.addEventListener('teamsVideoReady', () => {
    videoCaptureAvailable = true;
    console.log('[Teams Chat Extractor] Video capture ready');
  });

  // Developer aid (formerly a button in the panel): logs captured segment URL
  // patterns. Run from the page console:
  //   document.dispatchEvent(new CustomEvent('teamsVideoCommand', {detail: {command: 'analyzeUrls'}}))
  // or from the extension's content-script console: __tceAnalyzeVideoUrls()
  globalThis.__tceAnalyzeVideoUrls = async () => {
    const result = await sendVideoCommand('analyzeUrls');
    console.log('[URL Analysis]', result);
    return result;
  };

  // === DIRECT VIDEO DOWNLOAD (preferred, uses SharePoint API) ===
  const getVideoDriveItem = () => {
    // Read from hidden div (populated by transcriptAPIFetcher.js in page world)
    const div = document.getElementById('video-drive-data');
    if (!div) return null;
    try {
      const data = JSON.parse(div.getAttribute('data-drive-item') || '{}');
      if (data.driveId && data.itemId) return data;
    } catch (e) {}
    return null;
  };

  const DIRECT_DOWNLOAD_TOAST = 'The original file is downloading in a new tab.';

  const handleDirectVideoDownload = async () => {
    // Tier 1: Use pre-authenticated download URL if already captured by page-world script
    const driveItem = getVideoDriveItem();
    if (driveItem && driveItem.downloadUrl) {
      const filename = driveItem.fileName || 'recording.mp4';
      const sizeMB = driveItem.fileSize ? Math.round(driveItem.fileSize / 1024 / 1024) : '?';
      console.log(`[Teams Chat Extractor] Direct video download: ${filename} (${sizeMB}MB)`);
      window.open(driveItem.downloadUrl, '_blank');
      UI.toast({ kind: 'success', title: 'Video download started', message: DIRECT_DOWNLOAD_TOAST });
      return;
    }

    // Tier 2: Use download.aspx with the file path from URL (cookie-based, most reliable)
    const urlParams = new URLSearchParams(window.location.search);
    const filePath = urlParams.get('id');
    if (filePath) {
      const siteMatch = window.location.pathname.match(/^(\/personal\/[^/]+|\/sites\/[^/]+)/);
      const siteBase = siteMatch ? siteMatch[0] : '';
      const downloadUrl = `https://${window.location.hostname}${siteBase}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(filePath)}`;
      console.log('[Teams Chat Extractor] Video download via download.aspx');
      window.open(downloadUrl, '_blank');
      UI.toast({ kind: 'success', title: 'Video download started', message: DIRECT_DOWNLOAD_TOAST });
      return;
    }

    // Tier 3: Discover drive item from performance entries, then fetch download URL
    if (driveItem && driveItem.apiBase) {
      const busy = UI.toast({ kind: 'progress', message: 'Looking up the video download link…' });
      // Ask page world to fetch the download URL via CustomEvent
      document.dispatchEvent(new CustomEvent('tceRequestVideoDownloadUrl', {
        detail: { apiBase: driveItem.apiBase }
      }));
      // Wait briefly for response
      const waitResult = await new Promise((resolve) => {
        const handler = (e) => {
          document.removeEventListener('tceVideoDownloadUrlReady', handler);
          resolve(e.detail);
        };
        document.addEventListener('tceVideoDownloadUrlReady', handler);
        setTimeout(() => { document.removeEventListener('tceVideoDownloadUrlReady', handler); resolve(null); }, 5000);
      });
      busy.close();
      if (waitResult && waitResult.downloadUrl) {
        window.open(waitResult.downloadUrl, '_blank');
        UI.toast({ kind: 'success', title: 'Video download started', message: DIRECT_DOWNLOAD_TOAST });
        return;
      }
    }

    // Tier 4: Fall back to MSE capture panel
    console.log('[Teams Chat Extractor] No direct download available, opening capture panel');
    setupVideoDownloadPanel();
  };

  const formatDuration = (sec) => {
    const s = Math.max(0, Math.floor(sec || 0));
    const hh = Math.floor(s / 3600);
    const mm = Math.floor((s % 3600) / 60);
    const ss = String(s % 60).padStart(2, '0');
    return hh ? `${hh}:${String(mm).padStart(2, '0')}:${ss}` : `${mm}:${ss}`;
  };

  // Status line helper (aria-live region inside a panel).
  const makeStatus = (text) => {
    const el = UI.h('div', { class: 'tce-status', role: 'status', 'aria-live': 'polite', text });
    const set = (message, kind) => {
      el.textContent = message;
      el.className = 'tce-status' + (kind ? ` tce-status--${kind}` : '');
    };
    return { el, set };
  };

  const statRow = (label, valueEl) => [UI.h('dt', { text: label }), UI.h('dd', null, [valueEl])];

  const VIDEO_PANEL_ID = 'tce-video-panel';

  const setupVideoDownloadPanel = () => {
    const existing = document.getElementById(VIDEO_PANEL_ID);
    if (existing) {
      UI.applyTheme(UI.getStack());
      existing.querySelector('button:not([disabled])')?.focus();
      return;
    }

    const video = document.querySelector('video');
    if (!video) {
      UI.toast({
        kind: 'error',
        title: 'No video found',
        message: 'Open the recording so the player is on screen, then try again.'
      });
      return;
    }

    const totalDuration = video.duration || 0;
    // Estimate expected segments (roughly 2 seconds per segment)
    const expectedSegments = Math.ceil(totalDuration / 2);

    const videoCountEl = UI.h('span', { text: '0' });
    const audioCountEl = UI.h('span', { text: '0' });
    const durationEl = UI.h('span', { text: expectedSegments ? `0 / ${expectedSegments}` : '0' });
    const progress = UI.progressBar({ label: 'Capture progress', value: 0, max: 100 });
    const status = makeStatus('Start capture to play the video at high speed and collect its segments.');

    const captureBtn = UI.button({ label: 'Start capture', icon: 'play', variant: 'primary', className: 'tce-btn--grow',
      title: 'Best for DRM content - captures decrypted video at 16x speed' });
    const stopBtn = UI.button({ label: 'Stop', icon: 'stop', variant: 'secondary', disabled: true });
    const combinedBtn = UI.button({ label: 'Save combined', icon: 'download', variant: 'primary', className: 'tce-btn--grow', disabled: true,
      title: 'Save video and audio merged into a single MP4' });
    const downloadBtn = UI.button({ label: 'Save separately', variant: 'secondary', className: 'tce-btn--grow', disabled: true,
      title: 'Save video and audio as two files' });
    const directBtn = UI.button({ label: 'Try direct download (non-DRM only)', variant: 'subtle', className: 'tce-btn--small' });
    const mergeSlot = UI.h('div');

    let isCapturing = false;

    // Listen for segment updates from injected script
    const updateHandler = (e) => {
      const { videoCount, audioCount, videoPending, audioPending } = e.detail;
      videoCountEl.textContent = videoPending > 0 ? `${videoCount} (+${videoPending})` : videoCount;
      audioCountEl.textContent = audioPending > 0 ? `${audioCount} (+${audioPending})` : audioCount;

      // Show segment-based progress (more reliable than time)
      const capturedCount = videoCount || 0;
      durationEl.textContent = `${capturedCount} / ${expectedSegments}`;

      // Enable download if we have segments
      if (videoCount > 0 && !isCapturing) {
        downloadBtn.disabled = false;
        combinedBtn.disabled = false;
      }
    };
    window.addEventListener('teamsVideoSegmentUpdate', updateHandler);

    const panel = UI.createPanel({
      id: VIDEO_PANEL_ID,
      title: 'Download video',
      initialFocus: captureBtn,
      onClose: () => {
        if (isCapturing) sendVideoCommand('stopCapture');
        isCapturing = false;
        window.removeEventListener('teamsVideoSegmentUpdate', updateHandler);
      }
    });

    panel.body.append(
      UI.h('p', { text: 'Direct download isn’t available for this recording, so it will be captured while it plays at high speed. Keep this tab in the foreground.' }),
      UI.h('dl', { class: 'tce-stats' }, [
        ...statRow('Video segments', videoCountEl),
        ...statRow('Audio segments', audioCountEl),
        ...statRow('Captured', durationEl),
        ...statRow('Length', UI.h('span', { text: formatDuration(totalDuration) }))
      ]),
      progress.el,
      status.el,
      UI.h('div', { class: 'tce-row' }, [captureBtn, stopBtn]),
      UI.h('div', { class: 'tce-label', text: 'Save' }),
      UI.h('div', { class: 'tce-row' }, [combinedBtn, downloadBtn]),
      mergeSlot
    );
    panel.footer.append(directBtn);

    // Check if capture is available
    checkVideoCaptureAvailable().then(available => {
      if (available) {
        status.set('Ready. Start capture to begin.');
      } else {
        status.set('Waiting for video capture to initialize…');
        // Retry after a short delay
        setTimeout(async () => {
          if (await checkVideoCaptureAvailable()) {
            status.set('Ready. Start capture to begin.');
          }
        }, 1000);
      }
    });

    // Start high-speed capture
    captureBtn.addEventListener('click', async () => {
      // Check availability
      if (!videoCaptureAvailable) {
        const available = await checkVideoCaptureAvailable();
        if (!available) {
          status.set('Video capture isn’t available. Reload the page and try again.', 'error');
          return;
        }
      }

      isCapturing = true;
      captureBtn.disabled = true;
      stopBtn.disabled = false;
      downloadBtn.disabled = true;
      mergeSlot.replaceChildren();

      status.set('Starting capture…');

      // Clear previous capture
      await sendVideoCommand('clear');

      // Start capture
      await sendVideoCommand('startCapture', { downloadImmediately: true });

      // Configure video for high-speed playback
      const wasMuted = video.muted;
      const wasTime = video.currentTime;
      const wasRate = video.playbackRate;

      video.muted = true;
      video.currentTime = 0;

      try {
        await video.play();
        // Set playback rate AFTER play starts (some browsers reject high rates before play)
        video.playbackRate = 16; // Start with 16x (most reliable)
        const actualRate = video.playbackRate;
        status.set(`Playing at ${actualRate}x speed…`);
      } catch (e) {
        status.set('Click the video’s play button first, then try again.', 'warning');
        isCapturing = false;
        captureBtn.disabled = false;
        stopBtn.disabled = true;
        return;
      }

      // Monitor progress
      const progressInterval = setInterval(async () => {
        if (!isCapturing) {
          clearInterval(progressInterval);
          return;
        }

        const currentTime = video.currentTime;
        const pct = (currentTime / totalDuration) * 100;
        progress.set(Math.min(pct, 100), 100);

        const stats = await sendVideoCommand('getStats');
        if (stats) {
          videoCountEl.textContent = stats.videoCount || 0;
          audioCountEl.textContent = stats.audioCount || 0;
          status.set(`${Math.round(pct)}% · ${stats.videoCount || 0} video, ${stats.audioCount || 0} audio segments`);
        }

        // Check if done
        if (currentTime >= totalDuration - 1 || video.ended) {
          clearInterval(progressInterval);
          video.pause();
          video.muted = wasMuted;
          video.playbackRate = wasRate;
          video.currentTime = wasTime;

          await sendVideoCommand('stopCapture');
          isCapturing = false;
          captureBtn.disabled = false;
          stopBtn.disabled = true;

          const finalStats = await sendVideoCommand('getStats');
          if (finalStats && finalStats.videoCount > 0) {
            status.set(`Capture complete: ${finalStats.videoCount} video + ${finalStats.audioCount} audio segments. Choose how to save.`, 'success');
            downloadBtn.disabled = false;
            combinedBtn.disabled = false;
            combinedBtn.focus();
          } else {
            status.set('Capture finished but no segments were captured.', 'warning');
          }
        }
      }, 500);
    });

    // Stop capture
    stopBtn.addEventListener('click', async () => {
      await sendVideoCommand('stopCapture');
      video.pause();
      isCapturing = false;
      captureBtn.disabled = false;
      stopBtn.disabled = true;

      const stats = await sendVideoCommand('getStats');
      if (stats && stats.videoCount > 0) {
        status.set(`Stopped. ${stats.videoCount} segments captured. Choose how to save.`);
        downloadBtn.disabled = false;
        combinedBtn.disabled = false;
        combinedBtn.focus();
      } else {
        status.set('Capture stopped.');
        captureBtn.focus();
      }
    });

    // Direct download button
    directBtn.addEventListener('click', async () => {
      status.set('Testing direct API access…');

      const progressHandler = (e) => {
        status.set(e.detail?.message || 'Processing…');
      };
      document.addEventListener('teamsVideoDownloadProgress', progressHandler);

      const result = await sendVideoCommand('directDownload');
      document.removeEventListener('teamsVideoDownloadProgress', progressHandler);

      if (result) {
        console.log('[Direct Download Result]', result);
        if (result.success) {
          if (result.isDrmProtected) {
            // DRM detected - guide user to MSE capture
            status.set('This video is DRM-protected. Use Start capture instead.', 'warning');
          } else {
            status.set(result.message, 'success');
          }
        } else {
          status.set(result.error || result.message || 'Direct download failed.', 'error');
        }
      } else {
        status.set('No response from the page. Reload and try again.', 'error');
      }
    });

    // Save captured files (separate video + audio)
    downloadBtn.addEventListener('click', async () => {
      downloadBtn.disabled = true;
      status.set('Starting download…');

      // Listen for progress updates
      const progressHandler = (e) => {
        status.set(e.detail?.message || 'Processing…');
      };
      document.addEventListener('teamsVideoDownloadProgress', progressHandler);

      try {
        // Use the downloadFiles command (handled by videoDownloadOverride.js)
        const result = await sendVideoCommand('downloadFiles');

        document.removeEventListener('teamsVideoDownloadProgress', progressHandler);

        if (!result) {
          status.set('No response from the page. Reload and try again.', 'error');
        } else if (result.error) {
          status.set(result.error, 'error');
        } else if (result.success) {
          videoCountEl.textContent = result.videoCount;
          audioCountEl.textContent = result.audioCount || 'N/A';
          progress.set(100, 100);
          status.set('Saved video and audio as separate files.', 'success');
          mergeSlot.replaceChildren(UI.codeDisclosure(
            'Show manual merge command',
            `ffmpeg -i "${result.title}-video.mp4" -i "${result.title}-audio.mp4" -c copy "${result.title}.mp4"`
          ));
        }
      } catch (err) {
        document.removeEventListener('teamsVideoDownloadProgress', progressHandler);
        console.error('[Teams Chat Extractor] Save failed:', err);
        status.set(err.message, 'error');
      }

      downloadBtn.disabled = false;
    });

    // Save combined file (video+audio muxed together)
    combinedBtn.addEventListener('click', async () => {
      combinedBtn.disabled = true;
      downloadBtn.disabled = true;
      status.set('Preparing combined file…');

      const progressHandler = (e) => {
        status.set(e.detail?.message || 'Processing…');
      };
      document.addEventListener('teamsVideoDownloadProgress', progressHandler);

      try {
        const result = await sendVideoCommand('downloadCombined');

        document.removeEventListener('teamsVideoDownloadProgress', progressHandler);

        if (!result) {
          status.set('No response from the page. Reload and try again.', 'error');
        } else if (result.error) {
          status.set(result.error, 'error');
        } else if (result.success) {
          progress.set(100, 100);
          status.set(`Downloaded ${result.title}.mp4 (${result.sizeMB} MB)`, 'success');
        }
      } catch (err) {
        document.removeEventListener('teamsVideoDownloadProgress', progressHandler);
        console.error('[Teams Chat Extractor] Combined save failed:', err);
        status.set(err.message, 'error');
      }

      combinedBtn.disabled = false;
      downloadBtn.disabled = false;
    });
  };

  // === BATCH TRANSCRIPT PANEL ===
  const BATCH_PANEL_ID = 'tce-batch-panel';

  let zipModulePromise = null;
  const loadZip = () => {
    zipModulePromise = zipModulePromise || import(chrome.runtime.getURL('src/modules/zip.js'));
    return zipModulePromise;
  };

  const setupBatchTranscriptPanel = () => {
    const existing = document.getElementById(BATCH_PANEL_ID);
    if (existing) {
      UI.applyTheme(UI.getStack());
      existing.querySelector('button:not([disabled])')?.focus();
      return;
    }

    const totalEl = UI.h('span', { text: '–' });
    const foundEl = UI.h('span', { text: '–' });
    const currentEl = UI.h('span', { text: '0 / –' });
    const progress = UI.progressBar({ label: 'Batch progress', value: 0, max: 1 });
    const status = makeStatus('Start to find every meeting in this series and extract each transcript.');
    const logEl = UI.h('ul', { class: 'tce-log', 'aria-label': 'Activity log' });

    const startBtn = UI.button({ label: 'Start', icon: 'play', variant: 'primary', className: 'tce-btn--grow' });
    const cancelBtn = UI.button({ label: 'Cancel', variant: 'secondary', disabled: true });

    const formatName = UI.uid('fmt');
    const formatGroup = UI.h('div', { class: 'tce-segmented', role: 'radiogroup', 'aria-label': 'File format' },
      [['vtt', 'VTT'], ['txt', 'TXT'], ['both', 'Both']].map(([value, label], i) =>
        UI.h('label', null, [
          UI.h('input', { type: 'radio', name: formatName, value, checked: i === 0 }),
          UI.h('span', { text: label })
        ])
      )
    );
    const zipBtn = UI.button({ label: 'Download .zip', icon: 'zip', variant: 'primary', className: 'tce-btn--grow', disabled: true });
    const separateBtn = UI.button({ label: 'Save as separate files', variant: 'subtle', className: 'tce-btn--small', disabled: true });

    let batchResults = null;
    let total = 0;

    const addLog = (text, type = 'info') => {
      const line = UI.h('li', { class: `tce-log__${type}`, text, title: text });
      logEl.appendChild(line);
      logEl.scrollTop = logEl.scrollHeight;
    };

    // Listen for progress events from the injected script
    const progressHandler = (e) => {
      const d = e.detail;
      if (d.message) status.set(d.message);

      if (d.total) {
        total = d.total;
        totalEl.textContent = d.total;
      }
      if (d.current && total) {
        currentEl.textContent = `${d.current} / ${total}`;
        progress.set(d.current, total);
      }
      if (d.transcriptsFound !== undefined) {
        foundEl.textContent = d.transcriptsFound;
      }

      if (d.phase === 'extracted') {
        const src = d.source === 'api' ? '[API]' : '[DOM]';
        addLog(`${src} ${d.currentMeeting}: ${d.entryCount} entries`, 'success');
      } else if (d.phase === 'api_attempt') {
        addLog(`Trying API for ${d.currentMeeting}…`, 'info');
      } else if (d.phase === 'skipped') {
        addLog(`${d.currentMeeting}: no transcript`, 'warn');
      } else if (d.phase === 'complete') {
        const apiNote = d.apiSuccessCount > 0 ? ` (${d.apiSuccessCount} API, ${d.domFallbackCount} DOM)` : '';
        addLog(`Complete: ${d.transcriptsFound}/${d.total} meetings had transcripts${apiNote}`, 'success');
        foundEl.textContent = d.transcriptsFound;
        progress.set(total || 1, total || 1);
      }
    };
    document.addEventListener('teamsBatchTranscriptProgress', progressHandler);

    const panel = UI.createPanel({
      id: BATCH_PANEL_ID,
      title: 'Batch transcripts',
      initialFocus: startBtn,
      onClose: () => {
        sendBatchCommand('cancel');
        document.removeEventListener('teamsBatchTranscriptProgress', progressHandler);
      }
    });

    panel.body.append(
      UI.h('dl', { class: 'tce-stats' }, [
        ...statRow('Meetings', totalEl),
        ...statRow('With transcript', foundEl),
        ...statRow('Processed', currentEl)
      ]),
      progress.el,
      status.el,
      logEl,
      UI.h('div', { class: 'tce-row' }, [startBtn, cancelBtn]),
      UI.h('div', { class: 'tce-label', text: 'Save' }),
      UI.h('div', { class: 'tce-row' }, [formatGroup, zipBtn]),
      UI.h('div', { class: 'tce-row' }, [separateBtn])
    );

    const selectedFormats = () => {
      const v = formatGroup.querySelector('input:checked')?.value || 'vtt';
      return v === 'both' ? ['vtt', 'txt'] : [v];
    };

    const buildFiles = () => {
      const withTranscript = batchResults.results.filter((r) => r.hasTranscript);
      const series = UI.sanitizeFilename(batchResults.seriesName, 'Meeting', 80);
      const files = [];
      for (const r of withTranscript) {
        const base = `${series} - ${UI.sanitizeFilename(r.meetingDate, `Meeting ${r.index + 1}`, 80)}`;
        for (const fmt of selectedFormats()) {
          files.push({ name: `${base}.${fmt}`, data: fmt === 'txt' ? r.txt : r.vtt });
        }
      }
      return { series, files, count: withTranscript.length };
    };

    // Start
    startBtn.addEventListener('click', async () => {
      startBtn.disabled = true;
      cancelBtn.disabled = false;
      zipBtn.disabled = true;
      separateBtn.disabled = true;
      logEl.replaceChildren();
      batchResults = null;
      total = 0;
      progress.set(null);
      status.set('Finding meetings…');

      addLog('Starting batch extraction…');
      const result = await sendBatchCommand('start');

      startBtn.disabled = false;
      cancelBtn.disabled = true;

      if (!result) {
        status.set('No response from the batch script. Reload the page and try again.', 'error');
        addLog('Error: no response', 'error');
        progress.set(0, 1);
        return;
      }
      if (result.error) {
        status.set(result.error, 'error');
        addLog(`Error: ${result.error}`, 'error');
        progress.set(0, 1);
        return;
      }
      if (result.success) {
        batchResults = result;
        const withTranscript = result.results.filter((r) => r.hasTranscript);
        const n = result.results.length || 1;
        progress.set(n, n);
        zipBtn.disabled = withTranscript.length === 0;
        separateBtn.disabled = withTranscript.length === 0;
        if (withTranscript.length) {
          status.set(`${withTranscript.length} of ${result.results.length} meetings have transcripts. Ready to download.`, 'success');
          zipBtn.focus();
        } else {
          status.set('None of these meetings has a transcript.', 'warning');
        }
      }
    });

    // Cancel
    cancelBtn.addEventListener('click', () => {
      sendBatchCommand('cancel');
      cancelBtn.disabled = true;
      addLog('Cancelled by user', 'warn');
      status.set('Cancelling…', 'warning');
    });

    // Download everything as a single .zip
    zipBtn.addEventListener('click', async () => {
      if (!batchResults) return;
      zipBtn.disabled = true;
      try {
        const { createZipBlob } = await loadZip();
        const { series, files, count } = buildFiles();
        const zipName = `${series} transcripts.zip`;
        saveBlob(createZipBlob(files), zipName);
        addLog(`Saved ${zipName} (${count} meetings, ${files.length} files)`, 'success');
        UI.toast({ kind: 'success', message: `Saved ${zipName}` });
      } catch (err) {
        console.error('[Teams Chat Extractor] Zip failed:', err);
        addLog(`Error: ${err.message}`, 'error');
        UI.toast({ kind: 'error', title: 'Couldn’t create the .zip', message: `${err.message}. Try “Save as separate files”.` });
      } finally {
        zipBtn.disabled = false;
      }
    });

    // Secondary: one download per file (Chrome may ask to allow multiple downloads)
    separateBtn.addEventListener('click', async () => {
      if (!batchResults) return;
      separateBtn.disabled = true;
      const { files } = buildFiles();
      for (const f of files) {
        downloadFile(f.data, f.name);
        await new Promise((r) => setTimeout(r, 250));
      }
      addLog(`Downloaded ${files.length} files`);
      separateBtn.disabled = false;
    });
  };

  // Inject toolbar buttons into Teams native UI
  if (isVideoPage() || isTeamsPage()) {
    // Initial attempt after page loads
    setTimeout(tryInjectToolbarButtons, 2000);

    // Watch for toolbar elements appearing (React re-renders, navigation)
    const observer = new MutationObserver(() => {
      // Check if recap toolbar exists but hasn't been injected yet
      const recapToolbar = document.querySelector('[aria-label="Audio recap"]')?.closest('.fui-Toolbar');
      if (recapToolbar && !recapToolbar.hasAttribute('data-tce-injected')) {
        injectRecapToolbarButtons();
      }
      // Check if transcript actions bar exists but hasn't been injected yet
      const transcriptBar = document.querySelector('[aria-label="Transcript actions"]');
      if (transcriptBar && !transcriptBar.hasAttribute('data-tce-injected')) {
        injectTranscriptActionButtons();
      }
      // Check if Teams recap page needs buttons
      if (!document.querySelector('.tce-recap-bar')) {
        injectTeamsRecapButtons();
      }
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }
})();
