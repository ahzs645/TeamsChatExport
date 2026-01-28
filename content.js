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
  return window.location.href.includes('teams.microsoft.com');
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

  // --- Message Listener ---
  chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
    switch (request.action) {
      case 'sharePointTokenCaptured':
        // Forward from background script to page context
        forwardTokenToPage(request.host, request.token, request.capturedAt);
        sendResponse({ received: true });
        return true;

      case 'getState':
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
        (async () => {
          try {
            const result = await extractionEngine.extractActiveChat();
            if (result) {
              const response = await chrome.runtime.sendMessage({
                action: 'openResults',
                data: result
              });
              if (response && response.success) {
                console.log('Results page opened successfully');
              }
            } else {
              console.log('No messages extracted');
            }
          } catch (error) {
            console.error('Error extracting active chat:', error);
          }
        })();
        sendResponse({status: 'extracting'});
        return true;

      case 'getCurrentChat':
        const currentChat = TeamsVariantDetector.getCurrentChatTitle();
        sendResponse({chatTitle: currentChat});
        return true;

      case 'startBatchTranscript':
        setupBatchTranscriptPanel();
        sendResponse({ status: 'panel_opened' });
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

      case 'getTranscriptStatus':
        // Check if transcript data is available
        const transcriptContainer = document.getElementById('teams-chat-exporter-transcript-data');
        const hasTranscript = transcriptContainer && (transcriptContainer.textContent || transcriptContainer.getAttribute('data-vtt'));
        const hasVideo = !!document.querySelector('video');
        sendResponse({
          available: !!hasTranscript,
          hasVideo: hasVideo,
          isVideoPage: isVideoPage()
        });
        return true;

      case 'extractTranscript':
        // Extract transcript and return data
        const container = document.getElementById('teams-chat-exporter-transcript-data');
        if (!container) {
          sendResponse({
            success: false,
            error: 'Transcript not ready. Start playback to load the transcript.'
          });
          return true;
        }
        const vttData = container.getAttribute('data-vtt') || container.textContent;
        const txtData = container.getAttribute('data-txt') || container.textContent;

        // Get video title for filename
        const titleEl = document.querySelector('h1[class*="videoTitleViewModeHeading"] label');
        const videoTitle = titleEl?.innerText?.trim() || document.title?.trim() || 'transcript';
        const safeTitle = videoTitle.replace(/[^a-zA-Z0-9\s-]/g, '').trim();

        sendResponse({
          success: true,
          vtt: vttData,
          txt: txtData,
          title: safeTitle
        });
        return true;

      default:
        sendResponse({status: 'unknown action'});
        return true;
    }
  });

  console.log('Teams Chat Extractor ready');

  // === NATIVE TOOLBAR INTEGRATION ===

  // SVG icon paths (20x20 viewBox, clean outline style)
  const SVG_ICONS = {
    copy: '<svg viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg"><rect x="6" y="6" width="10" height="12" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M4 14V4.5A1.5 1.5 0 0 1 5.5 3H12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
    download: '<svg viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg"><path d="M10 3v10m0 0l-3.5-3.5M10 13l3.5-3.5M4 16h12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    batch: '<svg viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg"><rect x="4" y="5" width="10" height="12" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M7 3h7.5A1.5 1.5 0 0 1 16 4.5V14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/><path d="M7 10h4M7 13h2" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
    video: '<svg viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="5" width="11" height="10" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M13 9l5-3v8l-5-3" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/></svg>',
    chevron: '<svg viewBox="0 0 8 5" xmlns="http://www.w3.org/2000/svg" style="width:8px;height:5px;margin-left:2px"><path d="M1 1l3 3 3-3" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/></svg>'
  };

  // Shared helpers for transcript actions
  const getVideoTitle = () => {
    const heading = document.querySelector('h1[class*="videoTitleViewModeHeading"] label');
    const h2 = document.querySelector('h2');
    const videoTitle = heading?.innerText?.trim() || h2?.innerText?.trim() || document.title?.trim() || 'transcript';
    return videoTitle.replace(/[^a-zA-Z0-9\s-]/g, '').trim();
  };

  const getTranscriptData = () => {
    const container = document.getElementById('teams-chat-exporter-transcript-data');
    if (!container) return null;
    return {
      vtt: container.getAttribute('data-vtt') || container.textContent,
      txt: container.getAttribute('data-txt') || container.textContent
    };
  };

  const downloadFile = (content, filename) => {
    const a = document.createElement('a');
    a.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(content));
    a.setAttribute('download', filename);
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // Copy transcript handler
  const handleCopy = async (btn) => {
    const data = getTranscriptData();
    if (!data) {
      alert('Transcript not ready yet. Start playback to load the transcript, then try again.');
      return;
    }
    try {
      await navigator.clipboard.writeText(data.vtt);
      btn.classList.add('tce-copied');
      const origHTML = btn.innerHTML;
      if (btn.classList.contains('tce-cmd-btn')) {
        btn.textContent = 'Copied!';
      }
      setTimeout(() => {
        btn.classList.remove('tce-copied');
        if (btn.classList.contains('tce-cmd-btn')) {
          btn.innerHTML = origHTML;
        }
      }, 1500);
    } catch (err) {
      console.error('[Teams Chat Extractor] Failed to copy:', err);
    }
  };

  // Download VTT handler
  const handleDownloadVTT = () => {
    const data = getTranscriptData();
    if (!data) {
      alert('Transcript not ready yet. Start playback to load the transcript, then try again.');
      return;
    }
    downloadFile(data.vtt, `transcript-${getVideoTitle()}.vtt`);
  };

  // Download TXT handler
  const handleDownloadTXT = () => {
    const data = getTranscriptData();
    if (!data) {
      alert('Transcript not ready yet. Start playback to load the transcript, then try again.');
      return;
    }
    downloadFile(data.txt, `transcript-${getVideoTitle()}.txt`);
  };

  // Close any open download dropdown
  const closeDropdowns = () => {
    document.querySelectorAll('.tce-download-dropdown').forEach(d => d.remove());
  };

  // Create download dropdown anchored to a button
  const createDownloadDropdown = (anchorBtn) => {
    closeDropdowns();
    const dropdown = document.createElement('div');
    dropdown.className = 'tce-download-dropdown';

    const vttItem = document.createElement('button');
    vttItem.className = 'tce-download-dropdown-item';
    vttItem.textContent = 'Download VTT';
    vttItem.addEventListener('click', () => { handleDownloadVTT(); closeDropdowns(); });

    const txtItem = document.createElement('button');
    txtItem.className = 'tce-download-dropdown-item';
    txtItem.textContent = 'Download TXT';
    txtItem.addEventListener('click', () => { handleDownloadTXT(); closeDropdowns(); });

    dropdown.appendChild(vttItem);
    dropdown.appendChild(txtItem);
    anchorBtn.appendChild(dropdown);

    // Close on outside click
    const outsideHandler = (e) => {
      if (!anchorBtn.contains(e.target)) {
        closeDropdowns();
        document.removeEventListener('click', outsideHandler, true);
      }
    };
    setTimeout(() => document.addEventListener('click', outsideHandler, true), 0);
  };

  // === LOCATION 1: Recap Action Toolbar ===
  const injectRecapToolbarButtons = () => {
    // Find toolbar by anchor button
    const anchorBtn = document.querySelector('[aria-label="Audio recap"]');
    const toolbar = anchorBtn?.closest('.fui-Toolbar');
    if (!toolbar || toolbar.hasAttribute('data-tce-injected')) return;
    toolbar.setAttribute('data-tce-injected', 'true');

    // Create icon buttons
    const makeBtn = (icon, title, handler) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'tce-toolbar-btn';
      btn.title = title;
      btn.innerHTML = icon;
      btn.addEventListener('click', handler);
      return btn;
    };

    const copyBtn = makeBtn(SVG_ICONS.copy, 'Copy transcript', (e) => handleCopy(e.currentTarget));

    const downloadBtn = makeBtn(SVG_ICONS.download + SVG_ICONS.chevron, 'Download transcript', (e) => {
      e.stopPropagation();
      createDownloadDropdown(e.currentTarget);
    });

    const batchBtn = makeBtn(SVG_ICONS.batch, 'Batch download all transcripts', () => setupBatchTranscriptPanel());

    const videoBtn = makeBtn(SVG_ICONS.video, 'Download video', () => setupVideoDownloadPanel());

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
    menubar.appendChild(sep);

    const makeCmdBtn = (label, icon, handler) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'tce-cmd-btn';
      btn.innerHTML = icon + ' ' + label;
      btn.addEventListener('click', handler);
      return btn;
    };

    menubar.appendChild(makeCmdBtn('Copy', SVG_ICONS.copy, (e) => handleCopy(e.currentTarget)));
    menubar.appendChild(makeCmdBtn('Download VTT', SVG_ICONS.download, handleDownloadVTT));
    menubar.appendChild(makeCmdBtn('Download TXT', SVG_ICONS.download, handleDownloadTXT));
    menubar.appendChild(makeCmdBtn('Batch Download', SVG_ICONS.batch, () => setupBatchTranscriptPanel()));

    console.log('[Teams Chat Extractor] Injected buttons into transcript actions menubar');
  };

  // Try injecting into both toolbar locations
  const tryInjectToolbarButtons = () => {
    injectRecapToolbarButtons();
    injectTranscriptActionButtons();
  };

  // === VIDEO DOWNLOAD PANEL ===
  let videoDownloadPanelOpen = false;
  let videoCaptureAvailable = false;

  // Helper to send commands to the injected video capture script
  const sendVideoCommand = (command, data = {}, timeoutMs = 2000) => {
    // Use longer timeout for download operations
    if (command === 'downloadFiles' || command === 'directDownload') {
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

  const setupVideoDownloadPanel = () => {
    if (videoDownloadPanelOpen || document.getElementById('video-download-panel')) {
      return;
    }
    videoDownloadPanelOpen = true;

    const video = document.querySelector('video');
    if (!video) {
      alert('No video found on this page.');
      videoDownloadPanelOpen = false;
      return;
    }

    const totalDuration = video.duration || 0;
    const totalMins = Math.floor(totalDuration / 60);
    const totalSecs = Math.floor(totalDuration % 60);

    const panel = document.createElement('div');
    panel.id = 'video-download-panel';
    panel.innerHTML = `
      <div class="vdp-header">
        <span>🎬 Video Downloader</span>
        <button class="vdp-close" id="vdp-close">\u00D7</button>
      </div>
      <div class="vdp-stats">
        <div>📹 Video segments: <span id="vdp-video-count">0</span></div>
        <div>🎵 Audio segments: <span id="vdp-audio-count">0</span></div>
        <div>⏱️ Captured: <span id="vdp-duration">0:00</span> / ${totalMins}:${String(totalSecs).padStart(2, '0')}</div>
      </div>
      <div class="vdp-progress">
        <div class="vdp-progress-fill" id="vdp-progress-fill"></div>
      </div>
      <div class="vdp-status" id="vdp-status">Click Start - plays at high speed while downloading segments</div>
      <div class="vdp-buttons">
        <button class="vdp-btn vdp-capture" id="vdp-capture" title="Best for DRM content - captures decrypted video at 16x speed">🚀 Start Capture</button>
        <button class="vdp-btn vdp-stop" id="vdp-stop" disabled>⏹ Stop</button>
      </div>
      <div class="vdp-buttons" style="margin-top: 8px;">
        <button class="vdp-btn vdp-download" id="vdp-combined" disabled style="background: linear-gradient(135deg, #00ff88, #00cc6a);" title="Save combined video+audio as single file">💾 Save Combined</button>
        <button class="vdp-btn" id="vdp-download" disabled style="background: #666;" title="Save video and audio as separate files">📁 Separate</button>
      </div>
      <div class="vdp-buttons" style="margin-top: 8px;">
        <button class="vdp-btn" id="vdp-analyze" style="background: #555; font-size: 10px;" title="View captured URL patterns in console">🔍</button>
        <button class="vdp-btn" id="vdp-direct" style="background: linear-gradient(135deg, #ff6b6b, #ee5a5a); font-size: 10px;" title="Fast API download - only works for non-DRM content">⚡ Direct</button>
      </div>
      <div class="vdp-tip">🚀 Start → 💾 Save Combined (auto-merges video+audio)</div>
    `;

    document.body.appendChild(panel);

    // Get elements
    const closeBtn = document.getElementById('vdp-close');
    const captureBtn = document.getElementById('vdp-capture');
    const stopBtn = document.getElementById('vdp-stop');
    const downloadBtn = document.getElementById('vdp-download');
    const combinedBtn = document.getElementById('vdp-combined');
    const videoCountEl = document.getElementById('vdp-video-count');
    const audioCountEl = document.getElementById('vdp-audio-count');
    const durationEl = document.getElementById('vdp-duration');
    const progressEl = document.getElementById('vdp-progress-fill');
    const statusEl = document.getElementById('vdp-status');

    let isCapturing = false;

    // Estimate expected segments (roughly 2 seconds per segment)
    const expectedSegments = Math.ceil(totalDuration / 2);

    // Listen for segment updates from injected script
    const updateHandler = (e) => {
      const { videoCount, audioCount, videoPending, audioPending } = e.detail;
      videoCountEl.textContent = videoPending > 0 ? `${videoCount} (+${videoPending})` : videoCount;
      audioCountEl.textContent = audioPending > 0 ? `${audioCount} (+${audioPending})` : audioCount;

      // Show segment-based progress (more reliable than time)
      const capturedCount = videoCount || 0;
      durationEl.textContent = `${capturedCount}/${expectedSegments} segs`;

      // Enable download if we have segments
      if (videoCount > 0 && !isCapturing) {
        downloadBtn.disabled = false;
        combinedBtn.disabled = false;
      }
    };
    window.addEventListener('teamsVideoSegmentUpdate', updateHandler);

    // Check if capture is available
    checkVideoCaptureAvailable().then(available => {
      if (available) {
        statusEl.textContent = 'Ready! Click Start Capture to begin.';
      } else {
        statusEl.textContent = '⏳ Waiting for video capture to initialize...';
        // Retry after a short delay
        setTimeout(async () => {
          if (await checkVideoCaptureAvailable()) {
            statusEl.textContent = 'Ready! Click Start Capture to begin.';
          }
        }, 1000);
      }
    });

    // Close panel
    closeBtn.addEventListener('click', async () => {
      if (isCapturing) {
        await sendVideoCommand('stopCapture');
      }
      window.removeEventListener('teamsVideoSegmentUpdate', updateHandler);
      panel.remove();
      videoDownloadPanelOpen = false;
    });

    // Start high-speed capture
    captureBtn.addEventListener('click', async () => {
      // Check availability
      if (!videoCaptureAvailable) {
        const available = await checkVideoCaptureAvailable();
        if (!available) {
          statusEl.textContent = '❌ Video capture not available. Reload page and try again.';
          return;
        }
      }

      isCapturing = true;
      captureBtn.disabled = true;
      stopBtn.disabled = false;
      downloadBtn.disabled = true;

      statusEl.textContent = '🚀 Starting capture...';

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
        statusEl.textContent = `🚀 Playing at ${actualRate}x speed...`;
      } catch (e) {
        statusEl.textContent = '⚠️ Click the video play button first, then try again.';
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
        const progress = (currentTime / totalDuration) * 100;
        progressEl.style.width = `${Math.min(progress, 100)}%`;

        const stats = await sendVideoCommand('getStats');
        if (stats) {
          videoCountEl.textContent = stats.videoCount || 0;
          audioCountEl.textContent = stats.audioCount || 0;
          statusEl.textContent = `⚡ ${Math.round(progress)}% - ${stats.videoCount || 0} video, ${stats.audioCount || 0} audio segments`;
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
            statusEl.textContent = `✅ Complete! ${finalStats.videoCount} video + ${finalStats.audioCount} audio. Click Save.`;
            downloadBtn.disabled = false;
            combinedBtn.disabled = false;
          } else {
            statusEl.textContent = '⚠️ Capture finished but no segments captured.';
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
        statusEl.textContent = `⏹ Stopped. ${stats.videoCount} segments captured. Click Save.`;
        downloadBtn.disabled = false;
        combinedBtn.disabled = false;
      } else {
        statusEl.textContent = '⏹ Capture stopped.';
      }
    });

    // Analyze URLs button
    const analyzeBtn = document.getElementById('vdp-analyze');
    analyzeBtn.addEventListener('click', async () => {
      statusEl.textContent = '🔍 Analyzing captured URLs...';
      const result = await sendVideoCommand('analyzeUrls');
      if (result) {
        console.log('[URL Analysis]', result);
        const urlCount = result.capturedUrls?.length || 0;
        statusEl.textContent = `🔍 Found ${urlCount} URLs. Check console (F12) for details.`;
      } else {
        statusEl.textContent = '❌ No data - play video for a few seconds first';
      }
    });

    // Direct download button
    const directBtn = document.getElementById('vdp-direct');
    directBtn.addEventListener('click', async () => {
      statusEl.textContent = '⚡ Testing direct API access...';

      const progressHandler = (e) => {
        statusEl.textContent = '⚡ ' + (e.detail?.message || 'Processing...');
      };
      document.addEventListener('teamsVideoDownloadProgress', progressHandler);

      const result = await sendVideoCommand('directDownload');
      document.removeEventListener('teamsVideoDownloadProgress', progressHandler);

      if (result) {
        console.log('[Direct Download Result]', result);
        if (result.success) {
          if (result.isDrmProtected) {
            // DRM detected - guide user to MSE capture
            statusEl.textContent = `⚠️ DRM detected! Use 🚀 Start Capture instead`;
            statusEl.style.background = 'rgba(255, 193, 7, 0.3)';
          } else {
            statusEl.textContent = `✅ ${result.message}`;
            statusEl.style.background = 'rgba(40, 167, 69, 0.3)';
          }
        } else {
          statusEl.textContent = `❌ ${result.error || result.message}`;
        }
      } else {
        statusEl.textContent = '❌ No response - reload page';
      }
    });

    // Save captured files
    downloadBtn.addEventListener('click', async () => {
      downloadBtn.disabled = true;
      statusEl.textContent = '💾 Starting download...';

      // Listen for progress updates
      const progressHandler = (e) => {
        statusEl.textContent = '💾 ' + (e.detail?.message || 'Processing...');
      };
      document.addEventListener('teamsVideoDownloadProgress', progressHandler);

      try {
        // Use the downloadFiles command (handled by videoDownloadOverride.js)
        const result = await sendVideoCommand('downloadFiles');

        document.removeEventListener('teamsVideoDownloadProgress', progressHandler);

        if (!result) {
          statusEl.textContent = '❌ No response - reload page and try again';
        } else if (result.error) {
          statusEl.textContent = '❌ ' + result.error;
        } else if (result.success) {
          videoCountEl.textContent = result.videoCount;
          audioCountEl.textContent = result.audioCount || 'N/A';
          progressEl.style.width = '100%';
          statusEl.textContent = `✅ Saved! Merge: ffmpeg -i "${result.title}-video.mp4" -i "${result.title}-audio.mp4" -c copy output.mp4`;
        }
      } catch (err) {
        document.removeEventListener('teamsVideoDownloadProgress', progressHandler);
        console.error('[Teams Chat Extractor] Save failed:', err);
        statusEl.textContent = '❌ ' + err.message;
      }

      downloadBtn.disabled = false;
    });

    // Save combined file (video+audio muxed together)
    combinedBtn.addEventListener('click', async () => {
      combinedBtn.disabled = true;
      downloadBtn.disabled = true;
      statusEl.textContent = '💾 Preparing combined file...';

      const progressHandler = (e) => {
        statusEl.textContent = '💾 ' + (e.detail?.message || 'Processing...');
      };
      document.addEventListener('teamsVideoDownloadProgress', progressHandler);

      try {
        const result = await sendVideoCommand('downloadCombined');

        document.removeEventListener('teamsVideoDownloadProgress', progressHandler);

        if (!result) {
          statusEl.textContent = '❌ No response - reload page and try again';
        } else if (result.error) {
          statusEl.textContent = '❌ ' + result.error;
        } else if (result.success) {
          progressEl.style.width = '100%';
          statusEl.textContent = `✅ Downloaded: ${result.title}.mp4 (${result.sizeMB} MB)`;
        }
      } catch (err) {
        document.removeEventListener('teamsVideoDownloadProgress', progressHandler);
        console.error('[Teams Chat Extractor] Combined save failed:', err);
        statusEl.textContent = '❌ ' + err.message;
      }

      combinedBtn.disabled = false;
      downloadBtn.disabled = false;
    });
  }

  // === BATCH TRANSCRIPT PANEL ===
  let batchPanelOpen = false;

  const setupBatchTranscriptPanel = () => {
    if (batchPanelOpen || document.getElementById('batch-transcript-panel')) {
      return;
    }
    batchPanelOpen = true;

    const panel = document.createElement('div');
    panel.id = 'batch-transcript-panel';
    panel.innerHTML = `
      <div class="btp-header">
        <span>Batch Transcripts</span>
        <button class="btp-close" id="btp-close">\u00D7</button>
      </div>
      <div class="btp-stats">
        <div>Meetings: <span id="btp-total">--</span></div>
        <div>With transcript: <span id="btp-found">--</span></div>
        <div>Progress: <span id="btp-current">0</span> / <span id="btp-total2">--</span></div>
      </div>
      <div class="btp-progress">
        <div class="btp-progress-fill" id="btp-progress-fill"></div>
      </div>
      <div class="btp-status" id="btp-status">Click Start to enumerate meetings and extract all transcripts.</div>
      <div class="btp-log" id="btp-log"></div>
      <div class="btp-buttons">
        <button class="btp-btn btp-start" id="btp-start">Start</button>
        <button class="btp-btn btp-cancel" id="btp-cancel" disabled>Cancel</button>
      </div>
      <div class="btp-buttons" style="margin-top: 8px;">
        <button class="btp-btn btp-download-all" id="btp-download-vtt" disabled>Download All VTT</button>
        <button class="btp-btn btp-download-all" id="btp-download-txt" disabled>Download All TXT</button>
      </div>
    `;

    document.body.appendChild(panel);

    const closeBtn = document.getElementById('btp-close');
    const startBtn = document.getElementById('btp-start');
    const cancelBtn = document.getElementById('btp-cancel');
    const downloadVttBtn = document.getElementById('btp-download-vtt');
    const downloadTxtBtn = document.getElementById('btp-download-txt');
    const totalEl = document.getElementById('btp-total');
    const total2El = document.getElementById('btp-total2');
    const foundEl = document.getElementById('btp-found');
    const currentEl = document.getElementById('btp-current');
    const progressEl = document.getElementById('btp-progress-fill');
    const statusEl = document.getElementById('btp-status');
    const logEl = document.getElementById('btp-log');

    let batchResults = null;

    const addLog = (text, type = 'info') => {
      const line = document.createElement('div');
      line.className = `btp-log-line btp-log-${type}`;
      line.textContent = text;
      logEl.appendChild(line);
      logEl.scrollTop = logEl.scrollHeight;
    };

    // Listen for progress events from the injected script
    const progressHandler = (e) => {
      const d = e.detail;
      statusEl.textContent = d.message || '';

      if (d.total) {
        totalEl.textContent = d.total;
        total2El.textContent = d.total;
      }
      if (d.current) {
        currentEl.textContent = d.current;
        const pct = (d.current / d.total) * 100;
        progressEl.style.width = `${Math.min(pct, 100)}%`;
      }
      if (d.transcriptsFound !== undefined) {
        foundEl.textContent = d.transcriptsFound;
      }

      if (d.phase === 'extracted') {
        const src = d.source === 'api' ? '[API]' : '[DOM]';
        addLog(`${src} ${d.currentMeeting}: ${d.entryCount} entries`, 'success');
      } else if (d.phase === 'api_attempt') {
        addLog(`Trying API for ${d.currentMeeting}...`, 'info');
      } else if (d.phase === 'skipped') {
        addLog(`${d.currentMeeting}: no transcript`, 'warn');
      } else if (d.phase === 'complete') {
        const apiNote = d.apiSuccessCount > 0 ? ` (${d.apiSuccessCount} API, ${d.domFallbackCount} DOM)` : '';
        addLog(`Complete: ${d.transcriptsFound}/${d.total} meetings had transcripts${apiNote}`, 'success');
        foundEl.textContent = d.transcriptsFound;
      }
    };
    document.addEventListener('teamsBatchTranscriptProgress', progressHandler);

    // Helper to download a file
    const downloadFile = (content, filename) => {
      const a = document.createElement('a');
      a.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(content));
      a.setAttribute('download', filename);
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    };

    // Sanitize text for filenames
    const sanitize = (text) => text.replace(/[<>:"/\\|?*]/g, '-').replace(/\s+/g, ' ').trim().substring(0, 80);

    // Close
    closeBtn.addEventListener('click', () => {
      sendBatchCommand('cancel');
      document.removeEventListener('teamsBatchTranscriptProgress', progressHandler);
      panel.remove();
      batchPanelOpen = false;
    });

    // Start
    startBtn.addEventListener('click', async () => {
      startBtn.disabled = true;
      cancelBtn.disabled = false;
      downloadVttBtn.disabled = true;
      downloadTxtBtn.disabled = true;
      logEl.innerHTML = '';
      batchResults = null;

      addLog('Starting batch extraction...');
      const result = await sendBatchCommand('start');

      startBtn.disabled = false;
      cancelBtn.disabled = true;

      if (!result) {
        statusEl.textContent = 'No response from batch script. Reload page.';
        addLog('Error: no response', 'error');
        return;
      }
      if (result.error) {
        statusEl.textContent = result.error;
        addLog(`Error: ${result.error}`, 'error');
        return;
      }
      if (result.success) {
        batchResults = result;
        progressEl.style.width = '100%';
        const withTranscript = result.results.filter((r) => r.hasTranscript);
        downloadVttBtn.disabled = withTranscript.length === 0;
        downloadTxtBtn.disabled = withTranscript.length === 0;
      }
    });

    // Cancel
    cancelBtn.addEventListener('click', () => {
      sendBatchCommand('cancel');
      cancelBtn.disabled = true;
      addLog('Cancelled by user', 'warn');
    });

    // Download all VTT
    downloadVttBtn.addEventListener('click', () => {
      if (!batchResults) return;
      const withTranscript = batchResults.results.filter((r) => r.hasTranscript);
      const series = sanitize(batchResults.seriesName);
      withTranscript.forEach((r) => {
        const date = sanitize(r.meetingDate);
        downloadFile(r.vtt, `${series} - ${date}.vtt`);
      });
      addLog(`Downloaded ${withTranscript.length} VTT files`);
    });

    // Download all TXT
    downloadTxtBtn.addEventListener('click', () => {
      if (!batchResults) return;
      const withTranscript = batchResults.results.filter((r) => r.hasTranscript);
      const series = sanitize(batchResults.seriesName);
      withTranscript.forEach((r) => {
        const date = sanitize(r.meetingDate);
        downloadFile(r.txt, `${series} - ${date}.txt`);
      });
      addLog(`Downloaded ${withTranscript.length} TXT files`);
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
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }
})();
