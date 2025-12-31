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

  // --- Message Listener ---
  chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
    switch (request.action) {
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

  // === TRANSCRIPT UI SETUP ===
  // Track if user has manually closed the panel
  let transcriptPanelClosed = false;

  // Create floating panel for transcript extraction on video pages
  const setupTranscriptUI = () => {
    // Don't add if already exists or user has closed it
    if (document.querySelector('.transcript-extractor-wrapper') || transcriptPanelClosed) {
      return;
    }

    const wrapper = document.createElement('div');
    wrapper.className = 'transcript-extractor-wrapper';

    const copyButton = document.createElement('button');
    copyButton.className = 'transcript-extractor-button';
    copyButton.textContent = 'Copy';
    copyButton.title = 'Copy transcript to clipboard';

    const downloadVttButton = document.createElement('button');
    downloadVttButton.className = 'transcript-extractor-button';
    downloadVttButton.textContent = 'Download VTT';
    downloadVttButton.title = 'Download as WebVTT subtitle file';

    const downloadTxtButton = document.createElement('button');
    downloadTxtButton.className = 'transcript-extractor-button';
    downloadTxtButton.textContent = 'Download TXT';
    downloadTxtButton.title = 'Download as plain text file';

    const closeButton = document.createElement('button');
    closeButton.className = 'transcript-extractor-button transcript-extractor-close-button';
    closeButton.textContent = '\u00D7';
    closeButton.title = 'Close panel';

    wrapper.appendChild(copyButton);
    wrapper.appendChild(downloadVttButton);
    wrapper.appendChild(downloadTxtButton);
    wrapper.appendChild(closeButton);

    // Helper to get video title for filename
    const getVideoTitle = () => {
      const heading = document.querySelector('h1[class*="videoTitleViewModeHeading"] label');
      const videoTitle = heading?.innerText?.trim() || document.title?.trim() || 'transcript';
      return videoTitle.replace(/[^a-zA-Z0-9\s-]/g, '').trim();
    };

    // Helper to get transcript data
    const getTranscriptData = () => {
      const container = document.getElementById('teams-chat-exporter-transcript-data');
      if (!container) {
        return null;
      }
      return {
        vtt: container.getAttribute('data-vtt') || container.textContent,
        txt: container.getAttribute('data-txt') || container.textContent
      };
    };

    // Helper to download file
    const downloadFile = (content, filename) => {
      const element = document.createElement('a');
      element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(content));
      element.setAttribute('download', filename);
      element.style.display = 'none';
      document.body.appendChild(element);
      element.click();
      document.body.removeChild(element);
    };

    // Copy button handler
    copyButton.addEventListener('click', async () => {
      const data = getTranscriptData();
      if (!data) {
        alert('Transcript not ready yet. Start playback to load the transcript, then try again.');
        return;
      }
      try {
        await navigator.clipboard.writeText(data.vtt);
        copyButton.textContent = 'Copied!';
        setTimeout(() => { copyButton.textContent = 'Copy'; }, 2000);
      } catch (err) {
        console.error('[Teams Chat Extractor] Failed to copy:', err);
        alert('Failed to copy transcript to clipboard.');
      }
    });

    // Download VTT button handler
    downloadVttButton.addEventListener('click', () => {
      const data = getTranscriptData();
      if (!data) {
        alert('Transcript not ready yet. Start playback to load the transcript, then try again.');
        return;
      }
      const filename = `transcript-${getVideoTitle()}.vtt`;
      downloadFile(data.vtt, filename);
    });

    // Download TXT button handler
    downloadTxtButton.addEventListener('click', () => {
      const data = getTranscriptData();
      if (!data) {
        alert('Transcript not ready yet. Start playback to load the transcript, then try again.');
        return;
      }
      const filename = `transcript-${getVideoTitle()}.txt`;
      downloadFile(data.txt, filename);
    });

    // Close button handler
    closeButton.addEventListener('click', () => {
      transcriptPanelClosed = true;
      wrapper.remove();
    });

    // Add video download button
    const videoDownloadButton = document.createElement('button');
    videoDownloadButton.className = 'transcript-extractor-button video-download-btn';
    videoDownloadButton.textContent = '🎬 Video';
    videoDownloadButton.title = 'Download video (captures segments)';

    // Insert before close button
    wrapper.insertBefore(videoDownloadButton, closeButton);

    // Video download handler - opens capture panel
    videoDownloadButton.addEventListener('click', () => {
      setupVideoDownloadPanel();
    });

    document.body.appendChild(wrapper);
    console.log('[Teams Chat Extractor] Transcript UI panel added');
  };

  // === VIDEO DOWNLOAD PANEL ===
  let videoDownloadPanelOpen = false;
  let videoCaptureAvailable = false;

  // Helper to send commands to the injected video capture script
  const sendVideoCommand = (command, data = {}, timeoutMs = 2000) => {
    // Use longer timeout for download operations
    if (command === 'downloadFiles') {
      timeoutMs = 120000; // 2 minutes for downloads
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
        <button class="vdp-btn vdp-capture" id="vdp-capture">🚀 Start Capture</button>
        <button class="vdp-btn vdp-stop" id="vdp-stop" disabled>⏹ Stop</button>
        <button class="vdp-btn vdp-download" id="vdp-download" disabled>💾 Save Files</button>
      </div>
      <div class="vdp-tip">Plays at max speed (browser may cap at 16x). Merge with: ffmpeg -i video.mp4 -i audio.mp4 -c copy output.mp4</div>
    `;

    document.body.appendChild(panel);

    // Get elements
    const closeBtn = document.getElementById('vdp-close');
    const captureBtn = document.getElementById('vdp-capture');
    const stopBtn = document.getElementById('vdp-stop');
    const downloadBtn = document.getElementById('vdp-download');
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
            statusEl.textContent = `✅ Complete! ${finalStats.videoCount} video + ${finalStats.audioCount} audio. Click Save Files.`;
            downloadBtn.disabled = false;
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
        statusEl.textContent = `⏹ Stopped. ${stats.videoCount} segments captured. Click Save Files.`;
        downloadBtn.disabled = false;
      } else {
        statusEl.textContent = '⏹ Capture stopped.';
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
  }

  // Add transcript UI if on a video page
  if (isVideoPage() || isTeamsPage()) {
    // Wait a bit for the page to load, then check for video content
    setTimeout(() => {
      if (isVideoPage() || document.querySelector('video')) {
        setupTranscriptUI();
      }
    }, 2000);

    // Also watch for dynamic video elements
    const observer = new MutationObserver((mutations) => {
      if (document.querySelector('video') && !document.querySelector('.transcript-extractor-wrapper')) {
        setupTranscriptUI();
      }
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }
})();
