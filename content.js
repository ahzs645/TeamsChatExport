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
  // Create floating panel for transcript extraction on video pages
  const setupTranscriptUI = () => {
    // Don't add if already exists
    if (document.querySelector('.transcript-extractor-wrapper')) {
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
      wrapper.remove();
    });

    document.body.appendChild(wrapper);
    console.log('[Teams Chat Extractor] Transcript UI panel added');
  };

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
