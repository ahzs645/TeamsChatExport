/**
 * Microsoft Teams Chat Extractor - Main Content Script
 * 
 * This script automates the process of extracting full chat conversations,
 * including the author, timestamp, and message content for each message.
 *
 * VERSION 7 - IMPROVED MODULAR ARCHITECTURE
 * 
 * This version has been refactored with improved ES6 modules:
 * - TeamsVariantDetector for Teams detection and variant handling
 * - CheckboxManager for checkbox management and UI state
 * - MessageExtractor for message extraction
 * - ScrollManager for auto-scrolling functionality
 * - UIManager for button and UI management
 * - StorageManager for settings persistence
 * - ExtractionEngine for orchestrating the extraction process
 */

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
    checkboxModule,
    extractionModule,
    uiModule,
    storageModule
  ] = await Promise.all([
    import(chrome.runtime.getURL('src/modules/teamsVariantDetector.js')),
    import(chrome.runtime.getURL('src/modules/checkboxManager.js')),
    import(chrome.runtime.getURL('src/modules/extractionEngine.js')),
    import(chrome.runtime.getURL('src/modules/uiManager.js')),
    import(chrome.runtime.getURL('src/modules/storageManager.js'))
  ]);

  const { TeamsVariantDetector } = teamsModule;
  const { CheckboxManager } = checkboxModule;
  const { ExtractionEngine } = extractionModule;
  const { UIManager } = uiModule;
  const { StorageManager } = storageModule;

  console.log('🚀 Teams Chat Extractor - Improved Modular Version Starting...');
  
  // Initialize modules
  const storageManager = new StorageManager();
  const extractionEngine = new ExtractionEngine();
  const checkboxManager = new CheckboxManager();
  const uiManager = new UIManager(extractionEngine, checkboxManager, storageManager);
  
  // Initialize settings with defaults
  window.teamsExtractorSettings = { 
    showCheckboxes: false, 
    autoScroll: true 
  };

  // Load and migrate settings
  await storageManager.migrateSettings();
  const settings = await storageManager.loadSettings();
  
  // Apply loaded settings
  window.teamsExtractorSettings.showCheckboxes = settings.checkboxesEnabled;
  window.teamsExtractorSettings.autoScroll = settings.autoScrollEnabled;
  
  checkboxManager.settings.showCheckboxes = settings.checkboxesEnabled;
  extractionEngine.setAutoScroll(settings.autoScrollEnabled);

  // Inject UI styles
  uiManager.injectStylesheet();

  // --- Message Listener ---
  chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
    switch (request.action) {
      case 'toggleCheckboxes':
        window.teamsExtractorSettings.showCheckboxes = request.enabled;
        checkboxManager.settings.showCheckboxes = request.enabled;
        checkboxManager.toggleCheckboxes(request.enabled);
        
        // Save setting
        storageManager.saveCheckboxState(request.enabled);
        
        if (request.enabled) {
          uiManager.addExtractButton();
        } else {
          uiManager.removeExtractButton();
        }
        break;
        
      case 'getState':
        sendResponse({
          checkboxesEnabled: window.teamsExtractorSettings.showCheckboxes,
          currentChat: TeamsVariantDetector.getCurrentChatTitle()
        });
        break;
        
      case 'selectAll':
        checkboxManager.setAllCheckboxes(true);
        break;
        
      case 'deselectAll':
        checkboxManager.setAllCheckboxes(false);
        break;
        
      case 'queryState':
        checkboxManager.notifyPopupOfState();
        break;
        
      case 'extract':
        uiManager.startExtraction();
        break;
        
      case 'getCurrentChat':
        const currentChat = TeamsVariantDetector.getCurrentChatTitle();
        sendResponse({chatTitle: currentChat});
        break;
        
      case 'toggleAutoScroll':
        const newAutoScrollState = !extractionEngine.getAutoScroll();
        extractionEngine.setAutoScroll(newAutoScrollState);
        window.teamsExtractorSettings.autoScroll = newAutoScrollState;
        
        // Save setting
        storageManager.saveAutoScrollState(newAutoScrollState);
        
        sendResponse({autoScrollEnabled: newAutoScrollState});
        break;
    }
    sendResponse({status: 'received'});
    return true; // Keep the message channel open for async response
  });

  // Set up storage change listener
  storageManager.onStorageChanged((changedSettings) => {
    console.log('⚙️ Settings changed:', changedSettings);
    
    if (changedSettings.checkboxesEnabled !== undefined) {
      window.teamsExtractorSettings.showCheckboxes = changedSettings.checkboxesEnabled;
      checkboxManager.settings.showCheckboxes = changedSettings.checkboxesEnabled;
    }
    
    if (changedSettings.autoScrollEnabled !== undefined) {
      window.teamsExtractorSettings.autoScroll = changedSettings.autoScrollEnabled;
      extractionEngine.setAutoScroll(changedSettings.autoScrollEnabled);
    }
  });

  // Initialize UI if checkboxes are enabled
  if (settings.checkboxesEnabled) {
    // Wait a bit for Teams to load, then add UI elements
    setTimeout(() => {
      checkboxManager.addCheckboxes();
      uiManager.addExtractButton();
    }, 2000);
  }
  
  // Set up mutation observer for dynamic content
  const observer = new MutationObserver((mutations) => {
    // Only trigger if significant changes occurred
    let shouldUpdate = false;
    for (const mutation of mutations) {
      if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
        // Check if any added nodes contain relevant elements
        for (const node of mutation.addedNodes) {
          if (node.nodeType === Node.ELEMENT_NODE) {
            if (node.querySelector && (
              node.querySelector('[data-tid="titlebar-end-slot"]') ||
              node.querySelector('.fui-TreeItem') ||
              node.classList?.contains('fui-TreeItem')
            )) {
              shouldUpdate = true;
              break;
            }
          }
        }
      }
      if (shouldUpdate) break;
    }
    
    if (shouldUpdate) {
      checkboxManager.addCheckboxes();
      uiManager.addExtractButton();
    }
  });
  
  observer.observe(document.body, { childList: true, subtree: true });

  // --- Transcript Capture for Teams recap/Stream ---
  const TRANSCRIPT_DIV_ID = 'transcript-extractor-for-microsoft-stream-hidden-div-with-transcript';
  const TRANSCRIPT_WRAPPER_ID = 'teams-transcript-extractor-wrapper';
  const transcriptCaptureEnabled = (() => {
    try {
      const flag = window.localStorage?.getItem('teamsTranscriptCapture');
      if (flag === '1') return true;
      if (flag === '0') return false;
    } catch (_err) {
      // ignore storage errors
    }
    return false; // disabled by default
  })();

  const injectTranscriptFetchOverride = () => {
    if (window.__teamsTranscriptOverrideInjected) {
      return;
    }
    window.__teamsTranscriptOverrideInjected = true;

    const script = document.createElement('script');
    script.src = chrome.runtime.getURL('fetchOverride.js');
    script.dataset.teamsTranscriptOverride = 'true';
    script.onload = () => script.remove();
    (document.head || document.documentElement).appendChild(script);
  };

  const sanitizeFileName = (name) => name.replace(/[\\/:*?"<>|]+/g, '').trim();

  const getTranscriptFileName = () => {
    const heading = document.querySelector('h1[class*="videoTitle"] label') || document.querySelector('h1');
    const title = sanitizeFileName(heading?.innerText || document.title || '') || 'transcript';
    return `${title}.vtt`;
  };

  const getTranscriptText = () => {
    const transcriptContainer = document.getElementById(TRANSCRIPT_DIV_ID);
    if (!transcriptContainer) {
      return 'Transcript not ready yet. Start playback to load the transcript, then try again.';
    }

    const text = transcriptContainer.textContent || transcriptContainer.innerText || '';
    return text.trim() || 'Transcript not ready yet. Start playback to load the transcript, then try again.';
  };

  const downloadTranscript = (text) => {
    const element = document.createElement('a');
    element.setAttribute('href', 'data:text/vtt;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', getTranscriptFileName());
    element.style.display = 'none';
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
  };

  const attachTranscriptButtons = () => {
    if (document.getElementById(TRANSCRIPT_WRAPPER_ID)) {
      return;
    }

    const wrapper = document.createElement('div');
    wrapper.id = TRANSCRIPT_WRAPPER_ID;
    wrapper.style.position = 'fixed';
    wrapper.style.bottom = '16px';
    wrapper.style.right = '16px';
    wrapper.style.display = 'flex';
    wrapper.style.gap = '8px';
    wrapper.style.alignItems = 'center';
    wrapper.style.padding = '8px';
    wrapper.style.background = '#0b172a';
    wrapper.style.color = '#ffffff';
    wrapper.style.borderRadius = '8px';
    wrapper.style.boxShadow = '0 6px 18px rgba(0, 0, 0, 0.25)';
    wrapper.style.zIndex = '2147483647';
    wrapper.style.fontFamily = 'Segoe UI, system-ui, -apple-system, sans-serif';
    wrapper.style.fontSize = '12px';

    const makeButton = (label) => {
      const button = document.createElement('button');
      button.textContent = label;
      button.style.cursor = 'pointer';
      button.style.border = 'none';
      button.style.borderRadius = '6px';
      button.style.padding = '8px 10px';
      button.style.fontWeight = '600';
      button.style.fontSize = '12px';
      button.style.background = '#36cfc9';
      button.style.color = '#0b172a';
      button.style.boxShadow = '0 3px 8px rgba(0, 0, 0, 0.15)';
      button.style.transition = 'transform 120ms ease, box-shadow 120ms ease';
      button.onmouseenter = () => {
        button.style.transform = 'translateY(-1px)';
        button.style.boxShadow = '0 5px 12px rgba(0, 0, 0, 0.2)';
      };
      button.onmouseleave = () => {
        button.style.transform = '';
        button.style.boxShadow = '0 3px 8px rgba(0, 0, 0, 0.15)';
      };
      return button;
    };

    const copyBtn = makeButton('Copy Transcript');
    const downloadBtn = makeButton('Download .vtt');
    const closeBtn = makeButton('Close');
    closeBtn.style.background = '#ffb703';

    copyBtn.addEventListener('click', async () => {
      const textToCopy = getTranscriptText();
      try {
        await navigator.clipboard.writeText(textToCopy);
      } catch (err) {
        console.error('[Teams Transcript] Failed to copy transcript', err);
      }
    });

    downloadBtn.addEventListener('click', () => {
      const textToDownload = getTranscriptText();
      downloadTranscript(textToDownload);
    });

    closeBtn.addEventListener('click', () => {
      wrapper.remove();
    });

    wrapper.appendChild(copyBtn);
    wrapper.appendChild(downloadBtn);
    wrapper.appendChild(closeBtn);
    document.body.appendChild(wrapper);
  };

  if (transcriptCaptureEnabled) {
    injectTranscriptFetchOverride();
    attachTranscriptButtons();
  } else {
    console.log('[Teams Transcript] Disabled by default. Set localStorage "teamsTranscriptCapture" to "1" to enable.');
  }

  // --- Chat API capture (disabled by default) ---
  const chatApiModeEnabled = (() => {
    try {
      const flag = window.localStorage?.getItem('teamsChatApiMode');
      if (flag === '0') return false;
      if (flag === '1') return true;
    } catch (_err) {
      // ignore storage errors
    }
    return true; // enable by default
  })();

  const chatApiCaptureEnabled = (() => {
    try {
      const flag = window.localStorage?.getItem('teamsChatApiCapture');
      if (flag === '1') return true;
      if (flag === '0') return false;
    } catch (_err) {
      // ignore storage errors
    }
    return false || chatApiModeEnabled; // enable capture if API mode is on
  })();

  const injectChatFetchOverride = () => {
    if (window.__teamsChatOverrideInjected) {
      return;
    }
    window.__teamsChatOverrideInjected = true;

    const script = document.createElement('script');
    script.src = chrome.runtime.getURL('chatFetchOverride.js');
    script.dataset.teamsChatOverride = 'true';
    script.onload = () => script.remove();
    (document.head || document.documentElement).appendChild(script);
  };

  if (chatApiCaptureEnabled || chatApiModeEnabled) {
    injectChatFetchOverride();
  } else {
    console.log('[Teams Chat API capture] Disabled by default. Set localStorage "teamsChatApiCapture" to "1" to log/message API responses.');
  }

  if (chatApiModeEnabled) {
    console.log('[Teams Chat API capture] API-first mode enabled (will try API data before DOM).');
  }

  console.log('✅ Teams Chat Extractor - Improved Modular Version Initialized');
})();
