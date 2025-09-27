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

import { TeamsVariantDetector } from './src/modules/teamsVariantDetector.js';
import { CheckboxManager } from './src/modules/checkboxManager.js';
import { ExtractionEngine } from './src/modules/extractionEngine.js';
import { UIManager } from './src/modules/uiManager.js';
import { StorageManager } from './src/modules/storageManager.js';

(async () => {
  console.log('🚀 Teams Chat Extractor - Improved Modular Version Starting...');
  
  // Initialize modules
  const storageManager = new StorageManager();
  const extractionEngine = new ExtractionEngine();
  const checkboxManager = new CheckboxManager();
  const uiManager = new UIManager(extractionEngine, checkboxManager);
  
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

  console.log('✅ Teams Chat Extractor - Improved Modular Version Initialized');
})();