/**
 * Microsoft Teams Chat Extractor - Main Content Script
 * 
 * This script automates the process of extracting full chat conversations,
 * including the author, timestamp, and message content for each message.
 *
 * VERSION 6 - MODULAR ARCHITECTURE
 * 
 * This version has been refactored into a modular architecture with separate modules for:
 * - Teams detection and variant handling
 * - Checkbox management and UI state
 * - Message extraction and auto-scrolling
 * - Button and tooltip management
 */

import { TeamsDetection } from './src/modules/teamsDetection.js';
import { CheckboxManager } from './src/modules/checkboxManager.js';
import { ExtractionEngine } from './src/modules/extractionEngine.js';
import { ButtonManager } from './src/modules/buttonManager.js';

(async () => {
  console.log('🚀 Teams Chat Extractor - Modular Version Starting...');
  
  // Initialize modules
  const extractionEngine = new ExtractionEngine();
  const checkboxManager = new CheckboxManager();
  const buttonManager = new ButtonManager(extractionEngine, checkboxManager);
  
  // Initialize settings with defaults
  window.teamsExtractorSettings = { 
    showCheckboxes: false, 
    autoScroll: true 
  };

  // Inject checkbox styles
  checkboxManager.injectStylesheet();

  // --- Message Listener ---
  chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
    switch (request.action) {
      case 'toggleCheckboxes':
        window.teamsExtractorSettings.showCheckboxes = request.enabled;
        checkboxManager.settings.showCheckboxes = request.enabled;
        checkboxManager.toggleCheckboxes(request.enabled);
        
        if (request.enabled) {
          buttonManager.addExtractButton();
        } else {
          buttonManager.removeExtractButton();
        }
        break;
        
      case 'getState':
        sendResponse({
          checkboxesEnabled: window.teamsExtractorSettings.showCheckboxes,
          currentChat: TeamsDetection.getCurrentChatTitle()
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
        buttonManager.startExtraction();
        break;
        
      case 'getCurrentChat':
        const currentChat = TeamsDetection.getCurrentChatTitle();
        sendResponse({chatTitle: currentChat});
        break;
    }
    sendResponse({status: 'received'});
    return true; // Keep the message channel open for async response
  });

  // Load saved settings from Chrome storage
  chrome.storage.local.get(['checkboxesEnabled', 'autoScrollEnabled'], (result) => {
    if (result.checkboxesEnabled) {
      window.teamsExtractorSettings.showCheckboxes = true;
      checkboxManager.settings.showCheckboxes = true;
    }
    
    if (result.autoScrollEnabled !== undefined) {
      window.teamsExtractorSettings.autoScroll = result.autoScrollEnabled;
      extractionEngine.settings.autoScroll = result.autoScrollEnabled;
    }
    
    if (result.checkboxesEnabled) {
      // Wait a bit for Teams to load, then add UI elements
      setTimeout(() => {
        checkboxManager.addCheckboxes();
        buttonManager.addExtractButton();
      }, 2000);
    }
  });
  
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
      buttonManager.addExtractButton();
    }
  });
  
  observer.observe(document.body, { childList: true, subtree: true });

  console.log('✅ Teams Chat Extractor - Modular Version Initialized');
})();