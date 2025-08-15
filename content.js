/**
 * This script automates the process of extracting full chat conversations,
 * including the author, timestamp, and message content for each message.
 *
 * VERSION 5 - STABLE CHECKBOXES & SELECT ALL
 *
 * HOW IT WORKS:
 * 1. It injects styled checkboxes next to each chat list item.
 * 2. It listens for messages from the popup to select or deselect all conversations.
 * 3. It monitors for manual checkbox clicks to update the popup button.
 * 4. When "Extract Data" is clicked, it processes only the checked conversations.
 */
(async () => {
  // --- Configuration ---
  const DELAY_BETWEEN_CLICKS_MS = 2000; // 2 seconds. Increase if chats load slowly.

  // --- Selectors ---
  const CHAT_LIST_ITEM_SELECTOR = '[data-tid="chat-list-item"]';
  const CHAT_TITLE_SELECTOR = '[data-tid="chat-list-item-title"]';
  const MESSAGE_CONTAINER_SELECTOR = '[data-tid="chat-pane-item"]';
  const AUTHOR_SELECTOR = '[data-tid="message-author-name"]';
  const TIMESTAMP_SELECTOR = '.fui-ChatMessage__timestamp, .fui-ChatMyMessage__timestamp';
  const MESSAGE_BODY_SELECTOR = '.fui-ChatMessage__body, .fui-ChatMyMessage__body';

  // Additional selectors to try for chat list items
  // Try both Windows and Mac selectors
  const ALTERNATIVE_CHAT_SELECTORS = [
    '[data-tid="chat-list-item"]',  // Mac Teams (original working)
    '.fui-TreeItem',               // Windows Teams v2
    '[role="treeitem"]',           // Windows Teams v2 alternative
    '[data-testid="chat-list-item"]', 
    '.ms-List-cell',
    '[data-automationid*="chat"]',
    '[data-tid*="chat-item"]',
    '.chat-item',
    '.conversation-item',
    '[role="listitem"]',           // Generic fallback
    '.ms-FocusZone [role="button"]' // Another fallback
  ];

  /**
   * Injects a stylesheet to control the appearance of the checkboxes.
   */
  const injectStylesheet = () => {
    const style = document.createElement('style');
    style.textContent = `
      .conversation-checkbox {
        width: 20px;
        height: 20px;
        margin-right: 10px;
        cursor: pointer;
        vertical-align: middle;
        accent-color: #5B5FC5;
        flex-shrink: 0;
      }
      
      /* Ensure parent containers can accommodate checkboxes */
      [data-tid="chat-list-item"],
      .fui-TreeItem, 
      [role="treeitem"] {
        display: flex;
        align-items: center;
      }
      
      /* Specific adjustments for different platforms */
      [data-tid="chat-list-item"] .conversation-checkbox {
        margin-right: 15px;
      }
      
      .fui-TreeItem .conversation-checkbox,
      [role="treeitem"] .conversation-checkbox {
        margin-right: 8px;
        position: relative;
        z-index: 10;
      }
    `;
    document.head.appendChild(style);
  };

  /**
   * Helper function to pause execution.
   */
  const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

  /**
   * Extracts detailed information for all visible messages in the chat pane.
   */
  const extractVisibleMessages = () => {
    const messageContainers = document.querySelectorAll(MESSAGE_CONTAINER_SELECTOR);
    const extractedMessages = [];

    for (const container of messageContainers) {
      const author = container.querySelector(AUTHOR_SELECTOR)?.innerText.trim();
      const timestampElement = container.querySelector(TIMESTAMP_SELECTOR);
      const timestamp = timestampElement?.title || timestampElement?.innerText.trim();
      const message = container.querySelector(MESSAGE_BODY_SELECTOR)?.innerText.trim();
      const type = container.querySelector('.fui-ChatMyMessage__timestamp') ? 'sent' : 'received';

      if (author && timestamp && message) {
        extractedMessages.push({ author, timestamp, message, type });
      }
    }
    return extractedMessages;
  };

  /**
   * Notifies the popup of the current selection state.
   */
  const notifyPopupOfState = () => {
    const checkboxes = document.querySelectorAll('.conversation-checkbox');
    const selectedCount = Array.from(checkboxes).filter(c => c.checked).length;
    const allSelected = checkboxes.length > 0 && selectedCount === checkboxes.length;

    const extractButton = document.getElementById('extractSelectedButton');
    if (extractButton) {
      if (selectedCount === 0) {
        extractButton.disabled = true;
        extractButton.style.opacity = '0.5';
        extractButton.title = 'Select conversations to extract first';
      } else {
        extractButton.disabled = false;
        extractButton.style.opacity = '1';
        extractButton.title = '';
      }
    }

    chrome.runtime.sendMessage({action: 'updateSelectionState', allSelected: allSelected, selectedCount: selectedCount});
  };

  /**
   * Adds checkboxes to each conversation item.
   */
  const addCheckboxes = () => {
    if (!window.teamsExtractorSettings?.showCheckboxes) return;
    
    let chatListItems = [];
    
    // Try each selector until we find chat items
    for (const selector of ALTERNATIVE_CHAT_SELECTORS) {
      const items = document.querySelectorAll(selector);
      if (items.length > 0) {
        chatListItems = items;
        break;
      }
    }
    
    if (chatListItems.length === 0) {
      // Try to find any clickable list items in the chat area
      const fallbackItems = document.querySelectorAll('[role="listitem"], .ms-FocusZone [role="button"], [data-tid*="item"]');
      if (fallbackItems.length > 0) {
        chatListItems = fallbackItems;
      }
    }
    
    for (const item of chatListItems) {
      if (!item.querySelector('.conversation-checkbox')) {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'conversation-checkbox';
        checkbox.addEventListener('change', notifyPopupOfState);
        checkbox.addEventListener('click', (e) => e.stopPropagation());
        
        // Try to insert the checkbox at the beginning
        if (item.firstChild) {
          item.insertBefore(checkbox, item.firstChild);
        } else {
          item.appendChild(checkbox);
        }
      }
    }
    notifyPopupOfState();
  };

  /**
   * Removes all checkboxes from the page.
   */
  const removeCheckboxes = () => {
    const checkboxes = document.querySelectorAll('.conversation-checkbox');
    checkboxes.forEach(checkbox => checkbox.remove());
    notifyPopupOfState();
  };

  /**
   * Selects or deselects all checkboxes.
   */
  const setAllCheckboxes = (selected) => {
    const checkboxes = document.querySelectorAll('.conversation-checkbox');
    checkboxes.forEach(checkbox => checkbox.checked = selected);
    notifyPopupOfState();
  };

  // --- Main Extraction Logic ---
  const startExtraction = async () => {
    console.log("Starting detailed chat extraction...");

    const chatListItems = document.querySelectorAll(CHAT_LIST_ITEM_SELECTOR);
    const allConversations = {};

    for (const item of chatListItems) {
      const checkbox = item.querySelector('.conversation-checkbox');
      if (checkbox && checkbox.checked) {
        const conversationName = item.querySelector(CHAT_TITLE_SELECTOR)?.innerText.trim();

        if (conversationName) {
          console.log(`\n[->] Processing conversation with: "${conversationName}"`);
          item.click();
          await delay(DELAY_BETWEEN_CLICKS_MS);
          const messages = extractVisibleMessages();
          allConversations[conversationName] = messages;
          console.log(`    [✓] Extracted ${messages.length} detailed messages.`);
        }
      }
    }

    console.log("\n\n✅ --- EXTRACTION COMPLETE --- ✅");
    chrome.runtime.sendMessage({action: "download", data: allConversations});
  };

  // --- Message Listener ---
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    switch (request.action) {
      case 'toggleCheckboxes':
        window.teamsExtractorSettings.showCheckboxes = request.enabled;
        // Save state to Chrome storage
        chrome.storage.local.set({ checkboxesEnabled: request.enabled });
        if (request.enabled) {
          addCheckboxes();
          addExtractButton();
        } else {
          removeCheckboxes();
          removeExtractButton();
        }
        break;
      case 'getState':
        sendResponse({
          checkboxesEnabled: window.teamsExtractorSettings.showCheckboxes,
          currentChat: getCurrentChatTitle()
        });
        break;
      case 'selectAll':
        setAllCheckboxes(true);
        break;
      case 'deselectAll':
        setAllCheckboxes(false);
        break;
      case 'queryState':
        notifyPopupOfState();
        break;
      case 'extract':
        startExtraction();
        break;
      case 'getCurrentChat':
        const currentChat = getCurrentChatTitle();
        sendResponse({chatTitle: currentChat});
        break;
    }
    sendResponse({status: 'received'});
    return true; // Keep the message channel open for async response
  });

  /**
   * Gets the current chat title from the Teams interface
   */
  const getCurrentChatTitle = () => {
    const titleSelectors = [
      // Mac Teams selectors (original working ones first)
      '[data-tid="chat-header-title"]',
      '[data-tid="thread-header-title"]',
      '[data-tid="threadHeaderTitle"]',
      '[data-tid="chat-title"]',
      // Windows Teams v2 selectors
      '.fui-ThreadHeader__title',
      '[role="main"] h1',
      '.ms-Stack-inner h1',
      // Generic selectors (be more specific to avoid false matches)
      '.chat-header h1',
      '.ts-calling-thread-header',
      '[aria-label*="Chat with"]',
      'h1[data-tid*="title"]',
      '.thread-header-title',
      '[data-testid*="chat-title"]',
      '[data-testid*="thread-title"]'
    ];
    
    for (const selector of titleSelectors) {
      const element = document.querySelector(selector);
      if (element && element.textContent.trim()) {
        let title = element.textContent.trim();
        // Clean up common Teams UI text
        title = title.replace(/^\s*Chat\s*[-–]\s*/i, '');
        title = title.replace(/\s*\|\s*Microsoft Teams\s*$/i, '');
        if (title.length > 0) {
          return title;
        }
      }
    }
    
    // Fallback: try to get chat name from page title
    const pageTitle = document.title;
    if (pageTitle && pageTitle !== 'Microsoft Teams' && !pageTitle.includes('Teams')) {
      return pageTitle.replace(/\s*\|\s*Microsoft Teams\s*$/i, '');
    }
    
    return 'No chat selected';
  };

  /**
   * Adds the "Extract selected" button to the Teams page.
   */
  const addExtractButton = () => {
    // Only show button when checkboxes are enabled
    if (!window.teamsExtractorSettings?.showCheckboxes) {
      removeExtractButton();
      return;
    }

    const existingButton = document.getElementById('extractSelectedButton');
    if (existingButton) {
      return; // Button already exists
    }

    const extractButton = document.createElement('button');
    extractButton.id = 'extractSelectedButton';
    extractButton.textContent = 'Extract Selected';
    extractButton.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background-color: #5B5FC5;
      color: white;
      border: none;
      padding: 10px 16px;
      border-radius: 6px;
      cursor: pointer;
      z-index: 10000;
      font-size: 14px;
      font-weight: 600;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
      transition: all 0.2s ease;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    `;
    
    extractButton.addEventListener('click', () => {
      startExtraction();
    });
    
    extractButton.addEventListener('mouseenter', () => {
      extractButton.style.backgroundColor = '#4a4d9e';
      extractButton.style.transform = 'translateY(-1px)';
    });
    
    extractButton.addEventListener('mouseleave', () => {
      extractButton.style.backgroundColor = '#5B5FC5';
      extractButton.style.transform = 'translateY(0)';
    });

    document.body.appendChild(extractButton);
  };

  /**
   * Removes the extract button from the page.
   */
  const removeExtractButton = () => {
    const existingButton = document.getElementById('extractSelectedButton');
    if (existingButton) {
      existingButton.remove();
    }
  };

  // --- Initializer ---
  injectStylesheet();
  
  // Initialize settings
  window.teamsExtractorSettings = { showCheckboxes: false };
  
  // Load saved settings from Chrome storage
  chrome.storage.local.get(['checkboxesEnabled'], (result) => {
    if (result.checkboxesEnabled) {
      window.teamsExtractorSettings.showCheckboxes = true;
      // Wait a bit for Teams to load, then add UI elements
      setTimeout(() => {
        addCheckboxes();
        addExtractButton();
      }, 2000);
    }
  });
  
  const observer = new MutationObserver(() => {
    addCheckboxes();
    addExtractButton();
  });
  
  observer.observe(document.body, { childList: true, subtree: true });

})();