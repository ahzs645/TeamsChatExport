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

  /**
   * Detects the Teams variant and returns appropriate selectors
   */
  const getTeamsVariantSelectors = () => {
    // Detect Teams variant
    const isNewTeams = document.querySelector('.fui-FluentProvider');
    const isClassicTeams = document.querySelector('.ts-app-root');
    const isTeamsV2 = document.querySelector('[data-app-name*="teams"]');
    
    if (isNewTeams) {
      // New Teams with Fluent UI
      return {
        chatItems: ['.fui-TreeItem', '[role="treeitem"]', '[data-testid="list-item"]'],
        chatTitles: ['.fui-ThreadHeader__title', '[role="main"] h1', '.ms-Stack-inner h1'],
        variant: 'new-teams'
      };
    } else if (isClassicTeams) {
      // Classic Teams
      return {
        chatItems: ['[data-tid="chat-list-item"]', '.ms-List-cell', '[role="listitem"]'],
        chatTitles: ['[data-tid="chat-header-title"]', '[data-tid="thread-header-title"]'],
        variant: 'classic-teams'
      };
    } else {
      // Generic/unknown - try everything
      return {
        chatItems: [
          '[data-tid="chat-list-item"]',  // Classic
          '.fui-TreeItem',               // New Teams
          '[role="treeitem"]',           // New Teams alternative
          '[data-testid="list-item"]',   // New Teams data-testid
          '.ms-List-cell',
          '[role="listitem"]'
        ],
        chatTitles: [
          '[data-tid="chat-header-title"]',
          '.fui-ThreadHeader__title',
          '[role="main"] h1'
        ],
        variant: 'unknown'
      };
    }
  };

  /**
   * Injects a stylesheet to control the appearance of the checkboxes.
   */
  const injectStylesheet = () => {
    const style = document.createElement('style');
    style.textContent = `
      .conversation-checkbox {
        width: 16px;
        height: 16px;
        cursor: pointer;
        accent-color: #5B5FC5;
        flex-shrink: 0;
        z-index: 1000;
        margin: 0;
        position: absolute;
      }
      
      /* Classic Teams support */
      [data-tid="chat-list-item"] {
        display: flex !important;
        align-items: center !important;
        padding-left: 25px !important;
        position: relative !important;
      }
      [data-tid="chat-list-item"] .conversation-checkbox {
        left: 5px;
        top: 50%;
        transform: translateY(-50%);
      }
      
      /* New Teams support - consistent positioning for all items */
      .fui-TreeItem {
        position: relative !important;
        padding-left: 25px !important;
      }
      
      .fui-TreeItem .conversation-checkbox {
        left: 5px;
        top: 50%;
        transform: translateY(-50%);
      }
      
      /* Ensure the tree container doesn't interfere */
      .fui-Tree {
        position: relative !important;
      }
      
      /* Make checkboxes more visible when debugging */
      .conversation-checkbox:hover {
        transform: translateY(-50%) scale(1.2);
        border: 2px solid #5B5FC5;
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
    const teamsConfig = getTeamsVariantSelectors();
    const extractedMessages = [];
    
    // Try multiple message container selectors for different Teams variants
    const messageSelectors = [
      '[data-tid="chat-pane-item"]',      // Classic/current
      '.fui-unstable-ChatItem',          // New Teams
      '.fui-ChatMessage',                // Alternative
      '[data-tid="message"]'             // Generic
    ];
    
    let messageContainers = [];
    let workingSelector = null;
    
    for (const selector of messageSelectors) {
      const containers = document.querySelectorAll(selector);
      if (containers.length > 0) {
        messageContainers = containers;
        workingSelector = selector;
        console.log(`📨 Using message selector "${selector}" - found ${containers.length} messages`);
        break;
      }
    }

    if (messageContainers.length === 0) {
      console.log('❌ No message containers found');
      return extractedMessages;
    }

    for (const container of messageContainers) {
      // Try multiple author selectors
      const authorSelectors = [
        '[data-tid="message-author-name"]',
        '.fui-ChatMessage__author',
        '.message-author',
        '[data-testid="message-author"]',
        'span[title*="@"]',
        '.author'
      ];
      
      let author = null;
      for (const sel of authorSelectors) {
        const el = container.querySelector(sel);
        if (el && el.textContent.trim()) {
          author = el.textContent.trim();
          break;
        }
      }
      
      // Try multiple timestamp selectors
      const timestampSelectors = [
        '.fui-ChatMessage__timestamp',
        '.fui-ChatMyMessage__timestamp', 
        '[data-tid="timestamp"]',
        '.timestamp',
        'time',
        '[title*=":"]'
      ];
      
      let timestamp = null;
      for (const sel of timestampSelectors) {
        const el = container.querySelector(sel);
        if (el) {
          timestamp = el.title || el.textContent.trim();
          if (timestamp) break;
        }
      }
      
      // Try multiple message body selectors
      const messageSelectors = [
        '.fui-ChatMessage__body',
        '.fui-ChatMyMessage__body',
        '[data-tid="message-body"]',
        '.message-body',
        '.message-content',
        'p', 'div[dir]'
      ];
      
      let message = null;
      for (const sel of messageSelectors) {
        const el = container.querySelector(sel);
        if (el && el.textContent.trim()) {
          message = el.textContent.trim();
          break;
        }
      }
      
      // Determine message type (sent vs received)
      let type = 'received';
      
      // Check multiple indicators for sent messages
      if (container.querySelector('.fui-ChatMyMessage__timestamp') ||
          container.querySelector('.fui-ChatMyMessage__body') ||
          container.classList.contains('sent') ||
          container.getAttribute('data-type') === 'sent' ||
          container.querySelector('[data-tid*="my-message"]') ||
          container.closest('[data-tid*="my-message"]') ||
          (author && author.toLowerCase().includes('ahmad'))) {
        type = 'sent';
      }
      
      // Additional check: if the message is aligned to the right or has specific styling
      const computedStyle = window.getComputedStyle(container);
      if (computedStyle.textAlign === 'right' || 
          computedStyle.justifyContent === 'flex-end' ||
          container.style.marginLeft === 'auto') {
        type = 'sent';
      }

      // Only include messages with at least author and message content
      if (message && message.length > 0) {
        // Use fallback values if author/timestamp missing
        const finalAuthor = author || 'Unknown User';
        const finalTimestamp = timestamp || new Date().toLocaleString();
        
        extractedMessages.push({ 
          author: finalAuthor, 
          timestamp: finalTimestamp, 
          message: message, 
          type: type 
        });
      }
    }
    
    console.log(`📝 Extracted ${extractedMessages.length} messages`);
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
    
    const teamsConfig = getTeamsVariantSelectors();
    let chatListItems = [];
    let workingSelector = null;
    
    console.log(`🔍 Teams variant detected: ${teamsConfig.variant}`);
    
    // Try selectors specific to detected Teams variant
    for (const selector of teamsConfig.chatItems) {
      const items = document.querySelectorAll(selector);
      if (items.length > 0) {
        chatListItems = items;
        workingSelector = selector;
        console.log(`✅ Using selector "${selector}" found ${items.length} items`);
        break;
      }
    }
    
    if (chatListItems.length === 0) {
      console.log('❌ No chat items found with variant-specific selectors');
      return;
    }
    
    // Filter out non-chat items (folders, etc.)
    const validChatItems = Array.from(chatListItems).filter(item => {
      // Skip conversation folders
      if (item.getAttribute('data-conversation-folder') === 'true') return false;
      if (item.getAttribute('data-item-type') === 'custom-folder') return false;
      
      // For New Teams, ensure it has a valid conversation ID
      const value = item.getAttribute('data-fui-tree-item-value');
      if (value && (value.includes('Conversation') || value.includes('Chat'))) return true;
      if (value && value.includes('BizChatMetaOSConversation')) return true; // Copilot chats
      
      // For Classic Teams, check data-tid
      const tid = item.getAttribute('data-tid');
      if (tid && tid.includes('chat-list-item')) return true;
      
      return true; // Default to include if unsure
    });
    
    console.log(`📋 Adding checkboxes to ${validChatItems.length} valid chat items`);
    
    for (const item of validChatItems) {
      if (!item.querySelector('.conversation-checkbox')) {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'conversation-checkbox';
        checkbox.addEventListener('change', notifyPopupOfState);
        checkbox.addEventListener('click', (e) => e.stopPropagation());
        
        // Insert checkbox at the beginning
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
    console.log("🚀 Starting detailed chat extraction...");

    const teamsConfig = getTeamsVariantSelectors();
    const allConversations = {};
    
    // Find all chat items (not just the old selector)
    let chatListItems = [];
    for (const selector of teamsConfig.chatItems) {
      const items = document.querySelectorAll(selector);
      if (items.length > 0) {
        chatListItems = items;
        console.log(`📋 Using "${selector}" - found ${items.length} chat items`);
        break;
      }
    }

    // Filter to only valid chat items with checkboxes
    const validItems = Array.from(chatListItems).filter(item => {
      const checkbox = item.querySelector('.conversation-checkbox');
      const isSelected = checkbox && checkbox.checked;
      
      if (isSelected) {
        console.log(`✅ Found selected item:`, item);
      }
      
      return isSelected;
    });

    console.log(`📝 Processing ${validItems.length} selected conversations`);

    for (const item of validItems) {
      // Get conversation name using multiple approaches
      let conversationName = getConversationName(item);

      if (conversationName) {
        console.log(`\n[->] Processing conversation: "${conversationName}"`);
        
        // Click the item to open the conversation
        item.click();
        await delay(DELAY_BETWEEN_CLICKS_MS);
        
        // Extract messages using New Teams selectors
        const messages = extractVisibleMessages();
        allConversations[conversationName] = messages;
        console.log(`    [✓] Extracted ${messages.length} detailed messages.`);
      } else {
        console.log(`⚠️ Skipping item - no conversation name found:`, item);
      }
    }

    console.log("\n\n✅ --- EXTRACTION COMPLETE --- ✅");
    console.log(`📊 Total conversations extracted: ${Object.keys(allConversations).length}`);
    
    chrome.runtime.sendMessage({action: "download", data: allConversations});
  };

  /**
   * Extracts conversation name from a chat item using multiple methods
   */
  const getConversationName = (item) => {
    // First try to get the fui-tree-item-value attribute which has the conversation info
    const treeValue = item.getAttribute('data-fui-tree-item-value');
    if (treeValue) {
      // Extract conversation ID for lookup (we'll use a different approach)
      console.log(`🔍 Tree item value: ${treeValue}`);
    }

    // Try various selectors for conversation names - be more specific
    const nameSelectors = [
      'span[id^="title-chat-list-item_"]',  // Fluent UI conversation titles
      '[data-tid="chat-list-item-title"]',  // Classic Teams
      '.fui-TreeItem-content .conversation-title',  // More specific New Teams
      '.conversation-title',
      'span[title]:not([title*=":"]):not([title*="AM"]):not([title*="PM"])',  // Exclude timestamps
      'div[title]:not([title*=":"]):not([title*="AM"]):not([title*="PM"])'   // Exclude timestamps
    ];
    
    for (const selector of nameSelectors) {
      const el = item.querySelector(selector);
      if (el && el.textContent.trim()) {
        let name = el.textContent.trim();
        // Clean up the name
        name = cleanConversationName(name);
        if (name && name.length > 0) {
          return name;
        }
      }
    }
    
    // More targeted fallback: look for specific patterns in New Teams
    const textContent = item.textContent.trim();
    if (textContent) {
      // Split by lines and look for the name pattern
      const lines = textContent.split('\n').filter(line => line.trim().length > 0);
      
      for (const line of lines) {
        const cleanLine = line.trim();
        
        // Look for patterns like "Name [ORGANIZATION]" or just "Name"
        // Stop at first timestamp or preview text
        if (cleanLine.match(/^[A-Za-z]+,\s*[A-Za-z]+\s*\[[A-Z]+\]/) || 
            cleanLine.match(/^[A-Za-z\s,.'()-]+$/)) {
          
          // Clean and return the name
          let name = cleanConversationName(cleanLine);
          if (name && name.length > 2) {
            return name;
          }
        }
        
        // Stop processing at first timestamp or message preview
        if (cleanLine.match(/\d{1,2}[:/]\d{2}/) || cleanLine.includes('You:')) {
          break;
        }
      }
    }
    
    return 'Unknown Conversation';
  };

  /**
   * Cleans up conversation names by removing timestamps, previews, etc.
   */
  const cleanConversationName = (text) => {
    if (!text) return null;
    
    let cleaned = text.trim();
    console.log(`🧹 Cleaning name: "${cleaned}"`);
    
    // First, try to extract name pattern "LastName, FirstName [ORG]" from the beginning
    // Improved regex to handle hyphens, apostrophes, and various name formats
    const nameMatch = cleaned.match(/^([A-Za-z-']+,\s*[A-Za-z\s-']+\s*\[[A-Z]+\])/);
    if (nameMatch) {
      console.log(`✅ Found name pattern: "${nameMatch[1]}"`);
      return nameMatch[1];
    }
    
    // Alternative: try to match just the name pattern without being too strict about what follows
    const relaxedNameMatch = cleaned.match(/([A-Za-z-']+,\s*[A-Za-z\s-']+\s*\[[A-Z]+\])/);
    if (relaxedNameMatch) {
      console.log(`✅ Found relaxed name pattern: "${relaxedNameMatch[1]}"`);
      return relaxedNameMatch[1];
    }
    
    // Alternative: look for just "FirstName LastName" pattern
    const simpleNameMatch = cleaned.match(/^([A-Za-z\s-]+(?:\s+[A-Za-z]+)*)/);
    if (simpleNameMatch) {
      const simpleName = simpleNameMatch[1].trim();
      // Stop at common separators
      const beforeTimestamp = simpleName.split(/\d{1,2}[:/]\d{2}/)[0].trim();
      const beforeMessage = beforeTimestamp.split(/\b(I\s|Hey\s|Hi\s|Hello\s|Thanks\s|You:)/i)[0].trim();
      
      if (beforeMessage.length > 2 && beforeMessage.length < 50) {
        console.log(`✅ Found simple name: "${beforeMessage}"`);
        return beforeMessage;
      }
    }
    
    // Fallback cleaning approach
    // Remove timestamp patterns
    cleaned = cleaned.replace(/\d{1,2}[:/]\d{2}\s*(AM|PM)?/gi, '');
    
    // Remove date patterns
    cleaned = cleaned.replace(/\d{1,2}\/\d{1,2}/g, '');
    cleaned = cleaned.replace(/\b(Yesterday|Today|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b/gi, '');
    
    // Remove "You:" and common message starters
    cleaned = cleaned.replace(/^You:\s*/gi, '');
    cleaned = cleaned.replace(/\b(I have|I will|I think|I got|Hey|Hi|Hello|Thanks|Thank you|Sure|Yes|No|OK|Okay|hmm|lol).*$/gi, '');
    
    // Clean up whitespace
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    
    console.log(`🧹 Final cleaned: "${cleaned}"`);
    
    // If reasonable length, return it
    if (cleaned.length >= 2 && cleaned.length <= 50) {
      return cleaned;
    }
    
    return null;
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
    const teamsConfig = getTeamsVariantSelectors();
    
    // Enhanced selectors for different Teams variants
    const enhancedSelectors = [
      // Fluent UI specific selectors for active/selected chat
      'span[id^="title-chat-list-item_"]',  // Fluent UI conversation titles in sidebar
      
      // New Teams specific selectors for chat header
      '[data-tid="thread-header-title"]',
      '.fui-ThreadHeader__title',
      'h1[role="heading"]',
      '[data-tid="chat-header-title"]',
      
      // Classic Teams selectors  
      '.ts-calling-thread-header',
      '.thread-header-title',
      
      // Generic header selectors
      'main h1',
      '[role="main"] h1',
      '.chat-header h1',
      '[aria-label*="Chat with"]',
      'h1[data-tid*="title"]',
      '[data-testid*="chat-title"]',
      '[data-testid*="thread-title"]',
      
      // Fallback selectors
      'header h1',
      '.header-title',
      '[title]:not([title*=":"]):not([title*="AM"]):not([title*="PM"])'
    ];
    
    // First try enhanced selectors
    for (const selector of enhancedSelectors) {
      const element = document.querySelector(selector);
      if (element && element.textContent.trim()) {
        let title = element.textContent.trim();
        
        // Clean up the title
        title = cleanConversationName(title);
        
        if (title && title.length > 0 && title !== 'Unknown Conversation') {
          console.log(`📋 Found current chat: "${title}" using selector: ${selector}`);
          return title;
        }
      }
    }
    
    // Try to get from currently selected/highlighted chat item
    const selectedChatItem = document.querySelector('.chat-list-item.active, .chat-list-item[aria-selected="true"], .fui-TreeItem[aria-selected="true"]');
    if (selectedChatItem) {
      const chatName = getConversationName(selectedChatItem);
      if (chatName && chatName !== 'Unknown Conversation') {
        console.log(`📋 Found current chat from selected item: "${chatName}"`);
        return chatName;
      }
    }
    
    // Fallback: try to get chat name from page title
    const pageTitle = document.title;
    if (pageTitle && pageTitle !== 'Microsoft Teams' && !pageTitle.includes('Teams')) {
      const cleanTitle = pageTitle.replace(/\s*\|\s*Microsoft Teams\s*$/i, '').trim();
      if (cleanTitle.length > 0) {
        console.log(`📋 Found current chat from page title: "${cleanTitle}"`);
        return cleanTitle;
      }
    }
    
    console.log('⚠️ No current chat detected');
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

    // Try to find the titlebar end slot container
    const titlebarEndSlot = document.querySelector('[data-tid="titlebar-end-slot"]');
    const settingsContainer = document.querySelector('[data-tid="more-options-header"]')?.closest('.fui-Primitive')?.parentElement;
    
    if (titlebarEndSlot && settingsContainer) {
      // Create button that matches Teams styling exactly
      const extractButton = document.createElement('button');
      extractButton.id = 'extractSelectedButton';
      extractButton.type = 'button';
      extractButton.textContent = 'Extract Selected';
      extractButton.className = 'fui-Button r1alrhcs ___1tbxa1d fhovq9v f1p3nwhy f11589ue f1q5o8ev f1pdflbu f11d4kpn f146ro5n f1s2uweq fr80ssc f1ukrpxl fecsdlb f139oj5f ft1hn21 fuxngvv fy5bs14 fsv2rcd f1h0usnq fs4ktlq f16h9ulv fx2bmrt f1omzyqd f1dfjoow f1j98vj9 fj8yq94 f4xjyn1 f1et0tmh f9ddjv3 f1wi8ngl f18ktai2 fwbmr0d f44c6la f10pi13n f1gl81tg fbmy5og fkam630 f12awlo f1tvt47b f1fig6q3 f1h878h';
      
      // Match the exact height and styling of titlebar buttons
      extractButton.style.cssText = `
        background-color: #5B5FC5 !important;
        color: white !important;
        margin-right: 8px;
        font-size: 12px;
        font-weight: 600;
        border-radius: 4px;
        padding: 6px 12px;
        border: none;
        cursor: pointer;
        transition: all 0.2s ease;
        height: 32px;
        display: flex;
        align-items: center;
        justify-content: center;
      `;
      
      extractButton.addEventListener('click', () => {
        startExtraction();
      });
      
      extractButton.addEventListener('mouseenter', () => {
        extractButton.style.backgroundColor = '#4a4d9e !important';
        extractButton.style.transform = 'translateY(-1px)';
      });
      
      extractButton.addEventListener('mouseleave', () => {
        extractButton.style.backgroundColor = '#5B5FC5 !important';
        extractButton.style.transform = 'translateY(0)';
      });

      // Wrap in the exact same container structure as other titlebar buttons
      const buttonContainer = document.createElement('div');
      buttonContainer.className = 'fui-Primitive ___bp1bit0 f22iagw f122n59 f1p9o1ba f1tqdzup fvnq6it f1u5se5m fc397y7 f1mfizis f6s8nf5 frqcrf8';
      
      const buttonWrapper = document.createElement('div');
      buttonWrapper.className = 'fui-Primitive';
      buttonWrapper.appendChild(extractButton);
      
      buttonContainer.appendChild(buttonWrapper);
      
      // Insert before the settings container within the titlebar-end-slot
      titlebarEndSlot.insertBefore(buttonContainer, settingsContainer);
    } else {
      // Fallback to fixed positioning if header container not found
      const extractButton = document.createElement('button');
      extractButton.id = 'extractSelectedButton';
      extractButton.textContent = 'Extract Selected';
      extractButton.style.cssText = `
        position: fixed;
        top: 20px;
        right: 140px;
        background-color: #5B5FC5;
        color: white;
        border: none;
        padding: 8px 14px;
        border-radius: 4px;
        cursor: pointer;
        z-index: 10000;
        font-size: 12px;
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
    }
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