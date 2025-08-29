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

  // --- Configuration and Detection ---

  /**
   * Detects the Teams variant and returns appropriate selectors
   */
  const getTeamsVariantSelectors = () => {
    // Detect Teams variant
    const isNewTeams = document.querySelector('.fui-FluentProvider');
    const isClassicTeams = document.querySelector('.ts-app-root');
    
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
   * Scrolls to the top of the chat to ensure all messages are loaded
   */
  const scrollToTopOfChat = async () => {
    console.log('📜 Starting aggressive scroll to load all message history...');
    
    // Try multiple selectors for the chat container, with the viewport first
    const chatContainerSelectors = [
      '[data-tid="message-pane-list-viewport"]',  // The main scrollable viewport
      '[data-tid="message-pane-list-runway"]',
      '#chat-pane-list',
      '.fui-Chat',
      '[data-testid="virtualized-chat-list"]',
      '.fui-unstable-Chat',
      '[role="log"]'
    ];
    
    let chatContainer = null;
    for (const selector of chatContainerSelectors) {
      chatContainer = document.querySelector(selector);
      if (chatContainer) {
        console.log(`📜 Found chat container with: ${selector}`);
        break;
      }
    }
    
    if (!chatContainer) {
      console.log('⚠️ No chat container found for scrolling');
      return;
    }
    
    // Get initial message count
    let previousMessageCount = document.querySelectorAll('[data-tid="chat-pane-item"]').length;
    console.log(`📜 Starting with ${previousMessageCount} messages`);
    
    // Aggressive scrolling to load all message history
    const maxScrollAttempts = 20; // Maximum scroll attempts to prevent infinite loops
    let scrollAttempts = 0;
    let noNewMessagesCount = 0;
    
    while (scrollAttempts < maxScrollAttempts && noNewMessagesCount < 3) {
      // Scroll to top
      chatContainer.scrollTop = 0;
      scrollAttempts++;
      
      console.log(`📜 Scroll attempt ${scrollAttempts}, waiting for messages to load...`);
      
      // Wait for messages to load
      await delay(2000); // 2 seconds for each load
      
      // Check if new messages loaded
      const currentMessageCount = document.querySelectorAll('[data-tid="chat-pane-item"]').length;
      
      if (currentMessageCount > previousMessageCount) {
        console.log(`📜 Progress: ${currentMessageCount} messages (+${currentMessageCount - previousMessageCount} new)`);
        previousMessageCount = currentMessageCount;
        noNewMessagesCount = 0; // Reset counter
      } else {
        console.log(`📜 No new messages loaded (still ${currentMessageCount} messages)`);
        noNewMessagesCount++;
      }
      
      // If we've had 3 attempts with no new messages, we're probably done
      if (noNewMessagesCount >= 3) {
        console.log('📜 No new messages for 3 attempts, assuming we have all history');
        break;
      }
    }
    
    // Final scroll to ensure we're at the very top
    chatContainer.scrollTop = 0;
    await delay(1000);
    
    const finalMessageCount = document.querySelectorAll('[data-tid="chat-pane-item"]').length;
    console.log(`📜 Scroll complete! Final count: ${finalMessageCount} messages (${scrollAttempts} scroll attempts)`);
  };

  /**
   * Converts relative timestamps to actual dates
   */
  const convertRelativeTimestamp = (timestamp) => {
    if (!timestamp || timestamp.trim() === '') return timestamp;
    
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);
    
    let cleanTimestamp = timestamp.trim();
    
    // Handle "Today" timestamps
    if (cleanTimestamp.toLowerCase().startsWith('today')) {
      const timeMatch = cleanTimestamp.match(/(\d{1,2}:\d{2}(?:\s*[ap]m)?)/i);
      if (timeMatch) {
        const timeStr = timeMatch[1];
        const dateStr = today.toLocaleDateString('en-US', { 
          year: 'numeric', 
          month: 'long', 
          day: 'numeric' 
        });
        return `${dateStr} ${timeStr}`;
      }
      return today.toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      });
    }
    
    // Handle "Yesterday" timestamps
    if (cleanTimestamp.toLowerCase().startsWith('yesterday')) {
      const timeMatch = cleanTimestamp.match(/(\d{1,2}:\d{2}(?:\s*[ap]m)?)/i);
      if (timeMatch) {
        const timeStr = timeMatch[1];
        const dateStr = yesterday.toLocaleDateString('en-US', { 
          year: 'numeric', 
          month: 'long', 
          day: 'numeric' 
        });
        return `${dateStr} ${timeStr}`;
      }
      return yesterday.toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      });
    }
    
    // Handle day names (Monday, Tuesday, etc.)
    const dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    const lowerTimestamp = cleanTimestamp.toLowerCase();
    
    for (let i = 0; i < dayNames.length; i++) {
      if (lowerTimestamp.startsWith(dayNames[i])) {
        const timeMatch = cleanTimestamp.match(/(\d{1,2}:\d{2}(?:\s*[ap]m)?)/i);
        const dayOfWeek = i; // 0 = Sunday, 1 = Monday, etc.
        
        // Find the most recent occurrence of this day
        const currentDayOfWeek = now.getDay();
        let daysAgo = (currentDayOfWeek - dayOfWeek + 7) % 7;
        if (daysAgo === 0) daysAgo = 7; // If it's the same day, assume it was last week
        
        const targetDate = new Date(now.getTime() - daysAgo * 24 * 60 * 60 * 1000);
        const dateStr = targetDate.toLocaleDateString('en-US', { 
          year: 'numeric', 
          month: '2-digit', 
          day: '2-digit' 
        });
        
        if (timeMatch) {
          const timeStr = timeMatch[1];
          return `${dateStr} ${timeStr}`;
        }
        return dateStr;
      }
    }
    
    // Handle "X days ago" format
    const daysAgoMatch = cleanTimestamp.match(/(\d+)\s*days?\s*ago/i);
    if (daysAgoMatch) {
      const daysAgo = parseInt(daysAgoMatch[1]);
      const targetDate = new Date(now.getTime() - daysAgo * 24 * 60 * 60 * 1000);
      const timeMatch = cleanTimestamp.match(/(\d{1,2}:\d{2}(?:\s*[ap]m)?)/i);
      
      const dateStr = targetDate.toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: '2-digit', 
        day: '2-digit' 
      });
      
      if (timeMatch) {
        const timeStr = timeMatch[1];
        return `${dateStr} ${timeStr}`;
      }
      return dateStr;
    }
    
    // Handle "X hours ago" format
    const hoursAgoMatch = cleanTimestamp.match(/(\d+)\s*hours?\s*ago/i);
    if (hoursAgoMatch) {
      const hoursAgo = parseInt(hoursAgoMatch[1]);
      const targetDate = new Date(now.getTime() - hoursAgo * 60 * 60 * 1000);
      return targetDate.toLocaleString('en-US', { 
        year: 'numeric', 
        month: '2-digit', 
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
      });
    }
    
    // Handle "X minutes ago" format
    const minutesAgoMatch = cleanTimestamp.match(/(\d+)\s*min(?:utes?)?\s*ago/i);
    if (minutesAgoMatch) {
      const minutesAgo = parseInt(minutesAgoMatch[1]);
      const targetDate = new Date(now.getTime() - minutesAgo * 60 * 1000);
      return targetDate.toLocaleString('en-US', { 
        year: 'numeric', 
        month: '2-digit', 
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
      });
    }
    
    // Handle formats like "12/25/2023" or "Dec 25" with optional time
    if (cleanTimestamp.match(/\d{1,2}\/\d{1,2}\/?\d{0,4}|\w{3}\s+\d{1,2}/)) {
      // Already has proper date format, just return it
      return timestamp;
    }
    
    // If no conversion needed, return original
    return timestamp;
  };

  /**
   * Extracts detailed information for all visible messages in the chat pane.
   */
  const extractVisibleMessages = () => {
    const extractedMessages = [];
    
    // Try multiple message container selectors for different Teams variants
    const messageSelectors = [
      '[data-tid="chat-pane-item"]',      // Classic/current
      '.fui-unstable-ChatItem',          // New Teams
      '.fui-ChatMessage',                // Fluent UI received messages
      '.fui-ChatMyMessage',              // Fluent UI sent messages
      '[data-tid="message"]',            // Generic
      '[role="article"]',                // ARIA role for messages
      '.message-container'               // Generic message container
    ];
    
    let messageContainers = [];
    
    for (const selector of messageSelectors) {
      const containers = document.querySelectorAll(selector);
      if (containers.length > 0) {
        messageContainers = containers;
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
      // Primary indicators for Fluent UI
      if (container.querySelector('.fui-ChatMyMessage__timestamp') ||
          container.querySelector('.fui-ChatMyMessage__body') ||
          container.classList.contains('fui-ChatMyMessage')) {
        type = 'sent';
      }
      
      // Secondary indicators - data attributes and classes
      if (container.classList.contains('sent') ||
          container.getAttribute('data-type') === 'sent' ||
          container.querySelector('[data-tid*="my-message"]') ||
          container.closest('[data-tid*="my-message"]') ||
          container.getAttribute('data-is-from-me') === 'true' ||
          container.classList.contains('me') ||
          container.classList.contains('self')) {
        type = 'sent';
      }
      
      // Additional check: if the message is aligned to the right or has specific styling
      const computedStyle = window.getComputedStyle(container);
      if (computedStyle.textAlign === 'right' || 
          computedStyle.justifyContent === 'flex-end' ||
          computedStyle.alignSelf === 'flex-end' ||
          computedStyle.marginLeft === 'auto' ||
          container.style.marginLeft === 'auto') {
        type = 'sent';
      }
      
      // Check parent container for sent message indicators
      const parentContainer = container.parentElement;
      if (parentContainer) {
        if (parentContainer.classList.contains('sent') ||
            parentContainer.classList.contains('me') ||
            parentContainer.getAttribute('data-from-me') === 'true') {
          type = 'sent';
        }
      }

      // Only include messages with at least author and message content
      if (message && message.length > 0) {
        // Use fallback values if author/timestamp missing
        const finalAuthor = author || 'Unknown User';
        const rawTimestamp = timestamp || new Date().toLocaleString();
        const convertedTimestamp = convertRelativeTimestamp(rawTimestamp);
        
        console.log(`🕐 Timestamp conversion: "${rawTimestamp}" → "${convertedTimestamp}"`);
        
        extractedMessages.push({ 
          author: finalAuthor, 
          timestamp: convertedTimestamp, 
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

    try {
      if (chrome.runtime && chrome.runtime.sendMessage) {
        chrome.runtime.sendMessage({action: 'updateSelectionState', allSelected: allSelected, selectedCount: selectedCount});
      }
    } catch (error) {
      // Silently ignore extension context issues
    }
  };

  /**
   * Adds checkboxes to each conversation item.
   */
  const addCheckboxes = () => {
    if (!window.teamsExtractorSettings?.showCheckboxes) return;
    
    const teamsConfig = getTeamsVariantSelectors();
    let chatListItems = [];
    
    // Only log once to avoid spam
    if (!window.teamsExtractorLogged) {
      console.log(`🔍 Teams variant detected: ${teamsConfig.variant}`);
      window.teamsExtractorLogged = true;
    }
    
    // Try selectors specific to detected Teams variant
    for (const selector of teamsConfig.chatItems) {
      const items = document.querySelectorAll(selector);
      if (items.length > 0) {
        chatListItems = items;
        // Only log once to avoid spam
        if (!window.teamsExtractorLogged) {
          console.log(`✅ Using selector "${selector}" found ${items.length} items`);
        }
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
    
    // Only log once to avoid spam
    if (!window.teamsExtractorLogged) {
      console.log(`📋 Adding checkboxes to ${validChatItems.length} valid chat items`);
    }
    
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
        
        // Try multiple click strategies for Fluent UI
        let clicked = false;
        
        // Strategy 1: Look for the main clickable container (multiple selectors for different layouts)
        const containerSelectors = [
          '.fui-TreeItemLayout',
          '[data-inp*="chat-switch"]',
          '[data-inp*="collab-unified-chat-switch"]',
          '.fui-TreeItem-content',
          '[role="treeitem"] > div'
        ];
        
        for (const selector of containerSelectors) {
          const clickableContainer = item.querySelector(selector);
          if (clickableContainer && !clicked) {
            console.log(`🖱️ Clicking Fluent UI container with selector: ${selector}`);
            clickableContainer.click();
            clicked = true;
            break;
          }
        }
        
        // Strategy 2: Try clicking the item itself if container click failed
        if (!clicked) {
          console.log('🖱️ Clicking tree item directly');
          item.click();
          clicked = true;
        }
        
        await delay(DELAY_BETWEEN_CLICKS_MS);
        
        // Scroll to top to ensure we get all messages (if auto-scroll is enabled)
        if (window.teamsExtractorSettings.autoScroll) {
          await scrollToTopOfChat();
        }
        
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
    
    try {
      if (chrome.runtime && chrome.runtime.sendMessage) {
        chrome.runtime.sendMessage({action: "download", data: allConversations});
      } else {
        console.error("Extension context invalidated. Please refresh the page and try again.");
        alert("Extension context was invalidated. Please refresh the page and try again.");
      }
    } catch (error) {
      console.error("Extension context invalidated. Please refresh the page and try again.", error);
      alert("Extension context was invalidated. Please refresh the page and try again.");
    }
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
  chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
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
    console.log('🔍 Detecting current chat...');
    
    // Method 1: Try to find the currently selected/active chat item in the sidebar
    const activeChatSelectors = [
      // Fluent UI active chat items
      '.fui-TreeItem[aria-selected="true"]',
      '.fui-TreeItem[data-is-selected="true"]', 
      '.fui-TreeItem.is-selected',
      
      // Classic Teams active items
      '[data-tid="chat-list-item"][aria-selected="true"]',
      '[data-tid="chat-list-item"].active',
      '.chat-list-item.active'
    ];
    
    for (const selector of activeChatSelectors) {
      const activeItem = document.querySelector(selector);
      if (activeItem) {
        const chatName = getConversationName(activeItem);
        if (chatName && chatName !== 'Unknown Conversation') {
          console.log(`📋 Found current chat from active item: "${chatName}" (selector: ${selector})`);
          return chatName;
        }
      }
    }
    
    // Method 2: Try to find the main chat header title (what's currently open)
    const headerSelectors = [
      // Main chat header selectors for the current conversation
      '[data-tid="thread-header-title"] span',
      '[data-tid="thread-header-title"]',
      '.fui-ThreadHeader__title',
      'main [role="heading"]',
      '[data-tid="chat-header-title"]',
      
      // Additional header locations
      'main h1',
      '[role="main"] h1',
      '.chat-header h1'
    ];
    
    for (const selector of headerSelectors) {
      const element = document.querySelector(selector);
      if (element && element.textContent.trim()) {
        let title = element.textContent.trim();
        
        // Clean up the title
        title = cleanConversationName(title);
        
        if (title && title.length > 0 && title !== 'Unknown Conversation') {
          console.log(`📋 Found current chat from header: "${title}" (selector: ${selector})`);
          return title;
        }
      }
    }
    
    // Method 3: Look for selected conversations with checkboxes
    const selectedCheckboxes = document.querySelectorAll('.conversation-checkbox:checked');
    if (selectedCheckboxes.length === 1) {
      // If exactly one checkbox is selected, use that conversation
      const checkedItem = selectedCheckboxes[0].closest('.fui-TreeItem, [data-tid="chat-list-item"]');
      if (checkedItem) {
        const chatName = getConversationName(checkedItem);
        if (chatName && chatName !== 'Unknown Conversation') {
          console.log(`📋 Found current chat from single selected checkbox: "${chatName}"`);
          return chatName;
        }
      }
    } else if (selectedCheckboxes.length > 1) {
      // Multiple selections - show count and first conversation name
      const firstCheckedItem = selectedCheckboxes[0].closest('.fui-TreeItem, [data-tid="chat-list-item"]');
      if (firstCheckedItem) {
        const firstName = getConversationName(firstCheckedItem);
        if (firstName && firstName !== 'Unknown Conversation') {
          console.log(`📋 Found multiple selected chats (${selectedCheckboxes.length}), showing first: "${firstName}"`);
          return `${selectedCheckboxes.length} chats selected (${firstName}...)`;
        }
      }
      console.log(`📋 Found ${selectedCheckboxes.length} selected chats`);
      return `${selectedCheckboxes.length} conversations selected`;
    }
    
    // Method 4: Try to get from page title as last resort
    const pageTitle = document.title;
    if (pageTitle && pageTitle !== 'Microsoft Teams' && !pageTitle.includes('Teams')) {
      const cleanTitle = pageTitle.replace(/\s*\|\s*Microsoft Teams\s*$/i, '').trim();
      if (cleanTitle.length > 0 && cleanTitle !== 'Microsoft Teams') {
        console.log(`📋 Found current chat from page title: "${cleanTitle}"`);
        return cleanTitle;
      }
    }
    
    console.log('⚠️ No current chat detected - showing default message');
    return 'No chat selected';
  };

  /**
   * Adds the "Extract selected" button and auto-scroll toggle to the Teams page.
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

    // Add a small delay to allow Teams to fully load
    setTimeout(() => {
      addExtractButtonInternal();
    }, 1000);
  };

  const addExtractButtonInternal = () => {
    const existingButton = document.getElementById('extractSelectedButton');
    if (existingButton) {
      return; // Button already exists
    }

    // Debug: Find the correct titlebar location
    console.log('🔍 Debugging titlebar location...');
    
    // Try multiple selectors to find the titlebar
    const titlebarSelectors = [
      '.app-layout-area--title-bar .title-bar [data-tid="titlebar-end-slot"]',
      '.title-bar [data-tid="titlebar-end-slot"]',
      '[data-tid="titlebar-end-slot"]'
    ];
    
    let titlebarEndSlot = null;
    for (const selector of titlebarSelectors) {
      titlebarEndSlot = document.querySelector(selector);
      if (titlebarEndSlot) {
        console.log(`✅ Found titlebar with selector: ${selector}`);
        console.log('📍 Titlebar parent hierarchy:', titlebarEndSlot.parentElement?.className);
        break;
      } else {
        console.log(`❌ Not found with selector: ${selector}`);
      }
    }
    
    const settingsButton = document.querySelector('[data-tid="more-options-header"]');
    const settingsContainer = settingsButton?.closest('.fui-Primitive');
    console.log('🔧 Settings button found:', !!settingsButton);
    console.log('🔧 Settings container found:', !!settingsContainer);
    
    // Find the spacer div that comes after the settings container
    const spacerDiv = titlebarEndSlot?.querySelector('.fui-Primitive.___u5xge90');
    console.log('🔧 Spacer div found:', !!spacerDiv);
    
    if (titlebarEndSlot && settingsContainer && spacerDiv) {
      console.log('✅ All required elements found, proceeding with button creation...');

      // Create extract button container with hover toggle
      const extractButtonContainer = document.createElement('div');
      extractButtonContainer.id = 'extractButtonContainer';
      extractButtonContainer.style.cssText = `
        position: relative;
        display: inline-block;
      `;
      
      // Create extract button that matches Teams styling exactly
      const extractButton = document.createElement('button');
      extractButton.id = 'extractSelectedButton';
      extractButton.type = 'button';
      extractButton.className = 'fui-Button r1alrhcs ___1tbxa1d fhovq9v f1p3nwhy f11589ue f1q5o8ev f1pdflbu f11d4kpn f146ro5n f1s2uweq fr80ssc f1ukrpxl fecsdlb f139oj5f ft1hn21 fuxngvv fy5bs14 fsv2rcd f1h0usnq fs4ktlq f16h9ulv fx2bmrt f1omzyqd f1dfjoow f1j98vj9 fj8yq94 f4xjyn1 f1et0tmh f9ddjv3 f1wi8ngl f18ktai2 fwbmr0d f44c6la f10pi13n f1gl81tg fbmy5og fkam630 f12awlo f1tvt47b f1fig6q3 f1h878h';
      extractButton.setAttribute('aria-label', 'Extract selected conversations');
      extractButton.title = 'Extract selected conversations';
      
      // Create download icon
      const buttonIcon = document.createElement('span');
      buttonIcon.className = 'fui-Button__icon rywnvv2';
      buttonIcon.innerHTML = `
        <span class="___1gzszts f22iagw">
          <svg class="fui-Icon-filled ___1vjqft9 fjseox fez10in fg4l7m0" fill="currentColor" aria-hidden="true" width="1em" height="1em" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
            <path d="M10.5 2.75a.75.75 0 0 0-1.5 0v8.69L6.03 8.47a.75.75 0 0 0-1.06 1.06l4.5 4.5a.75.75 0 0 0 1.06 0l4.5-4.5a.75.75 0 0 0-1.06-1.06L11 11.44V2.75ZM2.5 13.5A.75.75 0 0 1 3.25 13h13.5a.75.75 0 0 1 0 1.5H3.25a.75.75 0 0 1-.75-.75Z" fill="currentColor"/>
          </svg>
          <svg class="fui-Icon-regular ___12fm75w f1w7gpdv fez10in fg4l7m0" fill="currentColor" aria-hidden="true" width="1em" height="1em" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
            <path d="M10.5 2.75a.75.75 0 0 0-1.5 0v8.69L6.03 8.47a.75.75 0 0 0-1.06 1.06l4.5 4.5a.75.75 0 0 0 1.06 0l4.5-4.5a.75.75 0 0 0-1.06-1.06L11 11.44V2.75ZM2.5 13.5A.75.75 0 0 1 3.25 13h13.5a.75.75 0 0 1 0 1.5H3.25a.75.75 0 0 1-.75-.75Z" fill="currentColor"/>
          </svg>
        </span>
      `;
      
      extractButton.appendChild(buttonIcon);
      
      // Match the exact styling of other titlebar icon buttons
      extractButton.style.cssText = `
        background-color: #5B5FC5 !important;
        color: white !important;
        border: none;
        cursor: pointer;
        transition: all 0.2s ease;
        border-radius: 4px;
        padding: 5px 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        min-width: 32px;
        height: 32px;
      `;
      
      // Create hover tooltip that looks like a Teams menu
      const autoScrollTooltip = document.createElement('div');
      autoScrollTooltip.id = 'autoScrollTooltip';
      autoScrollTooltip.className = 'fui-MenuPopover';
      autoScrollTooltip.style.cssText = `
        position: fixed;
        background: white !important;
        color: #424242 !important;
        padding: 8px !important;
        border-radius: 8px !important;
        font-size: 13px !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
        white-space: nowrap !important;
        opacity: 0;
        visibility: hidden;
        transition: opacity 0.2s ease, visibility 0.2s ease;
        z-index: 999999 !important;
        pointer-events: auto;
        cursor: pointer;
        box-shadow: 0 4px 16px rgba(0,0,0,0.16), 0 0 2px rgba(0,0,0,0.08) !important;
        border: 1px solid #e0e0e0 !important;
        min-width: 140px !important;
      `;
      
      // Add arrow to tooltip (pointing down since tooltip is above button)
      const tooltipArrow = document.createElement('div');
      tooltipArrow.style.cssText = `
        position: absolute;
        top: 100%;
        left: 50%;
        transform: translateX(-50%);
        width: 0;
        height: 0;
        border-left: 6px solid transparent;
        border-right: 6px solid transparent;
        border-top: 6px solid #333;
      `;
      autoScrollTooltip.appendChild(tooltipArrow);
      
      const updateTooltipText = () => {
        const isEnabled = window.teamsExtractorSettings.autoScroll;
        
        // Create a menu item that looks like Teams menu items
        const menuItem = document.createElement('div');
        menuItem.className = 'fui-MenuItem';
        menuItem.style.cssText = `
          display: flex !important;
          align-items: center !important;
          padding: 8px 12px !important;
          border-radius: 6px !important;
          cursor: pointer !important;
          transition: background-color 0.2s ease !important;
          font-size: 13px !important;
          color: #424242 !important;
        `;
        
        // Add hover effect
        menuItem.addEventListener('mouseenter', () => {
          menuItem.style.backgroundColor = '#f5f5f5';
        });
        menuItem.addEventListener('mouseleave', () => {
          menuItem.style.backgroundColor = 'transparent';
        });
        
        // Add icon and text
        menuItem.innerHTML = `
          <span style="margin-right: 8px; font-size: 14px;">
            ${isEnabled ? '✅' : '⏸️'}
          </span>
          <span>Auto-scroll: <strong>${isEnabled ? 'ON' : 'OFF'}</strong></span>
        `;
        
        // Clear tooltip and add the menu item
        autoScrollTooltip.innerHTML = '';
        autoScrollTooltip.appendChild(menuItem);
        
        // Add click handler to the menu item
        menuItem.addEventListener('click', (e) => {
          e.stopPropagation();
          window.teamsExtractorSettings.autoScroll = !window.teamsExtractorSettings.autoScroll;
          chrome.storage.local.set({ autoScrollEnabled: window.teamsExtractorSettings.autoScroll });
          console.log(`🔄 Auto-scroll ${window.teamsExtractorSettings.autoScroll ? 'enabled' : 'disabled'}`);
          updateTooltipText();
        });
      };
      
      updateTooltipText();
      
      // Tooltip click handler is now handled by the menu item
      
      // Button click handler
      extractButton.addEventListener('click', () => {
        startExtraction();
      });
      
      // Hover handlers with debugging
      extractButton.addEventListener('mouseenter', () => {
        console.log('🐭 Button mouseenter triggered');
        
        // Better hover styling for the button - keep white icon
        extractButton.style.setProperty('background-color', '#4a4d9e', 'important');
        extractButton.style.setProperty('color', 'white', 'important');
        extractButton.style.transform = 'translateY(-1px)';
        
        // Force icon to stay white
        const iconSvgs = extractButton.querySelectorAll('svg');
        iconSvgs.forEach(svg => {
          svg.style.setProperty('fill', 'white', 'important');
        });
        
        // Position tooltip relative to viewport - ensure it's below the titlebar
        const buttonRect = extractButton.getBoundingClientRect();
        autoScrollTooltip.style.position = 'fixed';
        autoScrollTooltip.style.top = (buttonRect.bottom + 10) + 'px'; // Below the button instead of above
        autoScrollTooltip.style.left = (buttonRect.left + buttonRect.width/2 - 100) + 'px'; // Center it (assuming 200px max width)
        autoScrollTooltip.style.transform = 'none';
        
        autoScrollTooltip.style.opacity = '1';
        autoScrollTooltip.style.visibility = 'visible';
        autoScrollTooltip.style.display = 'block';
        
        // Remove the debugging colors - use normal styling now
        
        console.log('🐭 Button rect:', buttonRect);
        console.log('🐭 Tooltip positioned at:', {
          top: autoScrollTooltip.style.top,
          left: autoScrollTooltip.style.left
        });
      });
      
      extractButton.addEventListener('mouseleave', () => {
        console.log('🐭 Button mouseleave triggered');
        extractButton.style.setProperty('background-color', '#5B5FC5', 'important');
        extractButton.style.setProperty('color', 'white', 'important');
        extractButton.style.transform = 'translateY(0)';
        
        // Ensure icon stays white
        const iconSvgs = extractButton.querySelectorAll('svg');
        iconSvgs.forEach(svg => {
          svg.style.setProperty('fill', 'white', 'important');
        });
        // Add a small delay to allow moving to tooltip
        setTimeout(() => {
          const tooltipHovered = autoScrollTooltip.matches(':hover');
          console.log('🐭 Tooltip hovered?', tooltipHovered);
          if (!tooltipHovered) {
            autoScrollTooltip.style.opacity = '0';
            autoScrollTooltip.style.visibility = 'hidden';
            console.log('🐭 Tooltip hidden');
          }
        }, 150);
      });
      
      // Keep tooltip visible when hovering over it
      autoScrollTooltip.addEventListener('mouseenter', () => {
        console.log('🐭 Tooltip mouseenter');
        autoScrollTooltip.style.opacity = '1';
        autoScrollTooltip.style.visibility = 'visible';
      });
      
      autoScrollTooltip.addEventListener('mouseleave', () => {
        console.log('🐭 Tooltip mouseleave');
        autoScrollTooltip.style.opacity = '0';
        autoScrollTooltip.style.visibility = 'hidden';
      });
      
      extractButtonContainer.appendChild(extractButton);
      extractButtonContainer.appendChild(autoScrollTooltip);

      // Create container structure for the button to match Teams structure
      const extractContainer = document.createElement('div');
      extractContainer.className = 'fui-Primitive ___bp1bit0 f22iagw f122n59 f1p9o1ba f1tqdzup fvnq6it f1u5se5m fc397y7 f1mfizis f6s8nf5 frqcrf8';
      
      const extractWrapper = document.createElement('div');
      extractWrapper.className = 'fui-Primitive';
      extractWrapper.appendChild(extractButtonContainer);
      extractContainer.appendChild(extractWrapper);
      
      // Find all direct children of titlebarEndSlot to understand the structure
      const children = Array.from(titlebarEndSlot.children);
      console.log('📍 Titlebar children count:', children.length);
      children.forEach((child, index) => {
        console.log(`📍 Child ${index}:`, child.className || child.tagName, child.id || '');
      });
      
      // Try to find the insertion point more carefully
      let insertionPoint = null;
      
      // Look for the profile avatar section (usually the last main section)
      for (let i = children.length - 1; i >= 0; i--) {
        const child = children[i];
        if (child.querySelector && child.querySelector('[data-tid*=\"avatar\"]')) {
          insertionPoint = child;
          console.log('📍 Found avatar section for insertion point');
          break;
        }
      }
      
      // If no avatar found, try to insert before the last non-empty div
      if (!insertionPoint) {
        for (let i = children.length - 1; i >= 0; i--) {
          const child = children[i];
          if (child.tagName === 'DIV' && child.children.length > 0) {
            insertionPoint = child;
            console.log('📍 Using last non-empty div as insertion point');
            break;
          }
        }
      }
      
      if (insertionPoint) {
        try {
          titlebarEndSlot.insertBefore(extractContainer, insertionPoint);
          console.log('✅ Button inserted successfully');
        } catch (error) {
          console.log('⚠️ Insert before failed, appending to end:', error.message);
          titlebarEndSlot.appendChild(extractContainer);
          console.log('✅ Button appended to end of titlebar');
        }
      } else {
        // Fallback: insert at the end
        titlebarEndSlot.appendChild(extractContainer);
        console.log('✅ Button appended to end of titlebar (no insertion point found)');
      }
    } else {
      console.log('❌ Required titlebar elements not found, using fallback positioning');
      // Fallback to fixed positioning if header container not found
      const controlsContainer = document.createElement('div');
      controlsContainer.id = 'extractControlsContainer';
      controlsContainer.style.cssText = `
        position: fixed;
        top: 20px;
        right: 140px;
        display: flex;
        align-items: center;
        z-index: 10000;
        gap: 8px;
      `;
      
      // Create auto-scroll toggle
      const autoScrollContainer = document.createElement('div');
      autoScrollContainer.style.cssText = `
        display: flex;
        align-items: center;
        background: rgba(255, 255, 255, 0.9);
        border-radius: 4px;
        padding: 4px 8px;
        font-size: 11px;
        color: #333;
        border: 1px solid #ccc;
      `;
      
      const autoScrollCheckbox = document.createElement('input');
      autoScrollCheckbox.type = 'checkbox';
      autoScrollCheckbox.id = 'autoScrollToggle';
      autoScrollCheckbox.checked = window.teamsExtractorSettings.autoScroll;
      autoScrollCheckbox.style.cssText = `
        margin-right: 4px;
        accent-color: #5B5FC5;
        cursor: pointer;
      `;
      
      const autoScrollLabel = document.createElement('label');
      autoScrollLabel.htmlFor = 'autoScrollToggle';
      autoScrollLabel.textContent = 'Auto-scroll';
      autoScrollLabel.style.cssText = `
        cursor: pointer;
        font-size: 11px;
        color: #333;
        font-weight: 500;
      `;
      
      autoScrollCheckbox.addEventListener('change', () => {
        window.teamsExtractorSettings.autoScroll = autoScrollCheckbox.checked;
        // Save setting to Chrome storage
        chrome.storage.local.set({ autoScrollEnabled: autoScrollCheckbox.checked });
        console.log(`🔄 Auto-scroll ${autoScrollCheckbox.checked ? 'enabled' : 'disabled'}`);
      });
      
      autoScrollContainer.appendChild(autoScrollCheckbox);
      autoScrollContainer.appendChild(autoScrollLabel);
      
      const extractButton = document.createElement('button');
      extractButton.id = 'extractSelectedButton';
      extractButton.textContent = 'Extract Selected';
      extractButton.style.cssText = `
        background-color: #5B5FC5;
        color: white;
        border: none;
        padding: 8px 14px;
        border-radius: 4px;
        cursor: pointer;
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
      
      controlsContainer.appendChild(autoScrollContainer);
      controlsContainer.appendChild(extractButton);
      document.body.appendChild(controlsContainer);
    }
  };

  /**
   * Removes the extract button and controls from the page.
   */
  const removeExtractButton = () => {
    const existingButton = document.getElementById('extractSelectedButton');
    const existingButtonContainer = document.getElementById('extractButtonContainer');
    const existingContainer = document.getElementById('extractControlsContainer');
    
    if (existingButtonContainer) {
      // Remove button container and its container hierarchy
      const buttonContainer = existingButtonContainer.closest('.fui-Primitive.___bp1bit0');
      if (buttonContainer) {
        buttonContainer.remove();
      } else {
        existingButtonContainer.remove();
      }
    } else if (existingButton) {
      // Fallback: remove button directly
      const buttonContainer = existingButton.closest('.fui-Primitive.___bp1bit0');
      if (buttonContainer) {
        buttonContainer.remove();
      } else {
        existingButton.remove();
      }
    }
    
    if (existingContainer) {
      existingContainer.remove();
    }
  };

  // --- Initializer ---
  injectStylesheet();
  
  // Initialize settings
  window.teamsExtractorSettings = { showCheckboxes: false, autoScroll: true };
  
  // Load saved settings from Chrome storage
  chrome.storage.local.get(['checkboxesEnabled', 'autoScrollEnabled'], (result) => {
    if (result.checkboxesEnabled) {
      window.teamsExtractorSettings.showCheckboxes = true;
    }
    if (result.autoScrollEnabled !== undefined) {
      window.teamsExtractorSettings.autoScroll = result.autoScrollEnabled;
    }
    
    if (result.checkboxesEnabled) {
      // Wait a bit for Teams to load, then add UI elements
      setTimeout(() => {
        addCheckboxes();
        addExtractButton();
      }, 2000);
    }
  });
  
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
      addCheckboxes();
      addExtractButton();
    }
  });
  
  observer.observe(document.body, { childList: true, subtree: true });

})();