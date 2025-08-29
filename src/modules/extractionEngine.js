/**
 * Extraction Engine Module
 * Handles message extraction, auto-scrolling, and data processing
 */

import { TeamsDetection } from './teamsDetection.js';

export class ExtractionEngine {
  constructor() {
    this.settings = { autoScroll: true };
    this.DELAY_BETWEEN_CLICKS_MS = 2000; // 2 seconds. Increase if chats load slowly.
  }

  /**
   * Helper function to pause execution.
   */
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Scrolls to the top of the chat to ensure all messages are loaded
   */
  async scrollToTopOfChat() {
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
      await this.delay(2000); // 2 seconds for each load
      
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
    await this.delay(1000);
    
    const finalMessageCount = document.querySelectorAll('[data-tid="chat-pane-item"]').length;
    console.log(`📜 Scroll complete! Final count: ${finalMessageCount} messages (${scrollAttempts} scroll attempts)`);
  }

  /**
   * Converts relative timestamps to actual dates
   */
  convertRelativeTimestamp(timestamp) {
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
  }

  /**
   * Extracts detailed information for all visible messages in the chat pane.
   */
  extractVisibleMessages() {
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
        const convertedTimestamp = this.convertRelativeTimestamp(rawTimestamp);
        
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
  }

  /**
   * Main extraction logic
   */
  async startExtraction(selectedChatItems) {
    console.log("🚀 Starting detailed chat extraction...");

    const allConversations = {};
    console.log(`📝 Processing ${selectedChatItems.length} selected conversations`);

    for (const item of selectedChatItems) {
      // Get conversation name using multiple approaches
      let conversationName = TeamsDetection.getConversationName(item);

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
        
        await this.delay(this.DELAY_BETWEEN_CLICKS_MS);
        
        // Scroll to top to ensure we get all messages (if auto-scroll is enabled)
        if (this.settings.autoScroll) {
          await this.scrollToTopOfChat();
        }
        
        // Extract messages using New Teams selectors
        const messages = this.extractVisibleMessages();
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
  }

  /**
   * Updates auto-scroll setting
   */
  setAutoScroll(enabled) {
    this.settings.autoScroll = enabled;
    chrome.storage.local.set({ autoScrollEnabled: enabled });
  }

  /**
   * Gets auto-scroll setting
   */
  getAutoScroll() {
    return this.settings.autoScroll;
  }
}