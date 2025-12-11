/**
 * Extraction Engine Module
 * Handles extraction coordination and orchestration between other modules
 */

import { TeamsVariantDetector } from './teamsVariantDetector.js';
import { MessageExtractor } from './messageExtractor.js';
import { ScrollManager } from './scrollManager.js';
import { ApiMessageExtractor } from './apiMessageExtractor.js';

export class ExtractionEngine {
  constructor() {
    this.messageExtractor = new MessageExtractor();
    this.scrollManager = new ScrollManager();
    this.apiMessageExtractor = new ApiMessageExtractor();
    this.DELAY_BETWEEN_CLICKS_MS = 2000; // 2 seconds. Increase if chats load slowly.
    this.embedAvatarsEnabled = true; // Enable avatar embedding by default
  }

  /**
   * Fetches an avatar image and converts it to base64
   */
  async fetchAvatarAsBase64(url) {
    if (!url || url.startsWith('data:')) {
      return url;
    }

    try {
      const response = await fetch(url, { credentials: 'include' });
      if (!response.ok) {
        console.warn(`[Avatar Fetch] HTTP ${response.status} for ${url.substring(0, 80)}...`);
        return null;
      }

      const buffer = await response.arrayBuffer();
      const bytes = new Uint8Array(buffer);
      let binary = '';
      for (let i = 0; i < bytes.length; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const base64 = btoa(binary);
      const contentType = response.headers.get('content-type') || 'image/png';
      return `data:${contentType};base64,${base64}`;
    } catch (err) {
      console.error(`[Avatar Fetch] Failed for ${url.substring(0, 80)}...`, err);
      return null;
    }
  }

  /**
   * Embeds avatars as base64 data URIs in messages
   */
  async embedAvatars(messages) {
    if (!this.embedAvatarsEnabled) {
      return messages;
    }

    // Collect unique avatar URLs
    const urlSet = new Set();
    for (const msg of messages) {
      if (msg.avatar && !msg.avatar.startsWith('data:')) {
        urlSet.add(msg.avatar);
      }
    }

    if (urlSet.size === 0) {
      return messages;
    }

    console.log(`🖼️ Fetching ${urlSet.size} unique avatar(s)...`);

    // Fetch all avatars and cache results
    const cache = new Map();
    for (const url of urlSet) {
      const base64 = await this.fetchAvatarAsBase64(url);
      cache.set(url, base64);
    }

    // Update messages with base64 avatars
    return messages.map(msg => {
      if (!msg.avatar || msg.avatar.startsWith('data:')) {
        return msg;
      }

      const base64Avatar = cache.get(msg.avatar);
      return {
        ...msg,
        avatarUrl: msg.avatar,  // Keep original URL
        avatar: base64Avatar || null
      };
    });
  }

  /**
   * Helper function to pause execution.
   */
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Gets auto-scroll setting
   */
  getAutoScroll() {
    return this.scrollManager.getAutoScroll();
  }

  /**
   * Sets auto-scroll setting
   */
  setAutoScroll(enabled) {
    this.scrollManager.setAutoScroll(enabled);
  }

  /**
   * Scrolls to the top of the chat to ensure all messages are loaded
   * Delegates to ScrollManager
   */
  async scrollToTopOfChat() {
    return await this.scrollManager.scrollToTopOfChat();
  }

  /**
   * Converts relative timestamps to actual dates
   * Delegates to MessageExtractor
   */
  convertRelativeTimestamp(timestamp) {
    return this.messageExtractor.convertRelativeTimestamp(timestamp);
  }

  /**
   * Extracts detailed information for all visible messages in the chat pane.
   * Delegates to MessageExtractor
   */
  extractVisibleMessages() {
    return this.messageExtractor.extractVisibleMessages();
  }

  /**
   * Main extraction logic
   */
  async startExtraction(selectedChatItems) {
    console.log("🚀 Starting detailed chat extraction...");

    const allConversations = {};
    console.log(`📝 Processing ${selectedChatItems.length} selected conversations`);

    for (let i = 0; i < selectedChatItems.length; i++) {
      const chatItem = selectedChatItems[i];
      console.log(`\n--- Processing conversation ${i + 1}/${selectedChatItems.length} ---`);

      const apiModeEnabled = this.apiMessageExtractor.isEnabled();
      if (apiModeEnabled) {
        this.apiMessageExtractor.startConversationWindow();
      }

      try {
        // Get the conversation name
        const conversationName = TeamsVariantDetector.getConversationName(chatItem);
        console.log(`📋 Conversation: ${conversationName}`);

        // Click on the conversation item to open it
        console.log("👆 Clicking on conversation...");
        chatItem.click();

        // Wait for the conversation to load
        await this.delay(this.DELAY_BETWEEN_CLICKS_MS);

        let combinedMessages = [];
        let usedApiCapture = false;

        if (apiModeEnabled) {
          const conversationId = this.getConversationId();
          if (conversationId) {
            console.log(`🌐 Attempting direct API fetch for conversation ${conversationId}`);
            const apiMessages = await this.apiMessageExtractor.fetchConversation(conversationId);
            if (apiMessages.length > 0) {
              combinedMessages = this.messageExtractor.prepareMessages(
                this.messageExtractor.mergeMessages(apiMessages)
              );
              usedApiCapture = true;
              console.log(`💾 Used API fetch for "${conversationName}" (${combinedMessages.length} messages, no scroll)`);
            } else {
              console.log('ℹ️ API fetch returned no messages; trying passive capture.');
            }
          } else {
            console.log('ℹ️ No conversation id found; skipping direct API fetch.');
          }

          if (!usedApiCapture) {
            const apiMessages = this.apiMessageExtractor.extractMessagesSince(this.apiMessageExtractor.conversationStartTs);
            if (apiMessages.length > 0) {
              combinedMessages = this.messageExtractor.prepareMessages(
                this.messageExtractor.mergeMessages(apiMessages)
              );
              usedApiCapture = true;
              console.log(`💾 Used passive API capture for "${conversationName}" (${combinedMessages.length} messages, no scroll)`);
            } else {
              console.log('ℹ️ API capture enabled but no messages captured; falling back to DOM scroll.');
            }
          }
        }

        // Auto-scroll to load all messages if enabled and API capture did not succeed
        if (!usedApiCapture && this.getAutoScroll()) {
          await this.scrollToTopOfChat();
        }

        if (typeof window.__teamsExtractorMessageSequence !== 'number') {
          window.__teamsExtractorMessageSequence = 0;
        }

        if (!usedApiCapture) {
          const cachedMessages = this.collectCachedMessages();

          // Extract messages currently visible after scrolling
          const visibleMessages = this.extractVisibleMessages();

          combinedMessages = this.messageExtractor.prepareMessages(
            this.messageExtractor.mergeMessages([
              ...cachedMessages,
              ...visibleMessages
            ])
          );
        }

        if (combinedMessages.length > 0) {
          // Embed avatars as base64
          console.log(`🖼️ Embedding avatars for "${conversationName}"...`);
          const messagesWithAvatars = await this.embedAvatars(combinedMessages);

          allConversations[conversationName] = messagesWithAvatars;
          const earliestMessage = messagesWithAvatars.find((msg) => msg.isoTimestamp) || messagesWithAvatars[0];
          const earliestTimestamp = earliestMessage?.timestamp || 'Unknown';
          console.log(`✅ Extracted ${messagesWithAvatars.length} messages from "${conversationName}" (earliest: ${earliestTimestamp})`);
        } else {
          console.log(`⚠️ No messages found in "${conversationName}"`);
        }

        window.__teamsExtractorMessageCache = undefined;
        window.__teamsExtractorMessageSequence = undefined;

      } catch (error) {
        console.error(`❌ Error processing conversation ${i + 1}:`, error);
        continue;
      }
    }

    console.log(`\n🎉 Extraction complete! Processed ${Object.keys(allConversations).length} conversations.`);

    // Send data to results page
    try {
      const response = await chrome.runtime.sendMessage({
        action: 'openResults',
        data: allConversations
      });

      if (response && response.success) {
        console.log('✅ Results page opened successfully');
      } else {
        console.log('⚠️ Results page may not have opened properly');
      }
    } catch (error) {
      console.error('❌ Error opening results page:', error);
      alert("Extension context was invalidated. Please refresh the page and try again.");
    }
  }

  collectCachedMessages() {
    const cache = window.__teamsExtractorMessageCache;
    if (!cache || !(cache instanceof Map)) {
      return [];
    }

    const entries = Array.from(cache.values()).sort((a, b) => (b.sequence ?? 0) - (a.sequence ?? 0));
    const parsedMessages = [];

    entries.forEach((value) => {
      if (!value || !value.html) return;
      const data = this.messageExtractor.extractFromHTML(value.html);
      if (data) {
        data.__sequence = value.sequence ?? (window.__teamsExtractorMessageSequence = (window.__teamsExtractorMessageSequence || 0) + 1);
        parsedMessages.push(data);
      }
    });

    return this.messageExtractor.prepareMessages(parsedMessages);
  }

  getConversationId() {
    // Manual override via localStorage
    try {
      const override = window.localStorage?.getItem('teamsChatConversationIdOverride') ||
        window.localStorage?.getItem('teamsChatApiConversationId') ||
        window.localStorage?.getItem('teamsChatApiConversationIdOverride');
      if (override && override.includes('19:')) {
        return this.cleanConversationId(override);
      }
    } catch (_err) {
      // ignore storage errors
    }

    // Try location href (in case the SPA exposes it)
    const hrefMatch = (window.location?.href || '').match(/(19:[^/?#"]+)/);
    if (hrefMatch) {
      return this.cleanConversationId(hrefMatch[1]);
    }

    // Try active chat list item
    const activeItem = document.querySelector('[data-tid="chat-list-item"][aria-selected="true"], [aria-selected="true"][data-qa-chat-id], [role="treeitem"][aria-selected="true"]');
    const idFromDataset = this.extractIdFromDataset(activeItem);
    if (idFromDataset) return idFromDataset;

    // Try any dataset fields on active item containing 19:
    if (activeItem) {
      const datasetValues = Object.values(activeItem.dataset || {});
      const match = datasetValues.find((v) => typeof v === 'string' && v.includes('19:'));
      if (match) return this.cleanConversationId(match);
    }

    // Try links containing conversation id
    const link = document.querySelector('a[href*="19:"]');
    if (link && link.href) {
      const match = link.href.match(/(19:[^/?#"]+)/);
      if (match) return this.cleanConversationId(match[1]);
    }

    // Try recent performance entries for chatsvc calls
    const perfEntries = performance?.getEntriesByType?.('resource') || [];
    for (let i = perfEntries.length - 1; i >= 0; i--) {
      const name = perfEntries[i].name || '';
      if (name.includes('/conversations/') && name.includes('19:')) {
        const match = name.match(/(19:[^/?#]+)/);
        if (match) {
          return this.cleanConversationId(match[1]);
        }
      }
      if (name.includes('/threads/') && name.includes('19:')) {
        const match = name.match(/(19:[^/?#]+)/);
        if (match) {
          return this.cleanConversationId(match[1]);
        }
      }
    }

    return null;
  }

  extractIdFromDataset(el) {
    if (!el || !el.dataset) return null;
    const candidates = ['qaChatId', 'chatId', 'conversationId', 'mid', 'id'];
    for (const key of candidates) {
      const value = el.dataset[key];
      if (value && value.includes('19:')) {
        return this.cleanConversationId(value);
      }
    }
    return null;
  }

  cleanConversationId(raw) {
    if (!raw) return null;
    const trimmed = raw.trim().replace(/["'<>]/g, '');
    const match = trimmed.match(/(19:[^#?\s"']+)/);
    if (match) return match[1];
    return trimmed;
  }
}
