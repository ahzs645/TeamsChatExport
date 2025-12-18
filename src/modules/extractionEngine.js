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
   * Simple extraction of the currently active chat (no checkboxes needed)
   */
  async extractActiveChat() {
    console.log("🚀 Extracting active chat...");

    // Trigger bridge update
    await this.updateBridge();
    await this.delay(300); // Give bridge time to find everything

    const bridgeData = this.readFromBridge();
    let { token, conversationId, activeChatName } = bridgeData;

    // If no chat name found, use a default or try to get from conversation ID
    if (!activeChatName) {
      console.log("⚠️ Chat name not detected, using fallback name");
      // Try to get from selected item text
      const selected = document.querySelector('[aria-selected="true"]');
      if (selected) {
        const text = selected.textContent?.split('\n')[0]?.trim();
        if (text && text.length < 50) {
          activeChatName = text;
        }
      }
      // Still no name? Use timestamp
      if (!activeChatName) {
        activeChatName = `Chat_${new Date().toISOString().slice(0, 10)}`;
      }
    }

    console.log(`📋 Active chat: ${activeChatName}`);

    if (!conversationId) {
      console.error("❌ Could not find conversation ID for this chat.");
      console.log("💡 Try clicking on a different chat first, then come back.");
      return null;
    }

    console.log(`🔑 Conversation ID: ${conversationId.substring(0, 50)}...`);

    if (!token) {
      console.error("❌ No auth token available. Please refresh the page and try again.");
      return null;
    }

    console.log(`🎫 Token available (${token.length} chars)`);

    // Fetch messages via API
    console.log("🌐 Fetching messages via API...");
    const apiMessages = await this.apiMessageExtractor.fetchConversation(conversationId);

    if (apiMessages.length === 0) {
      console.log("⚠️ API returned no messages. Falling back to DOM extraction...");
      // Fall back to DOM extraction
      await this.scrollToTopOfChat();
      const visibleMessages = this.extractVisibleMessages();
      const combinedMessages = this.messageExtractor.prepareMessages(
        this.messageExtractor.mergeMessages(visibleMessages)
      );

      if (combinedMessages.length === 0) {
        console.error("❌ No messages found via API or DOM.");
        return null;
      }

      console.log(`✅ Extracted ${combinedMessages.length} messages via DOM`);
      return { [activeChatName]: combinedMessages };
    }

    const preparedMessages = this.messageExtractor.prepareMessages(
      this.messageExtractor.mergeMessages(apiMessages)
    );

    console.log(`✅ Extracted ${preparedMessages.length} messages via API`);

    // Embed avatars
    console.log("🖼️ Embedding avatars...");
    const messagesWithAvatars = await this.embedAvatars(preparedMessages);

    return { [activeChatName]: messagesWithAvatars };
  }

  /**
   * Main extraction logic (with checkbox selection)
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

        // Check if API-only mode is enabled (disables DOM fallback for testing)
        const apiOnlyMode = (() => {
          try {
            return window.localStorage?.getItem('teamsChatApiOnly') === '1';
          } catch (_e) { return false; }
        })();

        if (apiOnlyMode) {
          console.log('🔬 API-ONLY MODE: DOM fallback is DISABLED for testing');
        }

        // Trigger bridge update before looking for conversation ID
        await this.updateBridge();

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
              console.log('⚠️ API fetch returned no messages.');
              if (apiOnlyMode) {
                console.log('🔬 API-ONLY MODE: Skipping passive capture, no DOM fallback.');
              }
            }
          } else {
            console.log('⚠️ No conversation ID found!');
            console.log('💡 Try setting manually: localStorage.setItem("teamsChatConversationIdOverride", "19:xxx...")');
          }

          if (!usedApiCapture && !apiOnlyMode) {
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
        // Skip if API-only mode is enabled
        if (!usedApiCapture && this.getAutoScroll() && !apiOnlyMode) {
          await this.scrollToTopOfChat();
        }

        if (typeof window.__teamsExtractorMessageSequence !== 'number') {
          window.__teamsExtractorMessageSequence = 0;
        }

        // DOM fallback - skip if API-only mode is enabled
        if (!usedApiCapture && !apiOnlyMode) {
          const cachedMessages = this.collectCachedMessages();

          // Extract messages currently visible after scrolling
          const visibleMessages = this.extractVisibleMessages();

          combinedMessages = this.messageExtractor.prepareMessages(
            this.messageExtractor.mergeMessages([
              ...cachedMessages,
              ...visibleMessages
            ])
          );
        } else if (!usedApiCapture && apiOnlyMode) {
          console.log('🔬 API-ONLY MODE: No messages extracted. API did not return data.');
          console.log('📋 Debug info:');
          console.log('   - Token:', window.__teamsAuthToken ? 'Present' : 'Missing');
          console.log('   - Conversation ID:', this.getConversationId() || 'Not found');
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

  /**
   * Reads data from the context bridge element (set by contextBridge.js in page context)
   */
  readFromBridge() {
    const bridge = document.getElementById('teams-extractor-context-bridge');
    if (!bridge) {
      console.log('[Bridge] Bridge element not found');
      return { token: null, conversationId: null, activeChatName: null };
    }

    const token = bridge.getAttribute('data-token') || null;
    const conversationId = bridge.getAttribute('data-conversation-id') || null;
    const activeChatName = bridge.getAttribute('data-active-chat-name') || null;
    const timestamp = bridge.getAttribute('data-timestamp');

    console.log('[Bridge] Read from bridge:', {
      hasToken: !!token,
      tokenLength: token?.length || 0,
      activeChatName: activeChatName || 'null',
      conversationId: conversationId ? conversationId.substring(0, 40) + '...' : 'null',
      age: timestamp ? `${Date.now() - parseInt(timestamp)}ms ago` : 'unknown'
    });

    return { token, conversationId, activeChatName };
  }

  /**
   * Triggers the bridge to update (if available)
   */
  async updateBridge() {
    // Try to call the bridge update function in page context via a custom event
    const event = new CustomEvent('teams-extractor-update-bridge');
    document.dispatchEvent(event);
    // Give it a moment to update
    await new Promise(r => setTimeout(r, 100));
  }

  getConversationId() {
    // === FIRST: Try reading from context bridge ===
    const bridgeData = this.readFromBridge();
    if (bridgeData.conversationId) {
      console.log('[ConvID] ✅ Found in context bridge:', bridgeData.conversationId.substring(0, 40) + '...');
      return this.cleanConversationId(bridgeData.conversationId);
    }

    // Manual override via localStorage
    try {
      const override = window.localStorage?.getItem('teamsChatConversationIdOverride') ||
        window.localStorage?.getItem('teamsChatApiConversationId') ||
        window.localStorage?.getItem('teamsChatApiConversationIdOverride');
      if (override && override.includes('19:')) {
        console.log('[ConvID] Found override in localStorage:', override.substring(0, 40) + '...');
        return this.cleanConversationId(override);
      }
    } catch (_err) {
      // ignore storage errors
    }

    // Try location href (in case the SPA exposes it)
    const hrefMatch = (window.location?.href || '').match(/(19:[^/?#"]+)/);
    if (hrefMatch) {
      console.log('[ConvID] Found in URL:', hrefMatch[1].substring(0, 40) + '...');
      return this.cleanConversationId(hrefMatch[1]);
    }

    // === FLUENT UI DETECTION (New Teams) ===
    // Check for Fluent UI tree items with aria-selected="true"
    const fluentSelectors = [
      '.fui-TreeItem[aria-selected="true"]',
      '[data-fui-tree-item-value][aria-selected="true"]',
      '[role="treeitem"][aria-selected="true"]'
    ];

    for (const selector of fluentSelectors) {
      const activeItems = document.querySelectorAll(selector);
      for (const item of activeItems) {
        // Check data-fui-tree-item-value attribute (Fluent UI stores IDs here)
        const treeItemValue = item.getAttribute('data-fui-tree-item-value');
        if (treeItemValue && treeItemValue.includes('19:')) {
          console.log('[ConvID] Found in Fluent UI tree item value:', treeItemValue.substring(0, 40) + '...');
          return this.cleanConversationId(treeItemValue);
        }

        // Check all data-* attributes
        for (const [key, val] of Object.entries(item.dataset || {})) {
          if (val && val.includes('19:') && val.length > 25) {
            console.log(`[ConvID] Found in dataset.${key}:`, val.substring(0, 40) + '...');
            return this.cleanConversationId(val);
          }
        }

        // Check parent elements for the value (sometimes nested)
        let parent = item.parentElement;
        for (let i = 0; i < 3 && parent; i++) {
          const parentValue = parent.getAttribute('data-fui-tree-item-value');
          if (parentValue && parentValue.includes('19:')) {
            console.log('[ConvID] Found in parent tree item value:', parentValue.substring(0, 40) + '...');
            return this.cleanConversationId(parentValue);
          }
          parent = parent.parentElement;
        }
      }
    }

    // === CLASSIC TEAMS DETECTION ===
    const activeItem = document.querySelector('[data-tid="chat-list-item"][aria-selected="true"], [aria-selected="true"][data-qa-chat-id]');
    const idFromDataset = this.extractIdFromDataset(activeItem);
    if (idFromDataset) {
      console.log('[ConvID] Found in classic Teams dataset:', idFromDataset.substring(0, 40) + '...');
      return idFromDataset;
    }

    // Try any dataset fields on active item containing 19:
    if (activeItem) {
      const datasetValues = Object.values(activeItem.dataset || {});
      const match = datasetValues.find((v) => typeof v === 'string' && v.includes('19:'));
      if (match) {
        console.log('[ConvID] Found in active item dataset:', match.substring(0, 40) + '...');
        return this.cleanConversationId(match);
      }
    }

    // Try links containing conversation id
    const link = document.querySelector('a[href*="19:"]');
    if (link && link.href) {
      const match = link.href.match(/(19:[^/?#"]+)/);
      if (match) {
        console.log('[ConvID] Found in link href:', match[1].substring(0, 40) + '...');
        return this.cleanConversationId(match[1]);
      }
    }

    // Try recent performance entries for chatsvc calls
    const perfEntries = performance?.getEntriesByType?.('resource') || [];
    for (let i = perfEntries.length - 1; i >= 0; i--) {
      const name = perfEntries[i].name || '';
      if (name.includes('/conversations/') && name.includes('19:')) {
        const match = name.match(/(19:[^/?#]+)/);
        if (match) {
          console.log('[ConvID] Found in performance entry:', match[1].substring(0, 40) + '...');
          return this.cleanConversationId(match[1]);
        }
      }
      if (name.includes('/threads/') && name.includes('19:')) {
        const match = name.match(/(19:[^/?#]+)/);
        if (match) {
          console.log('[ConvID] Found in performance entry:', match[1].substring(0, 40) + '...');
          return this.cleanConversationId(match[1]);
        }
      }
    }

    console.log('[ConvID] ❌ No conversation ID found by any method');
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
