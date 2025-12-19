/**
 * Extraction Engine Module
 * Simplified version for API-based extraction only
 */

import { ApiMessageExtractor } from './apiMessageExtractor.js';

export class ExtractionEngine {
  constructor() {
    this.apiMessageExtractor = new ApiMessageExtractor();
    this.embedAvatarsEnabled = true;
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

    const urlSet = new Set();
    for (const msg of messages) {
      if (msg.avatar && !msg.avatar.startsWith('data:')) {
        urlSet.add(msg.avatar);
      }
    }

    if (urlSet.size === 0) {
      return messages;
    }

    console.log(`Fetching ${urlSet.size} unique avatar(s)...`);

    const cache = new Map();
    for (const url of urlSet) {
      const base64 = await this.fetchAvatarAsBase64(url);
      cache.set(url, base64);
    }

    return messages.map(msg => {
      if (!msg.avatar || msg.avatar.startsWith('data:')) {
        return msg;
      }

      const base64Avatar = cache.get(msg.avatar);
      return {
        ...msg,
        avatarUrl: msg.avatar,
        avatar: base64Avatar || null
      };
    });
  }

  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Simple extraction of the currently active chat
   */
  async extractActiveChat() {
    console.log("Extracting active chat...");

    await this.updateBridge();
    await this.delay(300);

    const bridgeData = this.readFromBridge();
    let { token, conversationId, activeChatName } = bridgeData;

    if (!activeChatName) {
      const selected = document.querySelector('[aria-selected="true"]');
      if (selected) {
        const text = selected.textContent?.split('\n')[0]?.trim();
        if (text && text.length < 50) {
          activeChatName = text;
        }
      }
      if (!activeChatName) {
        activeChatName = `Chat_${new Date().toISOString().slice(0, 10)}`;
      }
    }

    console.log(`Active chat: ${activeChatName}`);

    let apiMessages = [];

    // Try API extraction first (works for v1)
    if (conversationId && token) {
      console.log(`Conversation ID: ${conversationId.substring(0, 50)}...`);
      console.log(`Token available (${token.length} chars)`);
      console.log("Trying API extraction...");
      apiMessages = await this.apiMessageExtractor.fetchConversation(conversationId);
    }

    // If API fails or returns nothing, try DOM extraction (works for v2)
    if (apiMessages.length === 0) {
      console.log("API returned no messages, trying DOM extraction (Teams v2 mode)...");

      // Scroll to load more messages first
      await this.scrollToLoadMessages();

      apiMessages = this.extractMessagesFromDOM();

      if (apiMessages.length > 0) {
        console.log(`Extracted ${apiMessages.length} messages via DOM`);
      }
    }

    if (apiMessages.length === 0) {
      console.error("No messages found via API or DOM.");
      return null;
    }

    const preparedMessages = this.prepareMessages(this.mergeMessages(apiMessages));
    console.log(`Prepared ${preparedMessages.length} messages`);

    console.log("Embedding avatars...");
    const messagesWithAvatars = await this.embedAvatars(preparedMessages);

    return { [activeChatName]: messagesWithAvatars };
  }

  /**
   * Scroll up in the chat to load more messages (for DOM extraction)
   */
  async scrollToLoadMessages() {
    const messageList = document.querySelector('[data-tid="message-pane-list-container"]') ||
                       document.querySelector('[role="main"] [data-virtualized]') ||
                       document.querySelector('#message-list');

    if (!messageList) {
      console.log("Could not find message container for scrolling");
      return;
    }

    console.log("Scrolling to load more messages...");
    const scrollAttempts = 10;

    for (let i = 0; i < scrollAttempts; i++) {
      messageList.scrollTop = 0;
      await this.delay(300);

      // Check if we've loaded more
      const currentCount = document.querySelectorAll('[data-tid="chat-pane-message"]').length;
      console.log(`Scroll ${i + 1}/${scrollAttempts}: ${currentCount} messages loaded`);
    }
  }

  /**
   * Extract messages directly from DOM (Teams v2 fallback)
   */
  extractMessagesFromDOM() {
    const messageElements = document.querySelectorAll('[data-tid="chat-pane-message"]');
    console.log(`Found ${messageElements.length} message elements in DOM`);

    const messages = [];

    messageElements.forEach((msg, index) => {
      try {
        const mid = msg.getAttribute('data-mid');

        // Try multiple methods to find author
        let author = 'Unknown';
        const authorEl = document.getElementById('author-' + mid) ||
                        msg.querySelector('[id*="author"]') ||
                        msg.querySelector('[data-tid*="author"]') ||
                        msg.querySelector('[class*="author"]') ||
                        msg.querySelector('[class*="sender"]');
        if (authorEl) {
          author = authorEl.textContent?.trim() || 'Unknown';
        }

        // Try multiple methods to find timestamp
        let timestamp = '';
        const timestampEl = document.getElementById('timestamp-' + mid) ||
                          msg.querySelector('[id*="timestamp"]') ||
                          msg.querySelector('time') ||
                          msg.querySelector('[data-tid*="time"]') ||
                          msg.querySelector('[class*="timestamp"]');
        if (timestampEl) {
          timestamp = timestampEl.getAttribute('datetime') || timestampEl.textContent?.trim() || '';
        }

        // Try multiple methods to find content
        let content = '';
        const contentEl = document.getElementById('content-' + mid) ||
                         msg.querySelector('[id*="content"]') ||
                         msg.querySelector('[data-tid*="message-text"]') ||
                         msg.querySelector('[class*="message-body"]') ||
                         msg.querySelector('[class*="text-content"]') ||
                         msg.querySelector('div[dir="auto"]');
        if (contentEl) {
          content = contentEl.textContent?.trim() || '';
        }

        // Extract attachments
        const attachments = [];
        const attachmentEls = msg.querySelectorAll('[id*="attachment"], [data-tid*="attachment"], [class*="attachment"]');
        attachmentEls.forEach(att => {
          const link = att.querySelector('a');
          attachments.push({
            type: 'attachment',
            name: att.textContent?.trim() || 'Attachment',
            href: link?.href || ''
          });
        });

        // Extract images
        const embeddedImages = [];
        const imgEls = msg.querySelectorAll('img:not([class*="emoji"]):not([class*="avatar"])');
        imgEls.forEach(img => {
          if (img.src && !img.src.includes('emoji') && !img.src.includes('avatar')) {
            embeddedImages.push({
              src: img.src,
              alt: img.alt || '',
              isEmoji: false
            });
          }
        });

        // Only add if there's content
        if (content || attachments.length > 0 || embeddedImages.length > 0) {
          messages.push({
            id: mid || `dom-${index}`,
            author,
            timestamp,
            isoTimestamp: this.parseTimestamp(timestamp),
            message: content,
            content,
            attachments,
            embeddedImages,
            reactions: [], // Could extract reactions too if needed
            type: null
          });
        }
      } catch (err) {
        console.warn('Error extracting message:', err);
      }
    });

    return messages;
  }

  /**
   * Parse various timestamp formats to ISO
   */
  parseTimestamp(timestamp) {
    if (!timestamp) return null;

    try {
      // Try direct parse first
      const date = new Date(timestamp);
      if (!isNaN(date.getTime())) {
        return date.toISOString();
      }

      // Handle "12/8 3:41 PM" format
      const match = timestamp.match(/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
      if (match) {
        const month = parseInt(match[1]) - 1;
        const day = parseInt(match[2]);
        let hours = parseInt(match[3]);
        const minutes = parseInt(match[4]);
        const ampm = match[5];

        if (ampm?.toUpperCase() === 'PM' && hours < 12) hours += 12;
        if (ampm?.toUpperCase() === 'AM' && hours === 12) hours = 0;

        const now = new Date();
        const date = new Date(now.getFullYear(), month, day, hours, minutes);
        return date.toISOString();
      }
    } catch (e) {
      // Ignore parse errors
    }

    return null;
  }

  /**
   * Merge messages and remove duplicates
   */
  mergeMessages(messages) {
    const seen = new Map();

    for (const msg of messages) {
      if (!msg) continue;
      const key = msg.id || `${msg.isoTimestamp || msg.timestamp}::${msg.author}::${msg.message?.substring(0, 50)}`;
      if (!seen.has(key)) {
        seen.set(key, msg);
      }
    }

    return Array.from(seen.values());
  }

  /**
   * Prepare and sort messages for display
   */
  prepareMessages(messages) {
    return messages
      .filter(msg => msg && (msg.message || msg.content || msg.attachments?.length))
      .sort((a, b) => {
        if (a.isoTimestamp && b.isoTimestamp) {
          return new Date(a.isoTimestamp) - new Date(b.isoTimestamp);
        }
        return (a.timestamp || '').localeCompare(b.timestamp || '');
      });
  }

  /**
   * Reads data from the context bridge element
   */
  readFromBridge() {
    const bridge = document.getElementById('teams-extractor-context-bridge');
    if (!bridge) {
      return { token: null, conversationId: null, activeChatName: null };
    }

    const token = bridge.getAttribute('data-token') || null;
    const conversationId = bridge.getAttribute('data-conversation-id') || null;
    const activeChatName = bridge.getAttribute('data-active-chat-name') || null;
    const timestamp = bridge.getAttribute('data-timestamp');

    console.log('[Bridge] Read:', {
      hasToken: !!token,
      tokenLength: token?.length || 0,
      activeChatName: activeChatName || 'null',
      conversationId: conversationId ? conversationId.substring(0, 40) + '...' : 'null',
      age: timestamp ? `${Date.now() - parseInt(timestamp)}ms ago` : 'unknown'
    });

    return { token, conversationId, activeChatName };
  }

  /**
   * Triggers the bridge to update
   */
  async updateBridge() {
    const event = new CustomEvent('teams-extractor-update-bridge');
    document.dispatchEvent(event);
    await new Promise(r => setTimeout(r, 100));
  }
}
