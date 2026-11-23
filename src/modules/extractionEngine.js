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
    const { token, conversationId } = bridgeData;
    let activeChatName = document.querySelector('[id^="chat-header-"] h2')?.textContent?.trim() ||
      bridgeData.activeChatName;

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

    let messages = [];

    // Try API extraction first when the active chat context is available.
    if (conversationId && token) {
      console.log("Trying API extraction...");
      messages = await this.apiMessageExtractor.fetchConversation(conversationId);
    } else {
      console.info('[Teams Chat API] Skipped API export', { hasChatId: !!conversationId, hasAuth: !!token });
    }

    // Use the chat pane when the API is unavailable or returns no messages.
    if (messages.length === 0) {
      console.info('[Teams Chat Extractor] Using chat pane fallback');

      messages = await this.scrollToLoadMessages();

      if (messages.length > 0) {
        console.log(`Extracted ${messages.length} messages via DOM`);
      }
    }

    if (messages.length === 0) {
      console.error("No messages found via API or DOM.");
      return null;
    }

    const preparedMessages = this.prepareMessages(this.mergeMessages(messages));
    console.log(`Prepared ${preparedMessages.length} messages`);

    console.log("Embedding avatars...");
    const messagesWithAvatars = await this.embedAvatars(preparedMessages);

    return { [activeChatName]: messagesWithAvatars };
  }

  /**
   * Collect each visible window before Teams recycles its message rows.
   */
  async scrollToLoadMessages() {
    const messageList = document.querySelector('[data-tid="message-pane-list-viewport"]') ||
                       document.querySelector('[data-tid="message-pane-list-container"]') ||
                       document.querySelector('[role="main"] [data-virtualized]') ||
                       document.querySelector('#message-list');
    if (!messageList) return this.extractMessagesFromDOM();

    const chatHeader = document.querySelector('[id^="chat-header-"]');
    const chatId = chatHeader?.id;
    const messages = new Map();
    const distanceFromBottom = messageList.scrollHeight - messageList.clientHeight - messageList.scrollTop;
    const collect = () => {
      if (!messageList.isConnected || (chatHeader && (!chatHeader.isConnected || chatHeader.id !== chatId))) {
        throw new Error('The open chat changed during extraction. Open the chat and try again.');
      }
      for (const message of this.extractMessagesFromDOM(messageList)) {
        const key = message.id || `${message.timestamp}::${message.author}::${message.message}`;
        if (!messages.has(key)) messages.set(key, message);
      }
    };

    try {
      collect();
      if (distanceFromBottom > 1) {
        messageList.scrollTop = messageList.scrollHeight;
        await this.delay(100);
        collect();
      }
      let historyLoads = 0;
      for (let step = 0; step < 200 && historyLoads < 10; step++) {
        const scrollStep = Math.max(1, Math.floor(messageList.clientHeight * 0.8));
        const target = Math.max(0, messageList.scrollTop - scrollStep);
        const heightBeforeScroll = messageList.scrollHeight;
        messageList.scrollTop = target;
        await this.delay(100);
        collect();
        if (target > 0) continue;

        // Reaching the top can trigger an asynchronous history request.
        historyLoads++;
        const countAtTop = messages.size;
        for (let attempt = 0; attempt < 6 && messageList.scrollTop <= 0 &&
             messageList.scrollHeight === heightBeforeScroll && messages.size === countAtTop; attempt++) {
          await this.delay(250);
          collect();
        }
        if (messageList.scrollTop > 0) continue;
        if (messageList.scrollHeight === heightBeforeScroll && messages.size === countAtTop) break;

        // Some layouts stay pinned at the top when older rows arrive.
        // Move a short distance so the next pass can reach the top again.
        messageList.scrollTop = Math.min(40, scrollStep);
        await this.delay(100);
        collect();
      }
    } finally {
      if (messageList.isConnected && (!chatHeader || (chatHeader.isConnected && chatHeader.id === chatId))) {
        messageList.scrollTop = messageList.scrollHeight - messageList.clientHeight - distanceFromBottom;
      }
    }
    return messages.size ? Array.from(messages.values()) : this.extractMessagesFromDOM();
  }

  /**
   * Extract messages directly from DOM (Teams v2 fallback)
   */
  extractMessagesFromDOM(root = document) {
    const messageElements = root.querySelectorAll('[data-tid="chat-pane-message"]');

    const messages = [];

    messageElements.forEach(msg => {
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
        const isoTimestamp = this.parseTimestamp(timestampEl?.getAttribute('datetime')) ||
          (/^\d{13}$/.test(mid) ? new Date(Number(mid)).toISOString() : this.parseTimestamp(timestamp));

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
            id: mid || null,
            author,
            timestamp,
            isoTimestamp,
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
