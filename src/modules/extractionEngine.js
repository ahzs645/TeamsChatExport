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

    if (!conversationId) {
      console.error("Could not find conversation ID for this chat.");
      console.log("Try clicking on a different chat first, then come back.");
      return null;
    }

    console.log(`Conversation ID: ${conversationId.substring(0, 50)}...`);

    if (!token) {
      console.error("No auth token available. Please refresh the page and try again.");
      return null;
    }

    console.log(`Token available (${token.length} chars)`);

    console.log("Fetching messages via API...");
    const apiMessages = await this.apiMessageExtractor.fetchConversation(conversationId);

    if (apiMessages.length === 0) {
      console.error("No messages found via API.");
      return null;
    }

    const preparedMessages = this.prepareMessages(this.mergeMessages(apiMessages));
    console.log(`Extracted ${preparedMessages.length} messages via API`);

    console.log("Embedding avatars...");
    const messagesWithAvatars = await this.embedAvatars(preparedMessages);

    return { [activeChatName]: messagesWithAvatars };
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
