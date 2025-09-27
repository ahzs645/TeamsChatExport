/**
 * Extraction Engine Module
 * Handles extraction coordination and orchestration between other modules
 */

import { TeamsVariantDetector } from './teamsVariantDetector.js';
import { MessageExtractor } from './messageExtractor.js';
import { ScrollManager } from './scrollManager.js';

export class ExtractionEngine {
  constructor() {
    this.messageExtractor = new MessageExtractor();
    this.scrollManager = new ScrollManager();
    this.DELAY_BETWEEN_CLICKS_MS = 2000; // 2 seconds. Increase if chats load slowly.
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

      try {
        // Get the conversation name
        const conversationName = TeamsVariantDetector.getConversationName(chatItem);
        console.log(`📋 Conversation: ${conversationName}`);

        // Click on the conversation item to open it
        console.log("👆 Clicking on conversation...");
        chatItem.click();

        // Wait for the conversation to load
        await this.delay(this.DELAY_BETWEEN_CLICKS_MS);

        // Auto-scroll to load all messages if enabled
        if (this.getAutoScroll()) {
          await this.scrollToTopOfChat();
        }

        if (typeof window.__teamsExtractorMessageSequence !== 'number') {
          window.__teamsExtractorMessageSequence = 0;
        }

        const cachedMessages = this.collectCachedMessages();

        // Extract messages currently visible after scrolling
        const visibleMessages = this.extractVisibleMessages();

        const combinedMessages = this.messageExtractor.prepareMessages(
          this.messageExtractor.mergeMessages([
            ...cachedMessages,
            ...visibleMessages
          ])
        );
        
        if (combinedMessages.length > 0) {
          allConversations[conversationName] = combinedMessages;
          const earliestMessage = combinedMessages.find((msg) => msg.isoTimestamp) || combinedMessages[0];
          const earliestTimestamp = earliestMessage?.timestamp || 'Unknown';
          console.log(`✅ Extracted ${combinedMessages.length} messages from "${conversationName}" (earliest: ${earliestTimestamp})`);
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
}
