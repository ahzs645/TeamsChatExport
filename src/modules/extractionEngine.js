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

        // Extract messages
        const messages = this.extractVisibleMessages();
        
        if (messages.length > 0) {
          allConversations[conversationName] = messages;
          console.log(`✅ Extracted ${messages.length} messages from "${conversationName}"`);
        } else {
          console.log(`⚠️ No messages found in "${conversationName}"`);
        }

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
}