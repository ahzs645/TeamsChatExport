/**
 * Scroll Manager Module
 * Handles auto-scrolling functionality to load all messages
 */

export class ScrollManager {
  constructor() {
    this.settings = { autoScroll: true };
    this.DELAY_BETWEEN_SCROLLS_MS = 500; // Delay between scroll steps
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
    return this.settings.autoScroll;
  }

  /**
   * Sets auto-scroll setting
   */
  setAutoScroll(enabled) {
    this.settings.autoScroll = enabled;
    console.log(`🔄 Auto-scroll ${enabled ? 'enabled' : 'disabled'}`);
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
    
    // Store initial scroll position
    const initialScrollTop = chatContainer.scrollTop;
    console.log(`📜 Initial scroll position: ${initialScrollTop}`);
    
    // Scroll to the very top
    chatContainer.scrollTop = 0;
    await this.delay(1000); // Wait for scroll to complete
    
    // If we're already at the top, try alternative scrolling methods
    if (initialScrollTop === 0) {
      console.log('📜 Already at top, trying alternative scroll methods...');
      
      // Try scrolling the page itself
      window.scrollTo(0, 0);
      await this.delay(500);
      
      // Try scrolling other potential containers
      const bodyScrollers = document.querySelectorAll('[data-tid*="scroll"], [class*="scroll"]');
      for (const scroller of bodyScrollers) {
        if (scroller.scrollTop > 0) {
          scroller.scrollTop = 0;
          await this.delay(200);
        }
      }
    }
    
    // Perform multiple scroll attempts to ensure all messages are loaded
    let previousMessageCount = 0;
    let currentMessageCount = this.getVisibleMessageCount();
    let attempts = 0;
    const maxAttempts = 20;
    
    console.log(`📜 Starting progressive scroll. Initial message count: ${currentMessageCount}`);
    
    while (attempts < maxAttempts) {
      // Scroll up in small increments
      const scrollStep = chatContainer.scrollHeight / 10;
      for (let i = 0; i < 10; i++) {
        chatContainer.scrollTop = Math.max(0, chatContainer.scrollTop - scrollStep);
        await this.delay(this.DELAY_BETWEEN_SCROLLS_MS);
      }
      
      // Wait for new messages to load
      await this.delay(2000);
      
      currentMessageCount = this.getVisibleMessageCount();
      console.log(`📜 Attempt ${attempts + 1}: Found ${currentMessageCount} messages`);
      
      // If no new messages loaded, we've reached the top
      if (currentMessageCount === previousMessageCount) {
        console.log('📜 No new messages loaded, scrolling complete');
        break;
      }
      
      previousMessageCount = currentMessageCount;
      attempts++;
    }
    
    console.log(`📜 Scroll complete. Final message count: ${currentMessageCount}`);
    
    // Final scroll to ensure we're at the very top
    chatContainer.scrollTop = 0;
    await this.delay(1000);
  }

  /**
   * Gets the count of visible messages
   */
  getVisibleMessageCount() {
    const messageSelectors = [
      '[data-tid="chat-pane-item"]',
      '.fui-unstable-ChatItem',
      '.fui-ChatMessage',
      '.fui-ChatMyMessage',
      '[data-tid="message"]',
      '[role="article"]'
    ];
    
    for (const selector of messageSelectors) {
      const messages = document.querySelectorAll(selector);
      if (messages.length > 0) {
        return messages.length;
      }
    }
    
    return 0;
  }

  /**
   * Scrolls to a specific message element
   */
  scrollToMessage(messageElement) {
    if (messageElement && messageElement.scrollIntoView) {
      messageElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }

  /**
   * Scrolls to the bottom of the chat
   */
  async scrollToBottom() {
    const chatContainerSelectors = [
      '[data-tid="message-pane-list-viewport"]',
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
        break;
      }
    }
    
    if (chatContainer) {
      chatContainer.scrollTop = chatContainer.scrollHeight;
      await this.delay(500);
    }
  }
}