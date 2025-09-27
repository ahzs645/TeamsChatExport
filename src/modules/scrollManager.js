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
    
    window.__teamsExtractorMessageCache = new Map();
    window.__teamsExtractorMessageSequence = 0;

    this.cacheVisibleMessages();

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
    let currentMessageCount = this.getVisibleMessageCount();
    let maxMessageCount = currentMessageCount;
    let stalledAttempts = 0;
    let attempts = 0;
    const maxAttempts = 25;
    
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

      this.cacheVisibleMessages();

      const clickedLoadMore = this.clickLoadMoreButtons();
      if (clickedLoadMore) {
        await this.delay(2500);
        currentMessageCount = this.getVisibleMessageCount();
        console.log(`📜 Load more triggered, updated message count: ${currentMessageCount}`);
        this.cacheVisibleMessages();
      }

      if (this.pulseVirtualLoader()) {
        await this.delay(2500);
        currentMessageCount = this.getVisibleMessageCount();
        console.log(`📜 Virtual loader pulsed, updated message count: ${currentMessageCount}`);
        this.cacheVisibleMessages();
      }

      if (currentMessageCount > maxMessageCount) {
        maxMessageCount = currentMessageCount;
        stalledAttempts = 0;
      } else {
        stalledAttempts += 1;
      }

      // If Teams has virtualised away the oldest DOM nodes while we scrolled, scroll down a bit to let it reuse the buffer
      if (currentMessageCount < maxMessageCount) {
        chatContainer.scrollTop = Math.min(chatContainer.scrollHeight, chatContainer.scrollTop + chatContainer.clientHeight * 0.5);
        await this.delay(750);
        chatContainer.scrollTop = 0;
        await this.delay(1500);
        currentMessageCount = this.getVisibleMessageCount();
        console.log(`📜 Recovered count after buffer adjustment: ${currentMessageCount}`);
        if (currentMessageCount > maxMessageCount) {
          maxMessageCount = currentMessageCount;
          stalledAttempts = 0;
        }
        this.cacheVisibleMessages();
      }

      if (stalledAttempts >= 3) {
        console.log('📜 No additional messages detected after multiple attempts, stopping scroll');
        break;
      }

      attempts++;
    }

    console.log(`📜 Scroll complete. Final message count: ${currentMessageCount}`);
    this.cacheVisibleMessages();

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
   * Cache currently visible messages for later reconstruction
   */
  cacheVisibleMessages() {
    if (!window.__teamsExtractorMessageCache) {
      window.__teamsExtractorMessageCache = new Map();
    }

    if (typeof window.__teamsExtractorMessageSequence !== 'number') {
      window.__teamsExtractorMessageSequence = 0;
    }

    const cache = window.__teamsExtractorMessageCache;

    const messageSelectors = [
      '[data-tid="chat-pane-item"]',
      '.fui-unstable-ChatItem'
    ];

    let nodes = [];
    messageSelectors.some((selector) => {
      nodes = Array.from(document.querySelectorAll(selector));
      return nodes.length > 0;
    });

    nodes.forEach((node) => {
      const mid = node.getAttribute('data-mid') || node.dataset?.mid || node.getAttribute('id');
      const author = node.querySelector('[data-tid="message-author-name"]')?.textContent?.trim() || '';
      const timestamp = node.querySelector('[data-tid="message-timestamp"], time')?.textContent?.trim() || '';
      const textPreview = node.textContent?.trim().slice(0, 60) || '';
      const key = mid || `${timestamp}::${author}::${textPreview}`;

      if (!cache.has(key)) {
        window.__teamsExtractorMessageSequence += 1;
        cache.set(key, {
          html: node.outerHTML,
          sequence: window.__teamsExtractorMessageSequence
        });
      }
    });
  }

  /**
   * Attempts to click any available load-more buttons to fetch older messages
   */
  clickLoadMoreButtons() {
    const loadButtonSelectors = [
      'button[data-tid*="load" i]',
      'button[data-testid*="load" i]',
      'button[aria-label*="older" i]',
      'button[aria-label*="load" i]',
      'div[role="button"][data-tid*="load" i]'
    ];

    let clicked = false;

    loadButtonSelectors.forEach((selector) => {
      document.querySelectorAll(selector).forEach((button) => {
        if (!button || typeof button.click !== 'function') {
          return;
        }

        if (button.dataset?.teamsExtractorClicked === 'true') {
          return;
        }

        const label = (button.textContent || button.getAttribute('aria-label') || '').toLowerCase();
        if (label.includes('older') || label.includes('load')) {
          console.log('🔁 Clicking load-more button to fetch older messages');
          button.dataset.teamsExtractorClicked = 'true';
          button.click();
          clicked = true;
        }
      });
    });

    return clicked;
  }

  pulseVirtualLoader() {
    let pulsed = false;
    const loaderSelectors = [
      '[data-testid="virtual-list-loader"]',
      '[data-testid="vl-placeholders"]'
    ];

    loaderSelectors.forEach((selector) => {
      document.querySelectorAll(selector).forEach((loader) => {
        if (!(loader instanceof HTMLElement)) {
          return;
        }

        console.log('🔁 Pulsing virtual loader to keep early messages in buffer');
        loader.style.display = 'none';
        pulsed = true;
      });
    });

    if (pulsed) {
      setTimeout(() => {
        loaderSelectors.forEach((selector) => {
          document.querySelectorAll(selector).forEach((loader) => {
            if (loader instanceof HTMLElement) {
              loader.style.display = '';
            }
          });
        });
      }, 200);
    }

    return pulsed;
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
