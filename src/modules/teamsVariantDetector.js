/**
 * Teams Variant Detector Module
 * Handles detection of different Teams variants and provides appropriate selectors
 */

export class TeamsVariantDetector {
  /**
   * Detects the Teams variant and returns appropriate selectors
   */
  static getTeamsVariantSelectors() {
    // Detect Teams variant
    const isNewTeams = document.querySelector('.fui-FluentProvider');
    const isClassicTeams = document.querySelector('.ts-app-root');
    
    if (isNewTeams) {
      // New Teams with Fluent UI
      return {
        chatItems: ['.fui-TreeItem', '[role="treeitem"]', '[data-testid="list-item"]'],
        chatTitles: ['.fui-ThreadHeader__title', '[role="main"] h1', '.ms-Stack-inner h1'],
        variant: 'new-teams'
      };
    } else if (isClassicTeams) {
      // Classic Teams
      return {
        chatItems: ['[data-tid="chat-list-item"]', '.ms-List-cell', '[role="listitem"]'],
        chatTitles: ['[data-tid="chat-header-title"]', '[data-tid="thread-header-title"]'],
        variant: 'classic-teams'
      };
    } else {
      // Generic/unknown - try everything
      return {
        chatItems: [
          '[data-tid="chat-list-item"]',  // Classic
          '.fui-TreeItem',               // New Teams
          '[role="treeitem"]',           // New Teams alternative
          '[data-testid="list-item"]',   // New Teams data-testid
          '.ms-List-cell',
          '[role="listitem"]'
        ],
        chatTitles: [
          '[data-tid="chat-header-title"]',
          '.fui-ThreadHeader__title',
          '[role="main"] h1'
        ],
        variant: 'unknown'
      };
    }
  }

  /**
   * Gets the conversation name from a conversation element
   */
  static getConversationName(conversationElement) {
    if (!conversationElement) return 'Unknown Conversation';
    
    // Try various selectors to get the conversation name
    const nameSelectors = [
      '.fui-TreeItemLayout__main .ms-TooltipHost span',
      '.fui-TreeItemLayout__main span',
      '[data-tid="chat-header-title"]',
      '.conversation-title',
      '.ts-conversation-title',
      '.fui-Link__content',
      '.conversation-name'
    ];
    
    for (const selector of nameSelectors) {
      const nameElement = conversationElement.querySelector(selector);
      if (nameElement && nameElement.textContent.trim()) {
        const name = nameElement.textContent.trim();
        if (name && name !== '' && name !== 'undefined') {
          return name;
        }
      }
    }
    
    // Fallback: try getting text content directly
    const textContent = conversationElement.textContent?.trim();
    if (textContent && textContent.length > 0 && textContent.length < 200) {
      return textContent;
    }
    
    return 'Unknown Conversation';
  }

  /**
   * Gets the current chat title from the Teams interface
   */
  static getCurrentChatTitle() {
    console.log('🔍 Detecting current chat...');
    
    // Method 1: Try to find the currently selected/active chat item in the sidebar
    const activeChatSelectors = [
      // Fluent UI active chat items
      '.fui-TreeItem[aria-selected="true"]',
      '.fui-TreeItem[data-is-selected="true"]', 
      '.fui-TreeItem.is-selected',
      
      // Classic Teams active items
      '[data-tid="chat-list-item"][aria-selected="true"]',
      '[data-tid="chat-list-item"].active',
      '.chat-list-item.active'
    ];
    
    for (const selector of activeChatSelectors) {
      const activeItem = document.querySelector(selector);
      if (activeItem) {
        const chatName = this.getConversationName(activeItem);
        if (chatName && chatName !== 'Unknown Conversation') {
          console.log(`📋 Found current chat from active item: "${chatName}" (selector: ${selector})`);
          return chatName;
        }
      }
    }
    
    // Method 2: Try to find the main chat header title (what's currently open)
    const headerSelectors = [
      // Main chat header selectors for the current conversation
      '[data-tid="thread-header-title"] span',
      '[data-tid="thread-header-title"]',
      '.fui-ThreadHeader__title',
      'main [role="heading"]',
      '[data-tid="chat-header-title"]',
      
      // Additional header locations
      'main h1',
      '[role="main"] h1',
      '.chat-header h1'
    ];
    
    for (const selector of headerSelectors) {
      const element = document.querySelector(selector);
      if (element && element.textContent.trim()) {
        const title = element.textContent.trim();
        if (title && title !== '' && title.length < 200) {
          console.log(`📋 Found current chat from header: "${title}" (selector: ${selector})`);
          return title;
        }
      }
    }
    
    // Method 3: Try URL-based detection (last resort)
    try {
      const url = window.location.href;
      const urlMatch = url.match(/\/l\/chat\/([^\/\?]+)/);
      if (urlMatch) {
        console.log(`📋 Found chat from URL: ${urlMatch[1]}`);
        return decodeURIComponent(urlMatch[1]);
      }
    } catch (error) {
      console.log('❌ Error parsing URL for chat name:', error);
    }
    
    console.log('❌ Could not determine current chat title');
    return 'Unknown Chat';
  }

  /**
   * Gets the Teams variant as a simple string
   */
  static getTeamsVariant() {
    return this.getTeamsVariantSelectors().variant;
  }
}