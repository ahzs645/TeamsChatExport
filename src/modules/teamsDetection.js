/**
 * Teams Detection Module
 * Handles detection of different Teams variants and provides appropriate selectors
 */

export class TeamsDetection {
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
        let title = element.textContent.trim();
        
        // Clean up the title
        title = this.cleanConversationName(title);
        
        if (title && title.length > 0 && title !== 'Unknown Conversation') {
          console.log(`📋 Found current chat from header: "${title}" (selector: ${selector})`);
          return title;
        }
      }
    }
    
    // Method 3: Look for selected conversations with checkboxes
    const selectedCheckboxes = document.querySelectorAll('.conversation-checkbox:checked');
    if (selectedCheckboxes.length === 1) {
      // If exactly one checkbox is selected, use that conversation
      const checkedItem = selectedCheckboxes[0].closest('.fui-TreeItem, [data-tid="chat-list-item"]');
      if (checkedItem) {
        const chatName = this.getConversationName(checkedItem);
        if (chatName && chatName !== 'Unknown Conversation') {
          console.log(`📋 Found current chat from single selected checkbox: "${chatName}"`);
          return chatName;
        }
      }
    } else if (selectedCheckboxes.length > 1) {
      // Multiple selections - show count and first conversation name
      const firstCheckedItem = selectedCheckboxes[0].closest('.fui-TreeItem, [data-tid="chat-list-item"]');
      if (firstCheckedItem) {
        const firstName = this.getConversationName(firstCheckedItem);
        if (firstName && firstName !== 'Unknown Conversation') {
          console.log(`📋 Found multiple selected chats (${selectedCheckboxes.length}), showing first: "${firstName}"`);
          return `${selectedCheckboxes.length} chats selected (${firstName}...)`;
        }
      }
      console.log(`📋 Found ${selectedCheckboxes.length} selected chats`);
      return `${selectedCheckboxes.length} conversations selected`;
    }
    
    // Method 4: Try to get from page title as last resort
    const pageTitle = document.title;
    if (pageTitle && pageTitle !== 'Microsoft Teams' && !pageTitle.includes('Teams')) {
      const cleanTitle = pageTitle.replace(/\s*\|\s*Microsoft Teams\s*$/i, '').trim();
      if (cleanTitle.length > 0 && cleanTitle !== 'Microsoft Teams') {
        console.log(`📋 Found current chat from page title: "${cleanTitle}"`);
        return cleanTitle;
      }
    }
    
    console.log('⚠️ No current chat detected - showing default message');
    return 'No chat selected';
  }

  /**
   * Extracts conversation name from a chat item using multiple methods
   */
  static getConversationName(item) {
    // First try to get the fui-tree-item-value attribute which has the conversation info
    const treeValue = item.getAttribute('data-fui-tree-item-value');
    if (treeValue) {
      // Extract conversation ID for lookup (we'll use a different approach)
      console.log(`🔍 Tree item value: ${treeValue}`);
    }

    // Try various selectors for conversation names - be more specific
    const nameSelectors = [
      'span[id^="title-chat-list-item_"]',  // Fluent UI conversation titles
      '[data-tid="chat-list-item-title"]',  // Classic Teams
      '.fui-TreeItem-content .conversation-title',  // More specific New Teams
      '.conversation-title',
      'span[title]:not([title*=":"]):not([title*="AM"]):not([title*="PM"])',  // Exclude timestamps
      'div[title]:not([title*=":"]):not([title*="AM"]):not([title*="PM"])'   // Exclude timestamps
    ];
    
    for (const selector of nameSelectors) {
      const el = item.querySelector(selector);
      if (el && el.textContent.trim()) {
        let name = el.textContent.trim();
        // Clean up the name
        name = this.cleanConversationName(name);
        if (name && name.length > 0) {
          return name;
        }
      }
    }
    
    // More targeted fallback: look for specific patterns in New Teams
    const textContent = item.textContent.trim();
    if (textContent) {
      // Split by lines and look for the name pattern
      const lines = textContent.split('\n').filter(line => line.trim().length > 0);
      
      for (const line of lines) {
        const cleanLine = line.trim();
        
        // Look for patterns like "Name [ORGANIZATION]" or just "Name"
        // Stop at first timestamp or preview text
        if (cleanLine.match(/^[A-Za-z]+,\s*[A-Za-z]+\s*\[[A-Z]+\]/) || 
            cleanLine.match(/^[A-Za-z\s,.'()-]+$/)) {
          
          // Clean and return the name
          let name = this.cleanConversationName(cleanLine);
          if (name && name.length > 2) {
            return name;
          }
        }
        
        // Stop processing at first timestamp or message preview
        if (cleanLine.match(/\d{1,2}[:/]\d{2}/) || cleanLine.includes('You:')) {
          break;
        }
      }
    }
    
    return 'Unknown Conversation';
  }

  /**
   * Cleans up conversation names by removing timestamps, previews, etc.
   */
  static cleanConversationName(text) {
    if (!text) return null;
    
    let cleaned = text.trim();
    console.log(`🧹 Cleaning name: "${cleaned}"`);
    
    // First, try to extract name pattern "LastName, FirstName [ORG]" from the beginning
    // Improved regex to handle hyphens, apostrophes, and various name formats
    const nameMatch = cleaned.match(/^([A-Za-z-']+,\s*[A-Za-z\s-']+\s*\[[A-Z]+\])/);
    if (nameMatch) {
      console.log(`✅ Found name pattern: "${nameMatch[1]}"`);
      return nameMatch[1];
    }
    
    // Alternative: try to match just the name pattern without being too strict about what follows
    const relaxedNameMatch = cleaned.match(/([A-Za-z-']+,\s*[A-Za-z\s-']+\s*\[[A-Z]+\])/);
    if (relaxedNameMatch) {
      console.log(`✅ Found relaxed name pattern: "${relaxedNameMatch[1]}"`);
      return relaxedNameMatch[1];
    }
    
    // Alternative: look for just "FirstName LastName" pattern
    const simpleNameMatch = cleaned.match(/^([A-Za-z\s-]+(?:\s+[A-Za-z]+)*)/);
    if (simpleNameMatch) {
      const simpleName = simpleNameMatch[1].trim();
      // Stop at common separators
      const beforeTimestamp = simpleName.split(/\d{1,2}[:/]\d{2}/)[0].trim();
      const beforeMessage = beforeTimestamp.split(/\b(I\s|Hey\s|Hi\s|Hello\s|Thanks\s|You:)/i)[0].trim();
      
      if (beforeMessage.length > 2 && beforeMessage.length < 50) {
        console.log(`✅ Found simple name: "${beforeMessage}"`);
        return beforeMessage;
      }
    }
    
    // Fallback cleaning approach
    // Remove timestamp patterns
    cleaned = cleaned.replace(/\d{1,2}[:/]\d{2}\s*(AM|PM)?/gi, '');
    
    // Remove date patterns
    cleaned = cleaned.replace(/\d{1,2}\/\d{1,2}/g, '');
    cleaned = cleaned.replace(/\b(Yesterday|Today|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b/gi, '');
    
    // Remove "You:" and common message starters
    cleaned = cleaned.replace(/^You:\s*/gi, '');
    cleaned = cleaned.replace(/\b(I have|I will|I think|I got|Hey|Hi|Hello|Thanks|Thank you|Sure|Yes|No|OK|Okay|hmm|lol).*$/gi, '');
    
    // Clean up whitespace
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    
    console.log(`🧹 Final cleaned: "${cleaned}"`);
    
    // If reasonable length, return it
    if (cleaned.length >= 2 && cleaned.length <= 50) {
      return cleaned;
    }
    
    return null;
  }
}