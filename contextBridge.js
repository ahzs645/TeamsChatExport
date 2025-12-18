/**
 * Context Bridge Script
 * Runs in PAGE context to find data and expose it to content script
 * via DOM elements that both contexts can access.
 */
(() => {
  const BRIDGE_DIV_ID = 'teams-extractor-context-bridge';

  const createBridgeElement = () => {
    let bridge = document.getElementById(BRIDGE_DIV_ID);
    if (!bridge) {
      bridge = document.createElement('div');
      bridge.id = BRIDGE_DIV_ID;
      bridge.style.display = 'none';
      document.body.appendChild(bridge);
    }
    return bridge;
  };

  const findToken = () => {
    // Check captured token first
    if (window.__teamsAuthToken) {
      return window.__teamsAuthToken;
    }

    // Check localStorage
    const stored = localStorage.getItem('teamsChatAuthToken');
    if (stored) {
      return stored;
    }

    // Scan MSAL storage for IC3 token
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.includes('accesstoken') && key.includes('ic3.teams.office.com')) {
        try {
          const val = JSON.parse(localStorage.getItem(key));
          if (val && val.secret) {
            return val.secret;
          }
        } catch (e) {}
      }
    }

    return null;
  };

  /**
   * Gets the name of the currently active chat from the header
   * Simple approach: The first <h2> in Teams is the active chat name
   */
  const getActiveChatName = () => {
    // Simple and reliable: first h2 is the chat name
    const h2 = document.querySelector('h2');
    if (h2 && h2.textContent) {
      const name = h2.textContent.trim();
      // Make sure it's not a date or system text
      if (name.length > 0 && name.length < 100 && !name.includes(',') && !name.includes('Microsoft')) {
        console.log('[Context Bridge] Found chat name from h2:', name);
        return name;
      }
    }

    // Fallback: Get from selected sidebar item
    const selectedItem = document.querySelector('[aria-selected="true"]');
    if (selectedItem) {
      const textContent = selectedItem.textContent || '';
      const firstLine = textContent.split('\n')[0]?.trim();
      if (firstLine && firstLine.length > 0 && firstLine.length < 50) {
        console.log('[Context Bridge] Found chat name from selected item:', firstLine);
        return firstLine;
      }
    }

    console.log('[Context Bridge] Could not find chat name');
    return null;
  };

  /**
   * Finds conversation ID by matching active chat name to sidebar items
   */
  const findConversationIdByName = (chatName) => {
    if (!chatName) return null;

    // Find all tree items with conversation IDs
    const treeItems = document.querySelectorAll('[data-fui-tree-item-value*="19:"]');

    for (const item of treeItems) {
      // Get the text content of this item (the chat name)
      const itemText = item.textContent || '';

      // Check if this item's text contains the chat name
      if (itemText.includes(chatName) || chatName.includes(itemText.trim().split('\n')[0])) {
        const treeValue = item.getAttribute('data-fui-tree-item-value');
        if (treeValue) {
          // Extract the 19:... conversation ID from the tree value
          // Format: OneGQL_.../OneGQL_...|19:conversation-id
          const match = treeValue.match(/(19:[^/\s|]+)/);
          if (match && match[1].length > 25) {
            console.log('[Context Bridge] Found conversation ID by matching name "' + chatName + '"');
            return match[1];
          }
        }
      }
    }
    return null;
  };

  const findConversationId = () => {
    // Check override first
    const override = localStorage.getItem('teamsChatConversationIdOverride') ||
                     localStorage.getItem('teamsChatApiConversationId');
    if (override && override.includes('19:')) {
      console.log('[Context Bridge] Found conversation ID in override');
      return override.match(/(19:[^"'\s<>]+)/)?.[1] || null;
    }

    // Check if chatFetchOverride captured the current conversation ID
    if (window.__teamsCurrentConversationId) {
      console.log('[Context Bridge] Found conversation ID from fetch intercept');
      return window.__teamsCurrentConversationId;
    }

    // === NEW: Match active chat name to sidebar item ===
    const activeChatName = getActiveChatName();
    if (activeChatName) {
      console.log('[Context Bridge] Active chat name:', activeChatName);
      const idByName = findConversationIdByName(activeChatName);
      if (idByName) {
        return idByName;
      }
    }

    // Check URL (some Teams versions include it)
    const urlMatch = window.location.href.match(/(19:[^/?#"]+)/);
    if (urlMatch && urlMatch[1].length > 25) {
      console.log('[Context Bridge] Found conversation ID in URL');
      return urlMatch[1];
    }

    // === CHECK PERFORMANCE ENTRIES ===
    const perfEntries = performance?.getEntriesByType?.('resource') || [];
    for (let i = perfEntries.length - 1; i >= Math.max(0, perfEntries.length - 100); i--) {
      const name = perfEntries[i].name || '';
      if ((name.includes('/conversations/') || name.includes('/threads/')) && name.includes('19:')) {
        const match = name.match(/(?:conversations|threads)\/(19:[^/?#&]+)/);
        if (match && match[1].length > 25) {
          console.log('[Context Bridge] Found conversation ID in performance entry');
          return decodeURIComponent(match[1]);
        }
      }
    }

    // === SCAN ALL TREE ITEMS WITH 19: ===
    // Look for any selected or focused item
    const treeItems = document.querySelectorAll('[data-fui-tree-item-value*="19:"]');
    for (const item of treeItems) {
      const isSelected = item.getAttribute('aria-selected') === 'true';
      const isFocused = item === document.activeElement || item.contains(document.activeElement);
      const hasSelectedChild = item.querySelector('[aria-selected="true"]');

      if (isSelected || isFocused || hasSelectedChild) {
        const val = item.getAttribute('data-fui-tree-item-value');
        const match = val.match(/(19:[^/\s|]+)/);
        if (match && match[1].length > 25) {
          console.log('[Context Bridge] Found conversation ID in selected/focused tree item');
          return match[1];
        }
      }
    }

    // Last resort: just get the first tree item with a conversation ID
    // (assuming the user is viewing it)
    for (const item of treeItems) {
      const val = item.getAttribute('data-fui-tree-item-value');
      // Only match chat conversations, not channels
      if (val && val.includes('Conversation|19:')) {
        const match = val.match(/(19:[^/\s|]+)/);
        if (match && match[1].length > 25) {
          console.log('[Context Bridge] Using first visible chat conversation ID');
          return match[1];
        }
      }
    }

    console.log('[Context Bridge] No conversation ID found');
    return null;
  };

  const updateBridge = () => {
    const bridge = createBridgeElement();
    const activeChatName = getActiveChatName();
    const conversationId = findConversationId();
    const token = findToken();

    const data = {
      token,
      conversationId,
      activeChatName,
      timestamp: Date.now()
    };

    bridge.setAttribute('data-token', data.token || '');
    bridge.setAttribute('data-conversation-id', data.conversationId || '');
    bridge.setAttribute('data-active-chat-name', data.activeChatName || '');
    bridge.setAttribute('data-timestamp', data.timestamp.toString());

    console.log('[Context Bridge] Updated:', {
      hasToken: !!data.token,
      tokenLength: data.token?.length || 0,
      activeChatName: data.activeChatName || 'NOT FOUND',
      conversationId: data.conversationId ? data.conversationId.substring(0, 40) + '...' : 'NOT FOUND'
    });

    return data;
  };

  // Update immediately
  updateBridge();

  // Update periodically and on visibility change
  setInterval(updateBridge, 2000);
  document.addEventListener('visibilitychange', updateBridge);

  // Update on click (chat selection)
  document.addEventListener('click', () => {
    setTimeout(updateBridge, 500);
  }, true);

  // Expose manual update function
  window.__teamsExtractorUpdateBridge = updateBridge;

  // Listen for update requests from content script
  document.addEventListener('teams-extractor-update-bridge', () => {
    console.log('[Context Bridge] Update requested via event');
    updateBridge();
  });

  console.log('[Context Bridge] Initialized - bridge element created');
})();
