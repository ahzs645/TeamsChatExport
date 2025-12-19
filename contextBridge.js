/**
 * Context Bridge Script
 * Runs in PAGE context to find data and expose it to content script
 * via DOM elements that both contexts can access.
 */
(() => {
  try {
  const BRIDGE_DIV_ID = 'teams-extractor-context-bridge';

  const createBridgeElement = () => {
    let bridge = document.getElementById(BRIDGE_DIV_ID);
    if (!bridge) {
      // Wait for body to exist
      if (!document.body) {
        console.warn('[Context Bridge] document.body not ready');
        return null;
      }
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
      // Make sure it's not a system text (allow commas for "Last, First" format)
      if (name.length > 0 && name.length < 100 && !name.includes('Microsoft')) {
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
   * Prioritizes: exact matches > 1-on-1 chats > partial matches
   */
  const findConversationIdByName = (chatName) => {
    if (!chatName) return null;

    const treeItems = document.querySelectorAll('[data-fui-tree-item-value*="19:"]');
    const normalizedChatName = chatName.toLowerCase().trim();

    let exactMatch = null;
    let oneOnOneMatch = null;
    let partialMatch = null;

    for (const item of treeItems) {
      const treeValue = item.getAttribute('data-fui-tree-item-value');
      if (!treeValue) continue;

      const idMatch = treeValue.match(/(19:[^/\s|]+)/);
      if (!idMatch || idMatch[1].length < 25) continue;

      const conversationId = idMatch[1];
      const isOneOnOne = conversationId.includes('@unq.gbl.spaces');

      // Get ONLY the first line (chat name), not preview text
      const itemText = item.textContent || '';
      const firstLine = itemText.split('\n')[0]?.trim() || '';
      const normalizedFirstLine = firstLine.toLowerCase();

      // Check for EXACT match (first line equals chat name)
      if (normalizedFirstLine === normalizedChatName || firstLine === chatName) {
        if (isOneOnOne) {
          // Exact match + 1-on-1 = perfect, return immediately
          return conversationId;
        }
        exactMatch = conversationId;
        continue;
      }

      // Check if first line STARTS with the chat name (handles "Name, Person [NH]" format)
      if (normalizedFirstLine.startsWith(normalizedChatName) || normalizedChatName.startsWith(normalizedFirstLine)) {
        if (isOneOnOne && !oneOnOneMatch) {
          oneOnOneMatch = conversationId;
        } else if (!partialMatch) {
          partialMatch = conversationId;
        }
      }
    }

    // Return in priority order: exact > 1-on-1 > partial
    if (exactMatch) return exactMatch;
    if (oneOnOneMatch) return oneOnOneMatch;
    if (partialMatch) return partialMatch;

    return null;
  };

  const findConversationId = () => {
    // PRIORITY 1: Match active chat name to sidebar item (most reliable for Teams v2)
    const activeChatName = getActiveChatName();
    if (activeChatName) {
      console.log('[Context Bridge] Active chat name:', activeChatName);
      const idByName = findConversationIdByName(activeChatName);
      if (idByName) {
        console.log('[Context Bridge] Found conversation ID by matching name');
        return idByName;
      }
    }

    // PRIORITY 2: Check URL (some Teams versions include conversation ID)
    const urlMatch = window.location.href.match(/(19:[^/?#"]+)/);
    if (urlMatch && urlMatch[1].length > 25) {
      console.log('[Context Bridge] Found conversation ID in URL');
      return urlMatch[1];
    }

    // PRIORITY 3: Check selected/focused tree item
    const treeItems = document.querySelectorAll('[data-fui-tree-item-value*="19:"]');
    for (const item of treeItems) {
      const isSelected = item.getAttribute('aria-selected') === 'true';
      const isFocused = item === document.activeElement || item.contains(document.activeElement);

      if (isSelected || isFocused) {
        const val = item.getAttribute('data-fui-tree-item-value');
        const match = val.match(/(19:[^/\s|]+)/);
        if (match && match[1].length > 25) {
          console.log('[Context Bridge] Found conversation ID in selected tree item');
          return match[1];
        }
      }
    }

    // PRIORITY 4: Check performance entries for recent API calls
    const perfEntries = performance?.getEntriesByType?.('resource') || [];
    for (let i = perfEntries.length - 1; i >= Math.max(0, perfEntries.length - 50); i--) {
      const name = perfEntries[i].name || '';
      if ((name.includes('/conversations/') || name.includes('/threads/')) && name.includes('19:')) {
        const match = name.match(/(?:conversations|threads)\/(19:[^/?#&]+)/);
        if (match && match[1].length > 25) {
          console.log('[Context Bridge] Found conversation ID in performance entry');
          return decodeURIComponent(match[1]);
        }
      }
    }

    // PRIORITY 5: Last resort - use fetch intercept (might be stale)
    if (window.__teamsCurrentConversationId) {
      console.log('[Context Bridge] Using fetch intercept ID (last resort)');
      return window.__teamsCurrentConversationId;
    }

    console.log('[Context Bridge] No conversation ID found');
    return null;
  };

  const updateBridge = () => {
    try {
      const bridge = createBridgeElement();
      if (!bridge) return null; // Body not ready yet

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

    // Only log if something changed (reduce console spam)
    if (window.__lastBridgeChat !== data.activeChatName) {
      console.log('[Context Bridge] Chat changed:', data.activeChatName || 'NOT FOUND');
      window.__lastBridgeChat = data.activeChatName;
    }

    return data;
    } catch (err) {
      console.error('[Context Bridge] Error in updateBridge:', err);
      return null;
    }
  };

  // Debounce helper to prevent excessive updates
  let updateTimeout = null;
  const debouncedUpdate = (delay = 500) => {
    if (updateTimeout) clearTimeout(updateTimeout);
    updateTimeout = setTimeout(updateBridge, delay);
  };

  // Update immediately on load
  updateBridge();

  // Update periodically (reduced frequency to prevent performance issues)
  setInterval(updateBridge, 5000);

  // Update on visibility change
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
      debouncedUpdate(300);
    }
  });

  // Update on click (chat selection) - debounced to prevent rapid-fire updates
  document.addEventListener('click', () => {
    debouncedUpdate(600);
  }, true);

  // Expose manual update function
  window.__teamsExtractorUpdateBridge = updateBridge;

  // Listen for update requests from content script
  document.addEventListener('teams-extractor-update-bridge', () => {
    console.log('[Context Bridge] Update requested via event');
    updateBridge();
  });

  console.log('[Context Bridge] Initialized - bridge element created');
  } catch (err) {
    console.error('[Context Bridge] FATAL ERROR:', err);
  }
})();
