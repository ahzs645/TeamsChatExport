/**
 * Checkbox Manager Module
 * Handles checkbox creation, management, and state tracking
 */

import { TeamsVariantDetector } from './teamsVariantDetector.js';

export class CheckboxManager {
  constructor() {
    this.settings = { showCheckboxes: false };
    this.hasLogged = false; // Track if we've logged once to avoid spam
  }

  /**
   * Injects a stylesheet to control the appearance of the checkboxes.
   */
  injectStylesheet() {
    const style = document.createElement('style');
    style.textContent = `
      .conversation-checkbox {
        width: 16px;
        height: 16px;
        cursor: pointer;
        accent-color: #5B5FC5;
        flex-shrink: 0;
        z-index: 1000;
        margin: 0;
        position: absolute;
      }
      
      /* Classic Teams support */
      [data-tid="chat-list-item"] {
        display: flex !important;
        align-items: center !important;
        padding-left: 25px !important;
        position: relative !important;
      }
      [data-tid="chat-list-item"] .conversation-checkbox {
        left: 5px;
        top: 50%;
        transform: translateY(-50%);
      }
      
      /* New Teams support - consistent positioning for all items */
      .fui-TreeItem {
        position: relative !important;
        padding-left: 25px !important;
      }
      
      .fui-TreeItem .conversation-checkbox {
        left: 5px;
        top: 50%;
        transform: translateY(-50%);
      }
      
      /* Ensure the tree container doesn't interfere */
      .fui-Tree {
        position: relative !important;
      }
      
      /* Make checkboxes more visible when debugging */
      .conversation-checkbox:hover {
        transform: translateY(-50%) scale(1.2);
        border: 2px solid #5B5FC5;
      }
    `;
    document.head.appendChild(style);
  }

  /**
   * Adds checkboxes to each conversation item.
   */
  addCheckboxes() {
    if (!this.settings.showCheckboxes) return;
    
    const teamsConfig = TeamsVariantDetector.getTeamsVariantSelectors();
    let chatListItems = [];
    
    // Only log once to avoid spam
    if (!this.hasLogged) {
      console.log(`🔍 Teams variant detected: ${teamsConfig.variant}`);
      this.hasLogged = true;
    }
    
    // Try selectors specific to detected Teams variant
    for (const selector of teamsConfig.chatItems) {
      const items = document.querySelectorAll(selector);
      if (items.length > 0) {
        chatListItems = items;
        // Only log once to avoid spam
        if (!this.hasLogged) {
          console.log(`✅ Using selector "${selector}" found ${items.length} items`);
        }
        break;
      }
    }
    
    if (chatListItems.length === 0) {
      console.log('❌ No chat items found with variant-specific selectors');
      return;
    }
    
    // Filter out non-chat items (folders, etc.)
    const validChatItems = Array.from(chatListItems).filter(item => {
      // Skip conversation folders
      if (item.getAttribute('data-conversation-folder') === 'true') return false;
      if (item.getAttribute('data-item-type') === 'custom-folder') return false;
      
      // For New Teams, ensure it has a valid conversation ID
      const value = item.getAttribute('data-fui-tree-item-value');
      if (value && (value.includes('Conversation') || value.includes('Chat'))) return true;
      if (value && value.includes('BizChatMetaOSConversation')) return true; // Copilot chats
      
      // For Classic Teams, check data-tid
      const tid = item.getAttribute('data-tid');
      if (tid && tid.includes('chat-list-item')) return true;
      
      return true; // Default to include if unsure
    });
    
    // Only log once to avoid spam
    if (!this.hasLogged) {
      console.log(`📋 Adding checkboxes to ${validChatItems.length} valid chat items`);
    }
    
    for (const item of validChatItems) {
      if (!item.querySelector('.conversation-checkbox')) {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'conversation-checkbox';
        checkbox.addEventListener('change', () => this.notifyPopupOfState());
        checkbox.addEventListener('click', (e) => e.stopPropagation());
        
        // Insert checkbox at the beginning
        if (item.firstChild) {
          item.insertBefore(checkbox, item.firstChild);
        } else {
          item.appendChild(checkbox);
        }
      }
    }
    this.notifyPopupOfState();
  }

  /**
   * Removes all checkboxes from the page.
   */
  removeCheckboxes() {
    const checkboxes = document.querySelectorAll('.conversation-checkbox');
    checkboxes.forEach(checkbox => checkbox.remove());
    this.notifyPopupOfState();
  }

  /**
   * Selects or deselects all checkboxes.
   */
  setAllCheckboxes(selected) {
    const checkboxes = document.querySelectorAll('.conversation-checkbox');
    checkboxes.forEach(checkbox => checkbox.checked = selected);
    this.notifyPopupOfState();
  }

  /**
   * Notifies the popup of the current selection state.
   */
  notifyPopupOfState() {
    const checkboxes = document.querySelectorAll('.conversation-checkbox');
    const selectedCount = Array.from(checkboxes).filter(c => c.checked).length;
    const allSelected = checkboxes.length > 0 && selectedCount === checkboxes.length;

    const extractButton = document.getElementById('extractSelectedButton');
    if (extractButton) {
      if (selectedCount === 0) {
        extractButton.disabled = true;
        extractButton.style.opacity = '0.5';
        extractButton.title = 'Select conversations to extract first';
      } else {
        extractButton.disabled = false;
        extractButton.style.opacity = '1';
        extractButton.title = '';
      }
    }

    try {
      if (chrome.runtime && chrome.runtime.sendMessage) {
        chrome.runtime.sendMessage({action: 'updateSelectionState', allSelected: allSelected, selectedCount: selectedCount});
      }
    } catch (error) {
      // Silently ignore extension context issues
    }
  }

  /**
   * Toggles checkbox visibility
   */
  toggleCheckboxes(enabled) {
    this.settings.showCheckboxes = enabled;
    // Save state to Chrome storage
    chrome.storage.local.set({ checkboxesEnabled: enabled });
    
    if (enabled) {
      this.addCheckboxes();
    } else {
      this.removeCheckboxes();
    }
  }

  /**
   * Gets the current state
   */
  getState() {
    return {
      checkboxesEnabled: this.settings.showCheckboxes,
      currentChat: TeamsVariantDetector.getCurrentChatTitle()
    };
  }

  /**
   * Gets selected chat items
   */
  getSelectedChatItems() {
    const teamsConfig = TeamsVariantDetector.getTeamsVariantSelectors();
    let chatListItems = [];
    
    for (const selector of teamsConfig.chatItems) {
      const items = document.querySelectorAll(selector);
      if (items.length > 0) {
        chatListItems = items;
        break;
      }
    }

    // Filter to only valid chat items with checkboxes
    return Array.from(chatListItems).filter(item => {
      const checkbox = item.querySelector('.conversation-checkbox');
      return checkbox && checkbox.checked;
    });
  }
}