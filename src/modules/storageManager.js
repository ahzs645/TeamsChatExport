/**
 * Storage Manager Module
 * Handles Chrome storage operations and settings persistence
 */

export class StorageManager {
  constructor() {
    this.defaultSettings = {
      checkboxesEnabled: false,
      autoScrollEnabled: true,
      lastUsedChat: null,
      extractionHistory: []
    };
  }

  /**
   * Loads settings from Chrome storage
   */
  async loadSettings() {
    return new Promise((resolve) => {
      chrome.storage.local.get(Object.keys(this.defaultSettings), (result) => {
        // Merge with defaults
        const settings = { ...this.defaultSettings, ...result };
        resolve(settings);
      });
    });
  }

  /**
   * Saves a setting to Chrome storage
   */
  async saveSetting(key, value) {
    return new Promise((resolve) => {
      chrome.storage.local.set({ [key]: value }, () => {
        console.log(`💾 Saved setting: ${key} = ${value}`);
        resolve();
      });
    });
  }

  /**
   * Saves multiple settings to Chrome storage
   */
  async saveSettings(settings) {
    return new Promise((resolve) => {
      chrome.storage.local.set(settings, () => {
        console.log('💾 Saved settings:', Object.keys(settings));
        resolve();
      });
    });
  }

  /**
   * Gets a specific setting
   */
  async getSetting(key, defaultValue = null) {
    return new Promise((resolve) => {
      chrome.storage.local.get([key], (result) => {
        resolve(result[key] !== undefined ? result[key] : defaultValue);
      });
    });
  }

  /**
   * Clears all settings
   */
  async clearSettings() {
    return new Promise((resolve) => {
      chrome.storage.local.clear(() => {
        console.log('🗑️ All settings cleared');
        resolve();
      });
    });
  }

  /**
   * Saves checkbox state
   */
  async saveCheckboxState(enabled) {
    await this.saveSetting('checkboxesEnabled', enabled);
  }

  /**
   * Saves auto-scroll state
   */
  async saveAutoScrollState(enabled) {
    await this.saveSetting('autoScrollEnabled', enabled);
  }

  /**
   * Gets checkbox state
   */
  async getCheckboxState() {
    return await this.getSetting('checkboxesEnabled', false);
  }

  /**
   * Gets auto-scroll state
   */
  async getAutoScrollState() {
    return await this.getSetting('autoScrollEnabled', true);
  }

  /**
   * Saves the last used chat for quick access
   */
  async saveLastUsedChat(chatTitle) {
    await this.saveSetting('lastUsedChat', chatTitle);
  }

  /**
   * Gets the last used chat
   */
  async getLastUsedChat() {
    return await this.getSetting('lastUsedChat', null);
  }

  /**
   * Adds an extraction to history
   */
  async addExtractionToHistory(extraction) {
    const history = await this.getSetting('extractionHistory', []);
    
    // Add new extraction with timestamp
    const newExtraction = {
      ...extraction,
      timestamp: new Date().toISOString(),
      id: Date.now().toString()
    };
    
    history.unshift(newExtraction);
    
    // Keep only last 10 extractions
    const trimmedHistory = history.slice(0, 10);
    
    await this.saveSetting('extractionHistory', trimmedHistory);
    return newExtraction;
  }

  /**
   * Gets extraction history
   */
  async getExtractionHistory() {
    return await this.getSetting('extractionHistory', []);
  }

  /**
   * Clears extraction history
   */
  async clearExtractionHistory() {
    await this.saveSetting('extractionHistory', []);
  }

  /**
   * Exports all settings as JSON
   */
  async exportSettings() {
    const settings = await this.loadSettings();
    return JSON.stringify(settings, null, 2);
  }

  /**
   * Imports settings from JSON
   */
  async importSettings(jsonString) {
    try {
      const settings = JSON.parse(jsonString);
      await this.saveSettings(settings);
      return true;
    } catch (error) {
      console.error('❌ Failed to import settings:', error);
      return false;
    }
  }

  /**
   * Migrates old settings format to new format if needed
   */
  async migrateSettings() {
    const allData = await new Promise((resolve) => {
      chrome.storage.local.get(null, resolve);
    });

    let needsMigration = false;
    const newSettings = {};

    // Check for old format settings and migrate them
    if (allData.showCheckboxes !== undefined) {
      newSettings.checkboxesEnabled = allData.showCheckboxes;
      needsMigration = true;
    }

    if (allData.autoScroll !== undefined) {
      newSettings.autoScrollEnabled = allData.autoScroll;
      needsMigration = true;
    }

    if (needsMigration) {
      console.log('🔄 Migrating old settings format...');
      await this.saveSettings(newSettings);
      
      // Remove old keys
      const oldKeys = ['showCheckboxes', 'autoScroll'];
      for (const key of oldKeys) {
        if (allData[key] !== undefined) {
          chrome.storage.local.remove(key);
        }
      }
      
      console.log('✅ Settings migration completed');
    }
  }

  /**
   * Watches for storage changes and notifies listeners
   */
  onStorageChanged(callback) {
    chrome.storage.onChanged.addListener((changes, namespace) => {
      if (namespace === 'local') {
        const changedSettings = {};
        for (const key in changes) {
          changedSettings[key] = changes[key].newValue;
        }
        callback(changedSettings);
      }
    });
  }
}