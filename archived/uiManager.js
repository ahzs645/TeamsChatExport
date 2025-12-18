/**
 * UI Manager Module
 * Handles UI creation, management, and interactions including buttons and tooltips
 */

export class UIManager {
  constructor(extractionEngine, checkboxManager, storageManager = null) {
    this.extractionEngine = extractionEngine;
    this.checkboxManager = checkboxManager;
    this.scrollManager = extractionEngine.scrollManager;
    this.storageManager = storageManager;
  }

  /**
   * Injects main stylesheet for the extension
   */
  injectStylesheet() {
    // Delegate to checkbox manager for checkbox styles
    this.checkboxManager.injectStylesheet();
    
    // Add additional UI styles
    const style = document.createElement('style');
    style.id = 'teams-extractor-ui-styles';
    style.textContent = `
      /* Extract button styles */
      #extractSelectedButton {
        background-color: #5B5FC5;
        color: white;
        border: none;
        padding: 8px 14px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 12px;
        font-weight: 600;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        transition: all 0.2s ease;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      }
      
      #extractSelectedButton:hover {
        background-color: #4a4d9e;
        transform: translateY(-1px);
      }
      
      /* Auto-scroll toggle styles */
      .auto-scroll-container {
        display: flex;
        align-items: center;
        margin-right: 10px;
        font-size: 11px;
        color: #333;
      }
      
      .auto-scroll-container input {
        margin-right: 4px;
        accent-color: #5B5FC5;
        cursor: pointer;
      }
      
      .auto-scroll-container label {
        cursor: pointer;
        font-weight: 500;
      }
      
      /* Fallback button container */
      .fallback-controls-container {
        position: fixed;
        top: 10px;
        right: 10px;
        z-index: 10000;
        display: flex;
        align-items: center;
        gap: 10px;
        background: rgba(255, 255, 255, 0.95);
        padding: 8px 12px;
        border-radius: 8px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.15);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(0,0,0,0.1);
      }
    `;
    
    // Remove existing styles to avoid duplicates
    const existingStyles = document.getElementById('teams-extractor-ui-styles');
    if (existingStyles) {
      existingStyles.remove();
    }
    
    document.head.appendChild(style);
  }

  /**
   * Adds the "Extract selected" button and auto-scroll toggle to the Teams page.
   */
  addExtractButton() {
    // Only show button when checkboxes are enabled
    if (!this.checkboxManager.settings.showCheckboxes) {
      this.removeExtractButton();
      return;
    }

    const existingButton = document.getElementById('extractSelectedButton');
    if (existingButton) {
      return; // Button already exists
    }

    // Add a small delay to allow Teams to fully load
    setTimeout(() => {
      this.addExtractButtonInternal();
    }, 1000);
  }

  addExtractButtonInternal() {
    const existingButton = document.getElementById('extractSelectedButton');
    if (existingButton) {
      return; // Button already exists
    }

    // Debug: Find the correct titlebar location
    console.log('🔍 Debugging titlebar location...');
    
    // Try multiple selectors to find the titlebar
    const titlebarSelectors = [
      '.app-layout-area--title-bar .title-bar [data-tid="titlebar-end-slot"]',
      '.title-bar [data-tid="titlebar-end-slot"]',
      '[data-tid="titlebar-end-slot"]'
    ];
    
    let titlebarEndSlot = null;
    for (const selector of titlebarSelectors) {
      titlebarEndSlot = document.querySelector(selector);
      if (titlebarEndSlot) {
        console.log(`✅ Found titlebar with selector: ${selector}`);
        break;
      }
    }
    
    if (titlebarEndSlot) {
      this.insertButtonInTitlebar(titlebarEndSlot);
    } else {
      console.log('❌ Titlebar not found, creating fallback button');
      this.createFallbackButton();
    }
  }

  /**
   * Inserts the extract button into the Teams titlebar
   */
  insertButtonInTitlebar(titlebarEndSlot) {
    // Create auto-scroll toggle
    const autoScrollContainer = document.createElement('div');
    autoScrollContainer.className = 'auto-scroll-container';
    
    const autoScrollCheckbox = document.createElement('input');
    autoScrollCheckbox.type = 'checkbox';
    autoScrollCheckbox.id = 'autoScrollToggle';
    autoScrollCheckbox.checked = this.scrollManager.getAutoScroll();
    
    const autoScrollLabel = document.createElement('label');
    autoScrollLabel.htmlFor = 'autoScrollToggle';
    autoScrollLabel.textContent = 'Auto-scroll';
    
    autoScrollCheckbox.addEventListener('change', () => {
      this.scrollManager.setAutoScroll(autoScrollCheckbox.checked);
      if (this.storageManager) {
        this.storageManager.saveAutoScrollState(autoScrollCheckbox.checked);
      }
    });
    
    autoScrollContainer.appendChild(autoScrollCheckbox);
    autoScrollContainer.appendChild(autoScrollLabel);
    
    // Create extract button
    const extractButton = document.createElement('button');
    extractButton.id = 'extractSelectedButton';
    extractButton.textContent = 'Extract Selected';
    
    extractButton.addEventListener('click', () => {
      this.startExtraction();
    });
    
    // Insert into titlebar
    titlebarEndSlot.appendChild(autoScrollContainer);
    titlebarEndSlot.appendChild(extractButton);
    
    console.log('✅ Extract button added to titlebar');
  }

  /**
   * Creates a fallback button when titlebar insertion fails
   */
  createFallbackButton() {
    const controlsContainer = document.createElement('div');
    controlsContainer.className = 'fallback-controls-container';
    
    // Auto-scroll toggle
    const autoScrollContainer = document.createElement('div');
    autoScrollContainer.className = 'auto-scroll-container';
    
    const autoScrollCheckbox = document.createElement('input');
    autoScrollCheckbox.type = 'checkbox';
    autoScrollCheckbox.id = 'autoScrollToggle';
    autoScrollCheckbox.checked = this.scrollManager.getAutoScroll();
    
    const autoScrollLabel = document.createElement('label');
    autoScrollLabel.htmlFor = 'autoScrollToggle';
    autoScrollLabel.textContent = 'Auto-scroll';
    
    autoScrollCheckbox.addEventListener('change', () => {
      this.scrollManager.setAutoScroll(autoScrollCheckbox.checked);
      if (this.storageManager) {
        this.storageManager.saveAutoScrollState(autoScrollCheckbox.checked);
      }
    });
    
    autoScrollContainer.appendChild(autoScrollCheckbox);
    autoScrollContainer.appendChild(autoScrollLabel);
    
    // Extract button
    const extractButton = document.createElement('button');
    extractButton.id = 'extractSelectedButton';
    extractButton.textContent = 'Extract Selected';
    
    extractButton.addEventListener('click', () => {
      this.startExtraction();
    });
    
    controlsContainer.appendChild(autoScrollContainer);
    controlsContainer.appendChild(extractButton);
    document.body.appendChild(controlsContainer);
    
    console.log('✅ Fallback extract button created');
  }

  /**
   * Removes the extract button from the page
   */
  removeExtractButton() {
    const button = document.getElementById('extractSelectedButton');
    if (button) {
      // Remove the entire container if it's a fallback button
      const fallbackContainer = button.closest('.fallback-controls-container');
      if (fallbackContainer) {
        fallbackContainer.remove();
      } else {
        // Remove just the button and auto-scroll toggle from titlebar
        button.remove();
        const autoScrollToggle = document.getElementById('autoScrollToggle');
        if (autoScrollToggle) {
          autoScrollToggle.closest('.auto-scroll-container')?.remove();
        }
      }
      console.log('🗑️ Extract button removed');
    }
  }

  /**
   * Starts the extraction process
   */
  async startExtraction() {
    const selectedChatItems = this.checkboxManager.getSelectedChatItems();
    
    if (selectedChatItems.length === 0) {
      alert('Please select at least one conversation to extract.');
      return;
    }

    console.log(`🚀 Starting extraction for ${selectedChatItems.length} selected conversations...`);
    
    try {
      await this.extractionEngine.startExtraction(selectedChatItems);
    } catch (error) {
      console.error('❌ Extraction failed:', error);
      alert(`Extraction failed: ${error.message}`);
    }
  }

  /**
   * Shows a loading indicator
   */
  showLoadingIndicator(message = 'Processing...') {
    const existingIndicator = document.getElementById('teams-extractor-loading');
    if (existingIndicator) {
      existingIndicator.textContent = message;
      return;
    }

    const loadingDiv = document.createElement('div');
    loadingDiv.id = 'teams-extractor-loading';
    loadingDiv.style.cssText = `
      position: fixed;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      background: rgba(0, 0, 0, 0.8);
      color: white;
      padding: 20px 40px;
      border-radius: 8px;
      z-index: 10000;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      font-size: 16px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
    `;
    loadingDiv.textContent = message;
    
    document.body.appendChild(loadingDiv);
  }

  /**
   * Hides the loading indicator
   */
  hideLoadingIndicator() {
    const loadingDiv = document.getElementById('teams-extractor-loading');
    if (loadingDiv) {
      loadingDiv.remove();
    }
  }

  /**
   * Shows a notification message
   */
  showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      padding: 12px 20px;
      border-radius: 6px;
      color: white;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      font-size: 14px;
      z-index: 10000;
      max-width: 300px;
      word-wrap: break-word;
      box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
      ${type === 'success' ? 'background-color: #28a745;' : 
        type === 'error' ? 'background-color: #dc3545;' : 
        'background-color: #007bff;'}
    `;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    // Auto-remove after 5 seconds
    setTimeout(() => {
      if (notification.parentNode) {
        notification.remove();
      }
    }, 5000);
  }
}
