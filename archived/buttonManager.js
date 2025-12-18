/**
 * Button Manager Module
 * Handles extract button creation, tooltip management, and UI interactions
 */

export class ButtonManager {
  constructor(extractionEngine, checkboxManager) {
    this.extractionEngine = extractionEngine;
    this.checkboxManager = checkboxManager;
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
        console.log('📍 Titlebar parent hierarchy:', titlebarEndSlot.parentElement?.className);
        break;
      } else {
        console.log(`❌ Not found with selector: ${selector}`);
      }
    }
    
    const settingsButton = document.querySelector('[data-tid="more-options-header"]');
    const settingsContainer = settingsButton?.closest('.fui-Primitive');
    console.log('🔧 Settings button found:', !!settingsButton);
    console.log('🔧 Settings container found:', !!settingsContainer);
    
    // Find the spacer div that comes after the settings container
    const spacerDiv = titlebarEndSlot?.querySelector('.fui-Primitive.___u5xge90');
    console.log('🔧 Spacer div found:', !!spacerDiv);
    
    if (titlebarEndSlot && settingsContainer && spacerDiv) {
      console.log('✅ All required elements found, proceeding with button creation...');

      // Create extract button container with hover toggle
      const extractButtonContainer = document.createElement('div');
      extractButtonContainer.id = 'extractButtonContainer';
      extractButtonContainer.style.cssText = `
        position: relative;
        display: inline-block;
      `;
      
      // Create extract button that matches Teams styling exactly
      const extractButton = document.createElement('button');
      extractButton.id = 'extractSelectedButton';
      extractButton.type = 'button';
      extractButton.className = 'fui-Button r1alrhcs ___1tbxa1d fhovq9v f1p3nwhy f11589ue f1q5o8ev f1pdflbu f11d4kpn f146ro5n f1s2uweq fr80ssc f1ukrpxl fecsdlb f139oj5f ft1hn21 fuxngvv fy5bs14 fsv2rcd f1h0usnq fs4ktlq f16h9ulv fx2bmrt f1omzyqd f1dfjoow f1j98vj9 fj8yq94 f4xjyn1 f1et0tmh f9ddjv3 f1wi8ngl f18ktai2 fwbmr0d f44c6la f10pi13n f1gl81tg fbmy5og fkam630 f12awlo f1tvt47b f1fig6q3 f1h878h';
      extractButton.setAttribute('aria-label', 'Extract selected conversations');
      extractButton.title = 'Extract selected conversations';
      
      // Create download icon
      const buttonIcon = document.createElement('span');
      buttonIcon.className = 'fui-Button__icon rywnvv2';
      buttonIcon.innerHTML = `
        <span class="___1gzszts f22iagw">
          <svg class="fui-Icon-filled ___1vjqft9 fjseox fez10in fg4l7m0" fill="currentColor" aria-hidden="true" width="1em" height="1em" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
            <path d="M10.5 2.75a.75.75 0 0 0-1.5 0v8.69L6.03 8.47a.75.75 0 0 0-1.06 1.06l4.5 4.5a.75.75 0 0 0 1.06 0l4.5-4.5a.75.75 0 0 0-1.06-1.06L11 11.44V2.75ZM2.5 13.5A.75.75 0 0 1 3.25 13h13.5a.75.75 0 0 1 0 1.5H3.25a.75.75 0 0 1-.75-.75Z" fill="currentColor"/>
          </svg>
          <svg class="fui-Icon-regular ___12fm75w f1w7gpdv fez10in fg4l7m0" fill="currentColor" aria-hidden="true" width="1em" height="1em" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
            <path d="M10.5 2.75a.75.75 0 0 0-1.5 0v8.69L6.03 8.47a.75.75 0 0 0-1.06 1.06l4.5 4.5a.75.75 0 0 0 1.06 0l4.5-4.5a.75.75 0 0 0-1.06-1.06L11 11.44V2.75ZM2.5 13.5A.75.75 0 0 1 3.25 13h13.5a.75.75 0 0 1 0 1.5H3.25a.75.75 0 0 1-.75-.75Z" fill="currentColor"/>
          </svg>
        </span>
      `;
      
      extractButton.appendChild(buttonIcon);
      
      // Match the exact styling of other titlebar icon buttons
      extractButton.style.cssText = `
        background-color: #5B5FC5 !important;
        color: white !important;
        border: none;
        cursor: pointer;
        transition: all 0.2s ease;
        border-radius: 4px;
        padding: 5px 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        min-width: 32px;
        height: 32px;
      `;
      
      // Create hover tooltip that looks like a Teams menu
      const autoScrollTooltip = document.createElement('div');
      autoScrollTooltip.id = 'autoScrollTooltip';
      autoScrollTooltip.className = 'fui-MenuPopover';
      autoScrollTooltip.style.cssText = `
        position: fixed;
        background: white !important;
        color: #424242 !important;
        padding: 8px !important;
        border-radius: 8px !important;
        font-size: 13px !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
        white-space: nowrap !important;
        opacity: 0;
        visibility: hidden;
        transition: opacity 0.2s ease, visibility 0.2s ease;
        z-index: 999999 !important;
        pointer-events: auto;
        cursor: pointer;
        box-shadow: 0 4px 16px rgba(0,0,0,0.16), 0 0 2px rgba(0,0,0,0.08) !important;
        border: 1px solid #e0e0e0 !important;
        min-width: 140px !important;
      `;
      
      // Add arrow to tooltip (pointing down since tooltip is above button)
      const tooltipArrow = document.createElement('div');
      tooltipArrow.style.cssText = `
        position: absolute;
        top: 100%;
        left: 50%;
        transform: translateX(-50%);
        width: 0;
        height: 0;
        border-left: 6px solid transparent;
        border-right: 6px solid transparent;
        border-top: 6px solid #333;
      `;
      autoScrollTooltip.appendChild(tooltipArrow);
      
      const updateTooltipText = () => {
        const isEnabled = this.extractionEngine.getAutoScroll();
        
        // Create a menu item that looks like Teams menu items
        const menuItem = document.createElement('div');
        menuItem.className = 'fui-MenuItem';
        menuItem.style.cssText = `
          display: flex !important;
          align-items: center !important;
          padding: 8px 12px !important;
          border-radius: 6px !important;
          cursor: pointer !important;
          transition: background-color 0.2s ease !important;
          font-size: 13px !important;
          color: #424242 !important;
        `;
        
        // Add hover effect
        menuItem.addEventListener('mouseenter', () => {
          menuItem.style.backgroundColor = '#f5f5f5';
        });
        menuItem.addEventListener('mouseleave', () => {
          menuItem.style.backgroundColor = 'transparent';
        });
        
        // Add icon and text
        menuItem.innerHTML = `
          <span style="margin-right: 8px; font-size: 14px;">
            ${isEnabled ? '✅' : '⏸️'}
          </span>
          <span>Auto-scroll: <strong>${isEnabled ? 'ON' : 'OFF'}</strong></span>
        `;
        
        // Clear tooltip and add the menu item
        autoScrollTooltip.innerHTML = '';
        autoScrollTooltip.appendChild(menuItem);
        
        // Add click handler to the menu item
        menuItem.addEventListener('click', (e) => {
          e.stopPropagation();
          const newValue = !this.extractionEngine.getAutoScroll();
          this.extractionEngine.setAutoScroll(newValue);
          console.log(`🔄 Auto-scroll ${newValue ? 'enabled' : 'disabled'}`);
          updateTooltipText();
        });
      };
      
      updateTooltipText();
      
      // Tooltip click handler is now handled by the menu item
      
      // Button click handler
      extractButton.addEventListener('click', () => {
        this.startExtraction();
      });
      
      // Hover handlers with debugging
      extractButton.addEventListener('mouseenter', () => {
        console.log('🐭 Button mouseenter triggered');
        
        // Better hover styling for the button - keep white icon
        extractButton.style.setProperty('background-color', '#4a4d9e', 'important');
        extractButton.style.setProperty('color', 'white', 'important');
        extractButton.style.transform = 'translateY(-1px)';
        
        // Force icon to stay white
        const iconSvgs = extractButton.querySelectorAll('svg');
        iconSvgs.forEach(svg => {
          svg.style.setProperty('fill', 'white', 'important');
        });
        
        // Position tooltip relative to viewport - ensure it's below the titlebar
        const buttonRect = extractButton.getBoundingClientRect();
        autoScrollTooltip.style.position = 'fixed';
        autoScrollTooltip.style.top = (buttonRect.bottom + 10) + 'px'; // Below the button instead of above
        autoScrollTooltip.style.left = (buttonRect.left + buttonRect.width/2 - 100) + 'px'; // Center it (assuming 200px max width)
        autoScrollTooltip.style.transform = 'none';
        
        autoScrollTooltip.style.opacity = '1';
        autoScrollTooltip.style.visibility = 'visible';
        autoScrollTooltip.style.display = 'block';
        
        // Remove the debugging colors - use normal styling now
        
        console.log('🐭 Button rect:', buttonRect);
        console.log('🐭 Tooltip positioned at:', {
          top: autoScrollTooltip.style.top,
          left: autoScrollTooltip.style.left
        });
      });
      
      extractButton.addEventListener('mouseleave', () => {
        console.log('🐭 Button mouseleave triggered');
        extractButton.style.setProperty('background-color', '#5B5FC5', 'important');
        extractButton.style.setProperty('color', 'white', 'important');
        extractButton.style.transform = 'translateY(0)';
        
        // Ensure icon stays white
        const iconSvgs = extractButton.querySelectorAll('svg');
        iconSvgs.forEach(svg => {
          svg.style.setProperty('fill', 'white', 'important');
        });
        // Add a small delay to allow moving to tooltip
        setTimeout(() => {
          const tooltipHovered = autoScrollTooltip.matches(':hover');
          console.log('🐭 Tooltip hovered?', tooltipHovered);
          if (!tooltipHovered) {
            autoScrollTooltip.style.opacity = '0';
            autoScrollTooltip.style.visibility = 'hidden';
            console.log('🐭 Tooltip hidden');
          }
        }, 150);
      });
      
      // Keep tooltip visible when hovering over it
      autoScrollTooltip.addEventListener('mouseenter', () => {
        console.log('🐭 Tooltip mouseenter');
        autoScrollTooltip.style.opacity = '1';
        autoScrollTooltip.style.visibility = 'visible';
      });
      
      autoScrollTooltip.addEventListener('mouseleave', () => {
        console.log('🐭 Tooltip mouseleave');
        autoScrollTooltip.style.opacity = '0';
        autoScrollTooltip.style.visibility = 'hidden';
      });
      
      extractButtonContainer.appendChild(extractButton);
      extractButtonContainer.appendChild(autoScrollTooltip);

      // Create container structure for the button to match Teams structure
      const extractContainer = document.createElement('div');
      extractContainer.className = 'fui-Primitive ___bp1bit0 f22iagw f122n59 f1p9o1ba f1tqdzup fvnq6it f1u5se5m fc397y7 f1mfizis f6s8nf5 frqcrf8';
      
      const extractWrapper = document.createElement('div');
      extractWrapper.className = 'fui-Primitive';
      extractWrapper.appendChild(extractButtonContainer);
      extractContainer.appendChild(extractWrapper);
      
      // Find all direct children of titlebarEndSlot to understand the structure
      const children = Array.from(titlebarEndSlot.children);
      console.log('📍 Titlebar children count:', children.length);
      children.forEach((child, index) => {
        console.log(`📍 Child ${index}:`, child.className || child.tagName, child.id || '');
      });
      
      // Try to find the insertion point more carefully
      let insertionPoint = null;
      
      // Look for the profile avatar section (usually the last main section)
      for (let i = children.length - 1; i >= 0; i--) {
        const child = children[i];
        if (child.querySelector && child.querySelector('[data-tid*="avatar"]')) {
          insertionPoint = child;
          console.log('📍 Found avatar section for insertion point');
          break;
        }
      }
      
      // If no avatar found, try to insert before the last non-empty div
      if (!insertionPoint) {
        for (let i = children.length - 1; i >= 0; i--) {
          const child = children[i];
          if (child.tagName === 'DIV' && child.children.length > 0) {
            insertionPoint = child;
            console.log('📍 Using last non-empty div as insertion point');
            break;
          }
        }
      }
      
      if (insertionPoint) {
        try {
          titlebarEndSlot.insertBefore(extractContainer, insertionPoint);
          console.log('✅ Button inserted successfully');
        } catch (error) {
          console.log('⚠️ Insert before failed, appending to end:', error.message);
          titlebarEndSlot.appendChild(extractContainer);
          console.log('✅ Button appended to end of titlebar');
        }
      } else {
        // Fallback: insert at the end
        titlebarEndSlot.appendChild(extractContainer);
        console.log('✅ Button appended to end of titlebar (no insertion point found)');
      }
    } else {
      console.log('❌ Required titlebar elements not found, using fallback positioning');
      this.createFallbackButton();
    }
  }

  /**
   * Creates a fallback button when titlebar insertion fails
   */
  createFallbackButton() {
    // Fallback to fixed positioning if header container not found
    const controlsContainer = document.createElement('div');
    controlsContainer.id = 'extractControlsContainer';
    controlsContainer.style.cssText = `
      position: fixed;
      top: 20px;
      right: 140px;
      display: flex;
      align-items: center;
      z-index: 10000;
      gap: 8px;
    `;
    
    // Create auto-scroll toggle
    const autoScrollContainer = document.createElement('div');
    autoScrollContainer.style.cssText = `
      display: flex;
      align-items: center;
      background: rgba(255, 255, 255, 0.9);
      border-radius: 4px;
      padding: 4px 8px;
      font-size: 11px;
      color: #333;
      border: 1px solid #ccc;
    `;
    
    const autoScrollCheckbox = document.createElement('input');
    autoScrollCheckbox.type = 'checkbox';
    autoScrollCheckbox.id = 'autoScrollToggle';
    autoScrollCheckbox.checked = this.extractionEngine.getAutoScroll();
    autoScrollCheckbox.style.cssText = `
      margin-right: 4px;
      accent-color: #5B5FC5;
      cursor: pointer;
    `;
    
    const autoScrollLabel = document.createElement('label');
    autoScrollLabel.htmlFor = 'autoScrollToggle';
    autoScrollLabel.textContent = 'Auto-scroll';
    autoScrollLabel.style.cssText = `
      cursor: pointer;
      font-size: 11px;
      color: #333;
      font-weight: 500;
    `;
    
    autoScrollCheckbox.addEventListener('change', () => {
      this.extractionEngine.setAutoScroll(autoScrollCheckbox.checked);
      console.log(`🔄 Auto-scroll ${autoScrollCheckbox.checked ? 'enabled' : 'disabled'}`);
    });
    
    autoScrollContainer.appendChild(autoScrollCheckbox);
    autoScrollContainer.appendChild(autoScrollLabel);
    
    const extractButton = document.createElement('button');
    extractButton.id = 'extractSelectedButton';
    extractButton.textContent = 'Extract Selected';
    extractButton.style.cssText = `
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
    `;
    
    extractButton.addEventListener('click', () => {
      this.startExtraction();
    });
    
    extractButton.addEventListener('mouseenter', () => {
      extractButton.style.backgroundColor = '#4a4d9e';
      extractButton.style.transform = 'translateY(-1px)';
    });
    
    extractButton.addEventListener('mouseleave', () => {
      extractButton.style.backgroundColor = '#5B5FC5';
      extractButton.style.transform = 'translateY(0)';
    });
    
    controlsContainer.appendChild(autoScrollContainer);
    controlsContainer.appendChild(extractButton);
    document.body.appendChild(controlsContainer);
  }

  /**
   * Removes the extract button and controls from the page.
   */
  removeExtractButton() {
    const existingButton = document.getElementById('extractSelectedButton');
    const existingButtonContainer = document.getElementById('extractButtonContainer');
    const existingContainer = document.getElementById('extractControlsContainer');
    
    if (existingButtonContainer) {
      // Remove button container and its container hierarchy
      const buttonContainer = existingButtonContainer.closest('.fui-Primitive.___bp1bit0');
      if (buttonContainer) {
        buttonContainer.remove();
      } else {
        existingButtonContainer.remove();
      }
    } else if (existingButton) {
      // Fallback: remove button directly
      const buttonContainer = existingButton.closest('.fui-Primitive.___bp1bit0');
      if (buttonContainer) {
        buttonContainer.remove();
      } else {
        existingButton.remove();
      }
    }
    
    if (existingContainer) {
      existingContainer.remove();
    }
  }

  /**
   * Starts the extraction process
   */
  async startExtraction() {
    const selectedItems = this.checkboxManager.getSelectedChatItems();
    if (selectedItems.length === 0) {
      console.log('⚠️ No conversations selected');
      return;
    }
    
    await this.extractionEngine.startExtraction(selectedItems);
  }
}