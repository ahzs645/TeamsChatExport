document.addEventListener('DOMContentLoaded', () => {
  const checkboxToggle = document.getElementById('checkboxToggle');
  const selectAllBtn = document.getElementById('selectAllBtn');
  const extractBtn = document.getElementById('extractBtn');
  const openResultsBtn = document.getElementById('openResultsBtn');
  const currentChatEl = document.getElementById('currentChat');
  
  let checkboxesEnabled = false;
  let allSelected = false;

  // Load current state from content script
  const loadCurrentState = () => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0] && tabs[0].url.includes('teams.microsoft.com')) {
        chrome.tabs.sendMessage(tabs[0].id, {action: 'getState'}, (response) => {
          if (chrome.runtime.lastError) {
            currentChatEl.textContent = 'Connect to Teams first';
          } else if (response) {
            // Update chat title
            currentChatEl.textContent = response.currentChat || 'No chat selected';
            // Update checkbox toggle state
            if (response.checkboxesEnabled !== undefined) {
              updateToggle(response.checkboxesEnabled);
            }
          } else {
            currentChatEl.textContent = 'No chat selected';
          }
        });
      } else {
        currentChatEl.textContent = 'Navigate to Teams';
      }
    });
  };

  // Update toggle state
  const updateToggle = (enabled) => {
    checkboxesEnabled = enabled;
    if (enabled) {
      checkboxToggle.classList.add('active');
      selectAllBtn.disabled = false;
    } else {
      checkboxToggle.classList.remove('active');
      selectAllBtn.disabled = true;
    }
  };

  // Update select all button state
  const updateSelectAllButton = (isSelected) => {
    allSelected = isSelected;
    if (allSelected) {
      selectAllBtn.textContent = 'Deselect All';
      selectAllBtn.classList.remove('btn-secondary');
      selectAllBtn.classList.add('btn-primary');
    } else {
      selectAllBtn.textContent = 'Select All';
      selectAllBtn.classList.remove('btn-primary');
      selectAllBtn.classList.add('btn-secondary');
    }
  };

  // Send message to content script
  const sendToContentScript = (message, callback) => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0] && tabs[0].id && tabs[0].url.includes('teams.microsoft.com')) {
        chrome.tabs.sendMessage(tabs[0].id, message, (response) => {
          if (chrome.runtime.lastError) {
            // Silently handle connection errors - content script may not be ready
          } else if (callback) {
            callback(response);
          }
        });
      }
    });
  };

  // Toggle checkboxes on/off
  checkboxToggle.addEventListener('click', () => {
    const newState = !checkboxesEnabled;
    updateToggle(newState);
    sendToContentScript({action: 'toggleCheckboxes', enabled: newState});
  });

  // Select/deselect all
  selectAllBtn.addEventListener('click', () => {
    if (!checkboxesEnabled) return;
    
    const newState = !allSelected;
    updateSelectAllButton(newState);
    sendToContentScript({action: newState ? 'selectAll' : 'deselectAll'});
  });

  // Extract data
  extractBtn.addEventListener('click', () => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0] && tabs[0].url.startsWith("https://teams.microsoft.com")) {
        chrome.tabs.sendMessage(tabs[0].id, {action: "extract"});
        window.close(); // Close popup after starting extraction
      } else {
        alert("Please navigate to teams.microsoft.com to use this extension.");
      }
    });
  });

  // Open results viewer
  openResultsBtn.addEventListener('click', () => {
    chrome.tabs.create({url: chrome.runtime.getURL("results.html")});
    window.close();
  });

  // Listen for updates from content script
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'updateSelectionState') {
      updateSelectAllButton(request.allSelected);
    }
  });

  // Initialize
  loadCurrentState();
  
  // Refresh current state every 2 seconds while popup is open
  const stateRefreshInterval = setInterval(loadCurrentState, 2000);
  
  // Clean up interval when popup closes
  window.addEventListener('beforeunload', () => {
    clearInterval(stateRefreshInterval);
  });
});