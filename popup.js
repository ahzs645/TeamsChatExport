document.addEventListener('DOMContentLoaded', () => {
  const extractActiveChatBtn = document.getElementById('extractActiveChatBtn');
  const openResultsBtn = document.getElementById('openResultsBtn');
  const currentChatEl = document.getElementById('currentChat');

  // Transcript elements
  const transcriptSection = document.getElementById('transcriptSection');
  const transcriptStatus = document.getElementById('transcriptStatus');
  const copyTranscriptBtn = document.getElementById('copyTranscriptBtn');
  const downloadVttBtn = document.getElementById('downloadVttBtn');
  const downloadTxtBtn = document.getElementById('downloadTxtBtn');

  // Load current chat name from content script
  const loadCurrentState = () => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0] && tabs[0].url.includes('teams.microsoft.com')) {
        chrome.tabs.sendMessage(tabs[0].id, {action: 'getState'}, (response) => {
          if (chrome.runtime.lastError) {
            currentChatEl.textContent = 'Connect to Teams first';
          } else if (response) {
            currentChatEl.textContent = response.currentChat || 'No chat selected';
          } else {
            currentChatEl.textContent = 'No chat selected';
          }
        });
      } else {
        currentChatEl.textContent = 'Navigate to Teams';
      }
    });
  };

  // Extract active chat
  extractActiveChatBtn.addEventListener('click', () => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0] && tabs[0].url.includes("teams.microsoft.com")) {
        extractActiveChatBtn.textContent = "Extracting...";
        extractActiveChatBtn.disabled = true;
        chrome.tabs.sendMessage(tabs[0].id, {action: "extractActiveChat"});
        setTimeout(() => window.close(), 500);
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

  // === TRANSCRIPT FUNCTIONS ===

  // Check transcript status
  const loadTranscriptState = () => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0] && (tabs[0].url.includes('teams.microsoft.com') || tabs[0].url.includes('stream.aspx'))) {
        chrome.tabs.sendMessage(tabs[0].id, {action: 'getTranscriptStatus'}, (response) => {
          if (chrome.runtime.lastError) {
            transcriptSection.style.display = 'none';
            return;
          }

          if (response && (response.isVideoPage || response.hasVideo)) {
            transcriptSection.style.display = 'block';
            if (response.available) {
              transcriptStatus.textContent = 'Transcript ready';
              transcriptStatus.style.backgroundColor = '#d4edda';
              transcriptStatus.style.borderColor = '#c3e6cb';
              copyTranscriptBtn.disabled = false;
              downloadVttBtn.disabled = false;
              downloadTxtBtn.disabled = false;
            } else {
              transcriptStatus.textContent = 'Start playback to load transcript';
              transcriptStatus.style.backgroundColor = '#fff3cd';
              transcriptStatus.style.borderColor = '#ffeeba';
              copyTranscriptBtn.disabled = true;
              downloadVttBtn.disabled = true;
              downloadTxtBtn.disabled = true;
            }
          } else {
            transcriptSection.style.display = 'none';
          }
        });
      } else {
        transcriptSection.style.display = 'none';
      }
    });
  };

  // Copy transcript button
  copyTranscriptBtn.addEventListener('click', () => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0]) {
        chrome.tabs.sendMessage(tabs[0].id, {action: 'extractTranscript'}, async (response) => {
          if (chrome.runtime.lastError || !response || !response.success) {
            alert(response?.error || 'Failed to get transcript');
            return;
          }
          try {
            await navigator.clipboard.writeText(response.vtt);
            copyTranscriptBtn.textContent = 'Copied!';
            setTimeout(() => { copyTranscriptBtn.textContent = 'Copy Transcript'; }, 2000);
          } catch (err) {
            alert('Failed to copy to clipboard');
          }
        });
      }
    });
  });

  // Download VTT button
  downloadVttBtn.addEventListener('click', () => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0]) {
        chrome.tabs.sendMessage(tabs[0].id, {action: 'extractTranscript'}, (response) => {
          if (chrome.runtime.lastError || !response || !response.success) {
            alert(response?.error || 'Failed to get transcript');
            return;
          }
          // Create download link
          const blob = new Blob([response.vtt], {type: 'text/vtt;charset=utf-8'});
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `transcript-${response.title}.vtt`;
          a.click();
          URL.revokeObjectURL(url);
        });
      }
    });
  });

  // Download TXT button
  downloadTxtBtn.addEventListener('click', () => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0]) {
        chrome.tabs.sendMessage(tabs[0].id, {action: 'extractTranscript'}, (response) => {
          if (chrome.runtime.lastError || !response || !response.success) {
            alert(response?.error || 'Failed to get transcript');
            return;
          }
          // Create download link
          const blob = new Blob([response.txt], {type: 'text/plain;charset=utf-8'});
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `transcript-${response.title}.txt`;
          a.click();
          URL.revokeObjectURL(url);
        });
      }
    });
  });

  // === SETTINGS FUNCTIONS ===
  const settingsHeader = document.getElementById('settingsHeader');
  const settingsContent = document.getElementById('settingsContent');
  const pageSizeInput = document.getElementById('pageSize');
  const maxPagesInput = document.getElementById('maxPages');
  const maxMessagesInfo = document.getElementById('maxMessagesInfo');
  const maxModeToggle = document.getElementById('maxModeToggle');

  // Toggle settings section
  settingsHeader.addEventListener('click', () => {
    settingsHeader.classList.toggle('expanded');
    settingsContent.classList.toggle('expanded');
  });

  // Update max messages display
  const updateMaxMessages = () => {
    const pageSize = parseInt(pageSizeInput.value, 10) || 200;
    const maxPages = parseInt(maxPagesInput.value, 10) || 15;
    const maxMessages = pageSize * maxPages;
    maxMessagesInfo.textContent = `Up to ${maxMessages.toLocaleString()} messages`;
  };

  // Update max mode toggle state based on values
  const updateMaxModeState = () => {
    const pageSize = parseInt(pageSizeInput.value, 10) || 200;
    const maxPages = parseInt(maxPagesInput.value, 10) || 15;
    const isMaxMode = pageSize === 500 && maxPages === 200;
    maxModeToggle.classList.toggle('active', isMaxMode);
  };

  // Load saved settings
  const loadSettings = () => {
    chrome.storage.local.get(['teamsChatApiPageSize', 'teamsChatApiMaxPages'], (result) => {
      pageSizeInput.value = result.teamsChatApiPageSize || 200;
      maxPagesInput.value = result.teamsChatApiMaxPages || 15;
      updateMaxMessages();
      updateMaxModeState();
    });
  };

  // Save settings on change
  const saveSettings = () => {
    const pageSize = Math.max(1, Math.min(500, parseInt(pageSizeInput.value, 10) || 200));
    const maxPages = Math.max(1, Math.min(200, parseInt(maxPagesInput.value, 10) || 15));

    // Normalize input values
    pageSizeInput.value = pageSize;
    maxPagesInput.value = maxPages;

    chrome.storage.local.set({
      teamsChatApiPageSize: pageSize,
      teamsChatApiMaxPages: maxPages
    });
    updateMaxMessages();
    updateMaxModeState();
  };

  // Max mode toggle
  maxModeToggle.addEventListener('click', () => {
    const isCurrentlyMax = maxModeToggle.classList.contains('active');
    if (isCurrentlyMax) {
      // Turn off max mode - restore defaults
      pageSizeInput.value = 200;
      maxPagesInput.value = 15;
    } else {
      // Turn on max mode
      pageSizeInput.value = 500;
      maxPagesInput.value = 200;
    }
    saveSettings();
  });

  pageSizeInput.addEventListener('change', saveSettings);
  maxPagesInput.addEventListener('change', saveSettings);
  pageSizeInput.addEventListener('input', () => { updateMaxMessages(); updateMaxModeState(); });
  maxPagesInput.addEventListener('input', () => { updateMaxMessages(); updateMaxModeState(); });

  // Initialize
  loadCurrentState();
  loadTranscriptState();
  loadSettings();

  // Refresh state every 2 seconds
  const stateRefreshInterval = setInterval(() => {
    loadCurrentState();
    loadTranscriptState();
  }, 2000);
  window.addEventListener('beforeunload', () => clearInterval(stateRefreshInterval));
});
