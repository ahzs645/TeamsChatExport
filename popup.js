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

  // Helper to download file
  const downloadFile = (content, filename) => {
    const blob = new Blob([content], {type: 'text/plain;charset=utf-8'});
    const url = URL.createObjectURL(blob);
    chrome.downloads.download({
      url: url,
      filename: filename,
      saveAs: true
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

  // Initialize
  loadCurrentState();
  loadTranscriptState();

  // Refresh state every 2 seconds
  const stateRefreshInterval = setInterval(() => {
    loadCurrentState();
    loadTranscriptState();
  }, 2000);
  window.addEventListener('beforeunload', () => clearInterval(stateRefreshInterval));
});
