document.addEventListener('DOMContentLoaded', () => {
  const extractActiveChatBtn = document.getElementById('extractActiveChatBtn');
  const openResultsBtn = document.getElementById('openResultsBtn');
  const currentChatEl = document.getElementById('currentChat');

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

  // Initialize
  loadCurrentState();

  // Refresh chat name every 2 seconds
  const stateRefreshInterval = setInterval(loadCurrentState, 2000);
  window.addEventListener('beforeunload', () => clearInterval(stateRefreshInterval));
});
