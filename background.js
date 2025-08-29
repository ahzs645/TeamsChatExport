chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === "extract") {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      if (tabs[0] && tabs[0].id) {
        chrome.tabs.sendMessage(tabs[0].id, {action: "extract", selectAll: request.selectAll}, (response) => {
          if (chrome.runtime.lastError) {
            // Silently ignore the error
          }
        });
      }
    });
  } else if (request.action === "download") {
    // Get existing saved extractions
    chrome.storage.local.get(['savedExtractions'], (result) => {
      const savedExtractions = result.savedExtractions || {};
      
      // Add new extraction with timestamp
      const timestamp = new Date().toLocaleString();
      Object.keys(request.data).forEach(conversationName => {
        const newName = `[${timestamp}] ${conversationName}`;
        savedExtractions[newName] = request.data[conversationName];
      });
      
      // Save accumulated data
      chrome.storage.local.set({
        teamsChatData: request.data, // Current extraction for immediate display
        savedExtractions: savedExtractions // All accumulated extractions
      }, () => {
        // Open in new tab each time
        chrome.tabs.create({url: chrome.runtime.getURL("results.html")});
      });
    });
  }
});