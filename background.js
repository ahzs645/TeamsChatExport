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
  } else if (request.action === "openResults") {
    // Persist the latest extraction and open the results viewer
    chrome.storage.local.get(['savedExtractions'], (result) => {
      const savedExtractions = { ...(result.savedExtractions || {}) };
      const timestamp = new Date().toLocaleString();

      if (request.data) {
        Object.keys(request.data).forEach((conversationName) => {
          const newName = `[${timestamp}] ${conversationName}`;
          savedExtractions[newName] = request.data[conversationName];
        });
      }

      chrome.storage.local.set({
        teamsChatData: request.data || {},
        savedExtractions
      }, () => {
        chrome.tabs.create({ url: chrome.runtime.getURL("results.html") }, (tab) => {
          if (tab && request.data && Object.keys(request.data).length > 0) {
            const handleUpdated = (tabId, changeInfo) => {
              if (tabId === tab.id && changeInfo.status === 'complete') {
                chrome.tabs.onUpdated.removeListener(handleUpdated);
                chrome.tabs.sendMessage(tab.id, {
                  action: 'displayData',
                  data: request.data
                }, () => {
                  if (chrome.runtime.lastError) {
                    console.warn('Results page message error:', chrome.runtime.lastError.message);
                  }
                });
              }
            };
            chrome.tabs.onUpdated.addListener(handleUpdated);
          }

          if (typeof sendResponse === 'function') {
            sendResponse({ success: true });
          }
        });
      });
    });
    return true;
  }
});
