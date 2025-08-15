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
    chrome.storage.local.set({teamsChatData: request.data}, () => {
      chrome.tabs.create({url: chrome.runtime.getURL("results.html")});
    });
  }
});