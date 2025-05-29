// Function to extract chat data from the current page
function extractChatData() {
  const chatData = {
    title: document.querySelector('[data-tid="chat-header-title"]')?.textContent || 'Unknown Chat',
    messages: []
  };

  // Get all message elements
  const messageElements = document.querySelectorAll('[data-tid="message-container"]');
  
  messageElements.forEach(element => {
    const message = {
      sender: element.querySelector('[data-tid="message-sender"]')?.textContent || 'Unknown',
      timestamp: element.querySelector('[data-tid="message-timestamp"]')?.textContent || '',
      content: element.querySelector('[data-tid="message-content"]')?.innerHTML || '',
      attachments: []
    };

    // Extract attachments if any
    const attachmentElements = element.querySelectorAll('[data-tid="attachment"]');
    attachmentElements.forEach(attachment => {
      const attachmentName = attachment.querySelector('[data-tid="attachment-name"]')?.textContent;
      const attachmentUrl = attachment.querySelector('a')?.href;
      if (attachmentName && attachmentUrl) {
        message.attachments.push({
          name: attachmentName,
          url: attachmentUrl
        });
      }
    });

    chatData.messages.push(message);
  });

  return chatData;
}

// Listen for messages from the background script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'extractChat') {
    const chatData = extractChatData();
    sendResponse(chatData);
  }
  return true;
}); 