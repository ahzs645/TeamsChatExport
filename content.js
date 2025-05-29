// Function to extract chat data from the current page
function extractChatData() {
  console.log('Starting chat data extraction...');
  
  // Debug: Log all possible chat title elements
  const possibleTitleElements = [
    document.querySelector('[data-tid="chat-header-title"]'),
    document.querySelector('[data-tid="chat-title"]'),
    document.querySelector('[data-tid="chat-header"]'),
    document.querySelector('.chat-header'),
    document.querySelector('.chat-title')
  ];
  console.log('Possible title elements:', possibleTitleElements);

  const chatData = {
    title: document.querySelector('[data-tid="chat-header-title"]')?.textContent?.trim() || 
           document.querySelector('[data-tid="chat-title"]')?.textContent?.trim() ||
           document.querySelector('.chat-header')?.textContent?.trim() ||
           'Unknown Chat',
    messages: []
  };

  console.log('Found chat title:', chatData.title);

  // Get all message elements - try different selectors
  const messageSelectors = [
    '[data-tid="message-container"]',
    '[data-tid="message"]',
    '.message-container',
    '.message'
  ];

  let messageElements = [];
  for (const selector of messageSelectors) {
    const elements = document.querySelectorAll(selector);
    console.log(`Found ${elements.length} messages with selector: ${selector}`);
    if (elements.length > 0) {
      messageElements = elements;
      break;
    }
  }

  console.log('Total messages found:', messageElements.length);
  
  messageElements.forEach((element, index) => {
    console.log(`Processing message ${index + 1}`);
    
    // Try different selectors for sender
    const senderSelectors = [
      '[data-tid="message-sender"]',
      '[data-tid="sender"]',
      '.message-sender',
      '.sender'
    ];
    
    let sender = 'Unknown';
    for (const selector of senderSelectors) {
      const senderElement = element.querySelector(selector);
      if (senderElement?.textContent) {
        sender = senderElement.textContent.trim();
        break;
      }
    }

    // Try different selectors for timestamp
    const timestampSelectors = [
      '[data-tid="message-timestamp"]',
      '[data-tid="timestamp"]',
      '.message-timestamp',
      '.timestamp'
    ];
    
    let timestamp = '';
    for (const selector of timestampSelectors) {
      const timestampElement = element.querySelector(selector);
      if (timestampElement?.textContent) {
        timestamp = timestampElement.textContent.trim();
        break;
      }
    }

    // Try different selectors for content
    const contentSelectors = [
      '[data-tid="message-content"]',
      '[data-tid="content"]',
      '.message-content',
      '.content'
    ];
    
    let content = '';
    for (const selector of contentSelectors) {
      const contentElement = element.querySelector(selector);
      if (contentElement?.innerHTML) {
        content = contentElement.innerHTML.trim();
        break;
      }
    }

    const message = {
      sender,
      timestamp,
      content,
      attachments: []
    };

    console.log(`Message ${index + 1} details:`, message);

    // Extract attachments if any
    const attachmentSelectors = [
      '[data-tid="attachment"]',
      '.attachment',
      '[data-tid="file-attachment"]',
      '.file-attachment'
    ];

    for (const selector of attachmentSelectors) {
      const attachmentElements = element.querySelectorAll(selector);
      attachmentElements.forEach(attachment => {
        const attachmentName = attachment.querySelector('[data-tid="attachment-name"]')?.textContent?.trim() ||
                             attachment.querySelector('.attachment-name')?.textContent?.trim();
        const attachmentUrl = attachment.querySelector('a')?.href;
        if (attachmentName && attachmentUrl) {
          message.attachments.push({
            name: attachmentName,
            url: attachmentUrl
          });
        }
      });
    }

    chatData.messages.push(message);
  });

  console.log('Final chat data:', chatData);
  return chatData;
}

// Listen for messages from the background script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('Content script received message:', request);
  if (request.action === 'extractChat') {
    const chatData = extractChatData();
    console.log('Sending chat data back:', chatData);
    sendResponse(chatData);
  }
  return true;
}); 