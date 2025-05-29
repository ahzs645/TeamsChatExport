// Handle messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('Received message:', request);
  if (request.action === 'startExport') {
    console.log('Starting export process...');
    startExport()
      .then(() => {
        console.log('Export completed successfully');
        sendResponse({ success: true });
      })
      .catch(error => {
        console.error('Export failed:', error);
        sendResponse({ error: error.message });
      });
    return true; // Required for async response
  }
});

async function startExport() {
  try {
    console.log('Getting active tab...');
    // Get the active tab
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    console.log('Active tab:', tab);
    
    if (!tab.url.includes('teams.microsoft.com')) {
      throw new Error('Please navigate to Microsoft Teams in your browser');
    }

    console.log('Extracting chat data...');
    // Extract chat data from the current page
    const chatData = await new Promise((resolve, reject) => {
      chrome.tabs.sendMessage(tab.id, { action: 'extractChat' }, (response) => {
        if (chrome.runtime.lastError) {
          console.error('Error sending message to tab:', chrome.runtime.lastError);
          reject(new Error(chrome.runtime.lastError));
          return;
        }
        console.log('Received chat data:', response);
        resolve(response);
      });
    });

    console.log('Generating HTML...');
    // Generate HTML for the chat
    const html = generateChatHTML(chatData);
    console.log('HTML generated, length:', html.length);
    
    console.log('Saving chat to file...');
    // Save the HTML file
    await saveChatToFile(html, chatData.title);

    chrome.runtime.sendMessage({
      type: 'complete',
      status: 'Export completed successfully!'
    });
  } catch (error) {
    console.error('Error in startExport:', error);
    chrome.runtime.sendMessage({
      type: 'error',
      error: error.message
    });
    throw error;
  }
}

function generateChatHTML(chatData) {
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>${chatData.title} - Teams Chat Export</title>
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          max-width: 800px;
          margin: 0 auto;
          padding: 20px;
          line-height: 1.6;
        }
        .message {
          margin: 10px 0;
          padding: 10px;
          border-radius: 4px;
          background-color: #f5f5f5;
        }
        .message-header {
          display: flex;
          justify-content: space-between;
          margin-bottom: 5px;
          color: #666;
          font-size: 0.9em;
        }
        .message-content {
          white-space: pre-wrap;
        }
        .attachment {
          margin-top: 10px;
          padding: 10px;
          background-color: #fff;
          border: 1px solid #ddd;
          border-radius: 4px;
        }
        img {
          max-width: 100%;
          height: auto;
        }
      </style>
    </head>
    <body>
      <h1>${chatData.title}</h1>
      <div class="chat-messages">
  `;

  chatData.messages.forEach(message => {
    html += `
      <div class="message">
        <div class="message-header">
          <span class="sender">${message.sender}</span>
          <span class="timestamp">${message.timestamp}</span>
        </div>
        <div class="message-content">${message.content}</div>
        ${message.attachments.length > 0 ? generateAttachmentsHTML(message.attachments) : ''}
      </div>
    `;
  });

  html += `
      </div>
    </body>
    </html>
  `;

  return html;
}

function generateAttachmentsHTML(attachments) {
  if (!attachments || attachments.length === 0) return '';
  
  return attachments.map(attachment => `
    <div class="attachment">
      <a href="${attachment.url}" target="_blank">${attachment.name}</a>
    </div>
  `).join('');
}

async function saveChatToFile(html, chatName) {
  try {
    console.log('Creating data URL for chat:', chatName);
    const dataUrl = 'data:text/html;charset=utf-8,' + encodeURIComponent(html);
    console.log('Data URL created, length:', dataUrl.length);
    
    console.log('Initiating download...');
    await chrome.downloads.download({
      url: dataUrl,
      filename: `${chatName.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.html`,
      saveAs: false
    });
    console.log('Download initiated successfully');
  } catch (error) {
    console.error('Error in saveChatToFile:', error);
    throw error;
  }
} 