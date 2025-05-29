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
    
    if (!tab?.url?.includes('teams.microsoft.com')) {
      throw new Error('Please navigate to Microsoft Teams in your browser');
    }

    console.log('Extracting chat data...');
    // Extract chat data from the current page
    const chatData = await new Promise((resolve, reject) => {
      chrome.tabs.sendMessage(tab.id, { action: 'extractChat' }, (response) => {
        if (chrome.runtime.lastError) {
          console.error('Error sending message to tab:', chrome.runtime.lastError);
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }
        if (!response) {
          reject(new Error('No response from content script'));
          return;
        }
        console.log('Received chat data:', response);
        resolve(response);
      });
    });

    if (!chatData || !chatData.messages) {
      throw new Error('No chat data received');
    }

    // If we have debug information, save it to a file
    if (chatData.debug) {
      console.log('Saving debug information...');
      const debugHtml = generateDebugHTML(chatData);
      await saveChatToFile(debugHtml, 'teams_debug_info');
      return;
    }

    // Split messages into sections of 100 messages each
    const MESSAGES_PER_SECTION = 100;
    const sections = [];
    for (let i = 0; i < (chatData.messages?.length || 0); i += MESSAGES_PER_SECTION) {
      sections.push(chatData.messages.slice(i, i + MESSAGES_PER_SECTION));
    }

    console.log(`Splitting chat into ${sections.length} sections...`);

    // Generate and save section files
    for (let i = 0; i < sections.length; i++) {
      console.log(`Generating HTML for section ${i + 1}...`);
      const sectionHtml = generateChatHTML(chatData, i, sections[i]);
      await saveChatToFile(sectionHtml, chatData.title || 'untitled_chat', i);
    }

    // Generate and save index file
    console.log('Generating index file...');
    const indexHtml = generateIndexHTML(chatData, sections.length);
    await saveChatToFile(indexHtml, `${chatData.title || 'untitled_chat'}_index`);

    chrome.runtime.sendMessage({
      type: 'complete',
      status: 'Export completed successfully!'
    });
  } catch (error) {
    console.error('Error in startExport:', error);
    chrome.runtime.sendMessage({
      type: 'error',
      error: error.message || 'Unknown error occurred'
    });
    throw error;
  }
}

function generateDebugHTML(chatData) {
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Teams Debug Information</title>
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          max-width: 1200px;
          margin: 0 auto;
          padding: 20px;
          line-height: 1.6;
        }
        .section {
          margin: 20px 0;
          padding: 20px;
          border: 1px solid #ddd;
          border-radius: 4px;
        }
        .element {
          margin: 10px 0;
          padding: 10px;
          background-color: #f5f5f5;
          border-radius: 4px;
        }
        .path {
          color: #666;
          font-family: monospace;
          margin-bottom: 5px;
        }
        .attributes {
          color: #0066cc;
          font-family: monospace;
        }
        .content {
          margin-top: 5px;
          white-space: pre-wrap;
        }
        .potential-chat {
          background-color: #e6f3ff;
        }
      </style>
    </head>
    <body>
      <h1>Teams Debug Information</h1>
      <div class="section">
        <h2>URL</h2>
        <p>${chatData.debug.url}</p>
      </div>
      
      <div class="section">
        <h2>Potential Chat Elements</h2>
        ${chatData.debug.potentialChatElements.map(el => `
          <div class="element potential-chat">
            <div class="path">${el.path}</div>
            <div class="attributes">${el.attributes}</div>
            ${el.textContent ? `<div class="content">${el.textContent}</div>` : ''}
          </div>
        `).join('')}
      </div>

      <div class="section">
        <h2>All Elements</h2>
        ${chatData.debug.elementsInfo.map(el => `
          <div class="element">
            <div class="path">${el.path}</div>
            <div class="attributes">
              ${Object.entries(el.attributes).map(([key, value]) => `${key}="${value}"`).join(' ')}
            </div>
            ${el.textContent ? `<div class="content">${el.textContent}</div>` : ''}
          </div>
        `).join('')}
      </div>
    </body>
    </html>
  `;

  return html;
}

function generateChatHTML(chatData, sectionIndex = null, sectionMessages = null) {
  if (!chatData || !chatData.title) {
    console.error('Invalid chat data:', chatData);
    return '';
  }

  const messages = sectionMessages || chatData.messages || [];
  const title = sectionIndex !== null ? `${chatData.title} - Section ${sectionIndex + 1}` : chatData.title;
  
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>${title} - Teams Chat Export</title>
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
        .navigation {
          margin: 20px 0;
          padding: 10px;
          background-color: #f0f0f0;
          border-radius: 4px;
        }
        .navigation a {
          color: #464775;
          text-decoration: none;
          margin-right: 15px;
        }
        .navigation a:hover {
          text-decoration: underline;
        }
      </style>
    </head>
    <body>
      <div class="navigation">
        <a href="index.html">Back to Index</a>
        ${sectionIndex !== null ? `<a href="section_${sectionIndex}.html">Current Section</a>` : ''}
      </div>
      <h1>${title}</h1>
      <div class="chat-messages">
  `;

  messages.forEach(message => {
    if (!message) return;
    
    html += `
      <div class="message">
        <div class="message-header">
          <span class="sender">${message.sender || 'Unknown'}</span>
          <span class="timestamp">${message.timestamp || 'Unknown time'}</span>
        </div>
        <div class="message-content">${message.content || ''}</div>
        ${(message.attachments && message.attachments.length > 0) ? generateAttachmentsHTML(message.attachments) : ''}
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

function generateIndexHTML(chatData, sectionCount) {
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>${chatData.title} - Teams Chat Export Index</title>
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          max-width: 800px;
          margin: 0 auto;
          padding: 20px;
          line-height: 1.6;
        }
        .section-link {
          display: block;
          margin: 10px 0;
          padding: 10px;
          background-color: #f5f5f5;
          border-radius: 4px;
          color: #464775;
          text-decoration: none;
        }
        .section-link:hover {
          background-color: #e0e0e0;
        }
      </style>
    </head>
    <body>
      <h1>${chatData.title} - Chat Sections</h1>
      <div class="sections">
  `;

  for (let i = 0; i < sectionCount; i++) {
    html += `
      <a href="section_${i}.html" class="section-link">Section ${i + 1}</a>
    `;
  }

  html += `
      </div>
    </body>
    </html>
  `;

  return html;
}

function generateAttachmentsHTML(attachments) {
  if (!attachments || !Array.isArray(attachments) || attachments.length === 0) return '';
  
  return attachments.map(attachment => {
    if (!attachment || !attachment.url || !attachment.name) return '';
    return `
      <div class="attachment">
        <a href="${attachment.url}" target="_blank">${attachment.name}</a>
      </div>
    `;
  }).join('');
}

async function saveChatToFile(html, chatName, sectionIndex = null) {
  try {
    console.log('Creating data URL for chat:', chatName);
    const dataUrl = 'data:text/html;charset=utf-8,' + encodeURIComponent(html);
    console.log('Data URL created, length:', dataUrl.length);
    
    const filename = sectionIndex !== null 
      ? `${chatName.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_section_${sectionIndex}.html`
      : `${chatName.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.html`;
    
    console.log('Initiating download...');
    await chrome.downloads.download({
      url: dataUrl,
      filename: filename,
      saveAs: false
    });
    console.log('Download initiated successfully');
  } catch (error) {
    console.error('Error in saveChatToFile:', error);
    throw error;
  }
} 