// Function to wait for an element to be present in the DOM
function waitForElement(selectors, timeout = 10000) {
  return new Promise((resolve, reject) => {
    // Check if any of the selectors match immediately
    for (const selector of selectors) {
      const element = document.querySelector(selector);
      if (element) {
        console.log(`Found element with selector: ${selector}`);
        return resolve(element);
      }
    }

    const observer = new MutationObserver(() => {
      for (const selector of selectors) {
        const element = document.querySelector(selector);
        if (element) {
          console.log(`Found element with selector: ${selector}`);
          observer.disconnect();
          return resolve(element);
        }
      }
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true
    });

    setTimeout(() => {
      observer.disconnect();
      reject(new Error(`Timeout waiting for elements: ${selectors.join(', ')}`));
    }, timeout);
  });
}

// Function to get all elements with their attributes and content
function getAllElementsInfo() {
  const elements = document.querySelectorAll('*');
  const elementsInfo = [];

  elements.forEach(element => {
    // Get all attributes
    const attributes = {};
    for (const attr of element.attributes) {
      attributes[attr.name] = attr.value;
    }

    // Get computed styles that might be relevant
    const style = window.getComputedStyle(element);
    const relevantStyles = {
      display: style.display,
      visibility: style.visibility,
      position: style.position,
      zIndex: style.zIndex
    };

    // Get element's text content if it's not empty
    const textContent = element.textContent?.trim();
    
    // Get element's HTML content
    const innerHTML = element.innerHTML;

    elementsInfo.push({
      tagName: element.tagName,
      id: element.id,
      className: element.className,
      attributes,
      styles: relevantStyles,
      textContent: textContent || null,
      innerHTML: innerHTML || null,
      path: getElementPath(element)
    });
  });

  return elementsInfo;
}

// Function to get the full DOM path to an element
function getElementPath(element) {
  const path = [];
  while (element && element.nodeType === Node.ELEMENT_NODE) {
    let selector = element.nodeName.toLowerCase();
    if (element.id) {
      selector += `#${element.id}`;
    } else if (element.className) {
      // Handle both string and DOMTokenList className types
      const classes = typeof element.className === 'string' 
        ? element.className.split(' ')
        : Array.from(element.className);
      if (classes.length > 0) {
        selector += `.${classes.join('.')}`;
      }
    }
    path.unshift(selector);
    element = element.parentNode;
  }
  return path.join(' > ');
}

// Function to extract chat data from the current page
async function extractChatData() {
  console.log('Starting chat data extraction...');
  
  try {
    // Log the current URL to verify we're on Teams
    console.log('Current URL:', window.location.href);
    
    // Get all text content from the page
    const pageContent = document.body.innerText;
    console.log('Page content length:', pageContent.length);

    // Extract messages using pattern matching
    const messages = extractMessagesFromText(pageContent);
    console.log('Extracted messages:', messages);

    // Extract user info
    const userInfo = extractUserInfo(pageContent);
    console.log('User info:', userInfo);

    // Extract chat list
    const chatList = extractChatList(pageContent);
    console.log('Chat list:', chatList);

    // Get the chat title from the page
    const title = document.title.replace(' | Microsoft Teams', '').trim();

    return {
      title: title,
      messages: messages,
      userInfo: userInfo,
      chatList: chatList,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('Error extracting chat data:', error);
    return {
      title: 'Error Extracting Chat',
      messages: [],
      error: error.message
    };
  }
}

function extractMessagesFromText(text) {
  if (!text) {
    console.log('No text content provided');
    return [];
  }

  const messages = [];
  
  // Split text into lines for better processing
  const lines = text.split('\n') || [];
  
  // Pattern to match message content and sender
  const messagePattern = /^(.+?)\s+by\s+([^,\n]+(?:\[[^\]]+\])?)$/;
  
  // Pattern to match timestamps
  const timePattern = /(\d{1,2}:\d{2}\s*(?:AM|PM)|Today|Yesterday|\d{1,2}\/\d{1,2}\/\d{4})/i;
  
  let currentMessage = null;
  let currentSender = null;
  let currentTimestamp = null;
  let messageBuffer = [];
  
  try {
    for (let i = 0; i < (lines?.length || 0); i++) {
      const line = (lines[i] || '').trim();
      
      // Skip empty lines and UI elements
      if (!line || line.length < 3 || 
          any(skip => (line || '').toLowerCase().includes(skip), 
              ['unread', 'meeting', 'recording', 'like reaction', 'user added', 'edited', 'begin reference'])) {
        continue;
      }
      
      // Check if this line contains a timestamp
      const timeMatch = (line || '').match(timePattern);
      if (timeMatch) {
        // If we have a message in buffer, save it
        if (messageBuffer?.length > 0 && currentSender) {
          messages.push({
            content: messageBuffer.join('\n'),
            sender: currentSender,
            timestamp: currentTimestamp || timeMatch[1],
            type: 'message'
          });
          messageBuffer = [];
        }
        currentTimestamp = timeMatch[1];
        continue;
      }
      
      // Check if this line matches the message pattern
      const messageMatch = (line || '').match(messagePattern);
      if (messageMatch) {
        // If we have a message in buffer, save it
        if (messageBuffer?.length > 0 && currentSender) {
          messages.push({
            content: messageBuffer.join('\n'),
            sender: currentSender,
            timestamp: currentTimestamp || 'Unknown time',
            type: 'message'
          });
          messageBuffer = [];
        }
        
        // Start new message
        currentMessage = (messageMatch[1] || '').trim();
        currentSender = (messageMatch[2] || '').trim();
        
        // Only add to buffer if it's not a duplicate sender line
        if (currentMessage && currentSender && !currentMessage.includes(currentSender)) {
          messageBuffer.push(currentMessage);
        }
      } else if (currentMessage && currentSender) {
        // This is a continuation of the current message
        // Skip if the line is just the sender's name
        if (line !== currentSender && !line.includes('by ' + currentSender)) {
          messageBuffer.push(line);
        }
      }
    }
    
    // Add the last message if exists
    if (messageBuffer?.length > 0 && currentSender) {
      messages.push({
        content: messageBuffer.join('\n'),
        sender: currentSender,
        timestamp: currentTimestamp || 'Unknown time',
        type: 'message'
      });
    }
  } catch (error) {
    console.error('Error processing messages:', error);
    return []; // Return empty array on error
  }
  
  // Filter out very short messages and UI elements
  return messages.filter(msg => {
    if (!msg || !msg.content || !msg.sender) return false;
    
    const content = (msg.content || '').toLowerCase();
    return (msg.content || '').length > 10 && 
           !any(skip => content.includes(skip), 
                ['chat', 'unread', 'meeting', 'recording', 'like reaction', 'user added', 'edited', 'begin reference']) &&
           (msg.sender || '').length > 0 &&
           // Ensure sender is not the DAX Copilot description
           !(msg.sender || '').includes('transcribing and documenting patient visits');
  });
}

function extractUserInfo(text) {
  // Look for "(You)" pattern to identify current user
  const youPattern = /([^,\n]+\s+\(You\))/;
  const match = text.match(youPattern);
  
  const currentUser = match ? match[1] : null;
  
  return {
    currentUser: currentUser,
    organization: text.includes('[NH]') ? 'Northern Health' : null
  };
}

function extractChatList(text) {
  const chats = [];
  const lines = text.split('\n');
  
  const chatIndicators = ['PM', 'AM', 'You:', 'Unread', 'Recording is ready'];
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    // Look for lines that end with time indicators
    if (chatIndicators.some(indicator => line.includes(indicator)) && line.length > 5) {
      // Try to find chat name in previous lines
      let chatName = null;
      for (let j = Math.max(0, i - 3); j < i; j++) {
        const potentialName = lines[j].trim();
        if (potentialName && 
            potentialName.length > 3 && 
            !chatIndicators.some(indicator => potentialName.includes(indicator))) {
          chatName = potentialName;
          break;
        }
      }
      
      chats.push({
        chatName: chatName,
        lastMessage: line,
        hasUnread: line.includes('Unread')
      });
    }
  }
  
  return chats;
}

function any(predicate, array) {
  return array.some(predicate);
}

// Listen for messages from the background script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('Content script received message:', request);
  if (request.action === 'extractChat') {
    extractChatData()
      .then(chatData => {
        console.log('Sending chat data back:', chatData);
        sendResponse(chatData);
      })
      .catch(error => {
        console.error('Error in message listener:', error);
        sendResponse({
          title: 'Error Extracting Chat',
          messages: [],
          error: error.message
        });
      });
    return true; // Keep the message channel open for async response
  }
}); 