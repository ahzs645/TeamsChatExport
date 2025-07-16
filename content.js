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

// Function to normalize dates by adding current year if missing
function normalizeDate(dateString) {
  const currentYear = new Date().getFullYear();
  
  // Check if date already has a year
  if (dateString.includes(currentYear.toString()) || dateString.includes('2024') || dateString.includes('2025')) {
    return dateString; // Already has year
  }
  
  // Check if it's a month + day format (e.g., "May 27", "December 17")
  const monthDayPattern = /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}$/i;
  if (monthDayPattern.test(dateString)) {
    return `${dateString}, ${currentYear}`;
  }
  
  return dateString; // Return as-is if no pattern matches
}

// Function to convert weekday names to actual dates
function convertWeekdayToDate(weekdayName) {
  const weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  
  const today = new Date();
  const currentDayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
  const targetDayOfWeek = weekdays.findIndex(day => day.toLowerCase() === weekdayName.toLowerCase());
  
  if (targetDayOfWeek === -1) return weekdayName; // Return original if not found
  
  // Calculate days difference (assuming it's referring to this week)
  let daysDiff = targetDayOfWeek - currentDayOfWeek;
  
  // If the weekday already passed this week, assume it's from last week
  if (daysDiff > 0) {
    daysDiff -= 7;
  }
  
  const targetDate = new Date(today);
  targetDate.setDate(today.getDate() + daysDiff);
  
  const month = months[targetDate.getMonth()];
  const day = targetDate.getDate();
  const year = targetDate.getFullYear();
  
  return `${month} ${day}, ${year}`;
}

// Function to extract reactions from message content
function extractReactions(content) {
  if (!content) {
    return { cleanedContent: content, reactions: [] };
  }
  
  const reactions = [];
  let cleanedContent = content;
  
  // Pattern for reactions: "2 Heart reactions.2" or "1 Like reaction." 
  const reactionPattern = /(\d+)\s+(Like|Heart|Laugh|Surprised?|Sad|Angry)\s+reactions?\s*(?:with [^.]*)?\.(\d*)/gi;
  
  let match;
  while ((match = reactionPattern.exec(content)) !== null) {
    const count = parseInt(match[1]);
    const type = match[2].toLowerCase();
    
    // Map reaction types to emojis
    const reactionEmojis = {
      'like': '👍',
      'heart': '❤️', 
      'laugh': '😆',
      'surprised': '😮',
      'surprise': '😮',
      'sad': '😢',
      'angry': '😠'
    };
    
    const emoji = reactionEmojis[type] || '👍';
    
    reactions.push({
      type: type,
      count: count,
      emoji: emoji
    });
  }
  
  // Remove all reaction text from content
  cleanedContent = cleanedContent.replace(/\d+\s+(Like|Heart|Laugh|Surprised?|Sad|Angry)\s+reactions?\s*(?:with [^.]*)?\.(\d*)/gi, '').trim();
  
  // Clean up any trailing periods or numbers
  cleanedContent = cleanedContent.replace(/\.\d*$/, '').trim();
  
  return { cleanedContent, reactions };
}

// Function to extract media (GIFs, images) from DOM element
function extractMediaFromElement(element) {
  const media = [];
  const seenSources = new Set(); // Track seen sources to avoid duplicates
  
  // Look for GIF elements
  const gifElements = element.querySelectorAll('[aria-label*="GIF"], img[src*="giphy"], img[src*="tenor"]');
  gifElements.forEach(gifEl => {
    const ariaLabel = gifEl.getAttribute('aria-label') || gifEl.parentElement?.getAttribute('aria-label') || '';
    let src = '';
    
    // Find the actual image source
    if (gifEl.tagName === 'IMG') {
      src = gifEl.src;
    } else {
      const img = gifEl.querySelector('img');
      if (img) src = img.src;
    }
    
    // Filter out Teams internal emoticons and duplicates
    if (src && 
        (src.includes('giphy') || src.includes('tenor') || ariaLabel.includes('GIF')) &&
        !src.includes('statics.teams.cdn.office.net') &&
        !src.includes('emoticons') &&
        !seenSources.has(src)) {
      
      seenSources.add(src);
      media.push({
        type: 'gif',
        src: src,
        alt: ariaLabel || 'GIF',
        title: ariaLabel || 'GIF'
      });
    }
  });
  
  // Look for other images that might be media
  const images = element.querySelectorAll('img:not([src*="giphy"]):not([src*="tenor"])');
  images.forEach(img => {
    if (img.src && 
        !img.src.includes('emoticons') && 
        !img.src.includes('avatar') &&
        !img.src.includes('statics.teams.cdn.office.net') &&
        !seenSources.has(img.src)) {
      
      const alt = img.alt || img.title || '';
      if (alt && !alt.includes('reaction') && !alt.includes('emoji')) {
        seenSources.add(img.src);
        media.push({
          type: 'image',
          src: img.src,
          alt: alt,
          title: alt
        });
      }
    }
  });
  
  return media;
}

// Function to extract reactions from DOM structure
function extractReactionsFromDOM(element) {
  const reactions = [];
  
  // Look for reaction elements in Teams DOM structure
  const reactionElements = element.querySelectorAll('[id*="message-"][id*="-reaction"], [aria-labelledby*="message-"][aria-labelledby*="-reaction"]');
  reactionElements.forEach(reactionEl => {
    const textContent = reactionEl.textContent?.trim() || '';
    const match = textContent.match(/(\d+)\s+(Like|Heart|Laugh|Surprised?|Sad|Angry)\s+reactions?/i);
    
    if (match) {
      const count = parseInt(match[1]);
      const type = match[2].toLowerCase();
      
      const reactionEmojis = {
        'like': '👍',
        'heart': '❤️', 
        'laugh': '😆',
        'surprised': '😮',
        'surprise': '😮',
        'sad': '😢',
        'angry': '😠'
      };
      
      const emoji = reactionEmojis[type] || '👍';
      
      reactions.push({
        type: type,
        count: count,
        emoji: emoji
      });
    }
  });
  
  // Also look for reaction summary areas
  const reactionSummary = element.querySelector('[data-tid="diverse-reaction-summary"]');
  if (reactionSummary) {
    const reactionButtons = reactionSummary.querySelectorAll('[data-tid="diverse-reaction-pill-button"]');
    reactionButtons.forEach(button => {
      const reactionText = button.textContent?.trim() || '';
      const match = reactionText.match(/(\d+)\s+(Like|Heart|Laugh|Surprised?|Sad|Angry)\s+reactions?/i);
      
      if (match) {
        const count = parseInt(match[1]);
        const type = match[2].toLowerCase();
        
        const reactionEmojis = {
          'like': '👍',
          'heart': '❤️', 
          'laugh': '😆',
          'surprised': '😮',
          'surprise': '😮',
          'sad': '😢',
          'angry': '😠'
        };
        
        const emoji = reactionEmojis[type] || '👍';
        
        // Avoid duplicates
        if (!reactions.find(r => r.type === type)) {
          reactions.push({
            type: type,
            count: count,
            emoji: emoji
          });
        }
      }
    });
  }
  
  return reactions;
}

// Function to extract chat data from the current page
async function extractChatData() {
  console.log('Starting chat data extraction...');
  
  try {
    // Log the current URL to verify we're on Teams
    console.log('Current URL:', window.location.href);
    
    // Try to find the main chat container using common Teams selectors
    const chatContainer = findChatContainer();
    console.log('Found chat container:', chatContainer);
    
    if (!chatContainer) {
      // Fallback to debug mode - return all elements for analysis
      console.log('No chat container found, enabling debug mode...');
      return {
        title: 'Debug Mode - No Chat Container Found',
        messages: [],
        debug: {
          url: window.location.href,
          potentialChatElements: findPotentialChatElements(),
          elementsInfo: getAllElementsInfo().slice(0, 50) // Limit to first 50 elements
        }
      };
    }
    
    // Extract messages from the chat container
    const messages = extractMessagesFromDOM(chatContainer);
    console.log('Extracted messages:', messages);

    // Get the chat title
    const title = getChatTitle();

    return {
      title: title,
      messages: messages,
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

// Function to find the main chat container
function findChatContainer() {
  // Updated selectors based on debug analysis
  const selectors = [
    'div#chat-pane-list',
    'div[id*="chat-pane"]',
    '.fui-Flex.___1ccp5kb',
    '[data-tid="chat-pane-body"]',
    '[data-tid="chat-container"]', 
    '[role="main"] [role="log"]'
  ];
  
  for (const selector of selectors) {
    const element = document.querySelector(selector);
    if (element) {
      console.log(`Found chat container with selector: ${selector}`);
      return element;
    }
  }
  
  return null;
}

// Function to get chat title
function getChatTitle() {
  // Try to find chat title in various locations
  const titleSelectors = [
    '[data-tid="chat-header-title"]',
    '[data-tid="chat-title"]',
    '.chat-header h1',
    '.chat-header h2',
    'h1[data-tid*="title"]'
  ];
  
  for (const selector of titleSelectors) {
    const element = document.querySelector(selector);
    if (element && element.textContent.trim()) {
      return element.textContent.trim();
    }
  }
  
  // Fallback to page title
  return document.title.replace(' | Microsoft Teams', '').trim();
}

// Function to find potential chat elements for debugging
function findPotentialChatElements() {
  const potentialSelectors = [
    '[data-tid*="chat"]',
    '[data-tid*="message"]',
    '[role="log"]',
    '[role="main"]',
    '.ui-chat',
    '.message',
    '[data-testid*="message"]',
    '[data-testid*="chat"]'
  ];
  
  const elements = [];
  potentialSelectors.forEach(selector => {
    const found = document.querySelectorAll(selector);
    found.forEach(el => {
      elements.push({
        selector: selector,
        path: getElementPath(el),
        attributes: Object.fromEntries(Array.from(el.attributes).map(attr => [attr.name, attr.value])),
        textContent: el.textContent?.trim().substring(0, 200) || null
      });
    });
  });
  
  return elements.slice(0, 20); // Limit results
}

// Function to extract messages from DOM elements
function extractMessagesFromDOM(container) {
  const messages = [];
  
  // Updated selectors based on debug analysis
  const messageSelectors = [
    '.fui-unstable-ChatItem',
    'div[class*="ChatItem"]',
    'div[id*="message-body"]',
    '.fui-ChatMessageCompact',
    '[data-tid*="message"]',
    '.message'
  ];
  
  let messageElements = [];
  for (const selector of messageSelectors) {
    messageElements = container.querySelectorAll(selector);
    if (messageElements.length > 0) {
      console.log(`Found ${messageElements.length} messages with selector: ${selector}`);
      break;
    }
  }
  
  if (messageElements.length === 0) {
    console.log('No message elements found, falling back to text extraction');
    return extractMessagesFromText(container.textContent || '');
  }
  
  // Track current date for message context
  let currentDate = '';
  
  messageElements.forEach((msgEl, index) => {
    try {
      const message = extractMessageFromElement(msgEl);
      if (message && message.content && message.content.length > 5) {
        
        // Check if this message is a date marker
        const datePattern = /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:,?\s+\d{4})?$|^\d{1,2}\/\d{1,2}\/\d{4}$|^(Today|Yesterday)$|^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*,?\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2}$/i;
        const weekdayPattern = /^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$/i;
        
        if (datePattern.test(message.content.trim())) {
          currentDate = normalizeDate(message.content.trim());
          // Don't add date markers as regular messages
          return;
        } else if (weekdayPattern.test(message.content.trim())) {
          // Convert weekday to actual date
          currentDate = convertWeekdayToDate(message.content.trim());
          // Don't add weekday markers as regular messages
          return;
        }
        
        // Add current date context to timestamp if we have one
        if (currentDate && message.timestamp !== 'Unknown time') {
          message.timestamp = `${currentDate}, ${message.timestamp}`;
        }
        
        messages.push(message);
      }
    } catch (error) {
      console.error(`Error extracting message ${index}:`, error);
    }
  });
  
  return messages;
}

// Function to extract a single message from a DOM element
function extractMessageFromElement(element) {
  const fullText = element.textContent?.trim() || '';
  
  // Extract GIFs and media from DOM
  const media = extractMediaFromElement(element);
  
  // Extract reactions from DOM structure
  const domReactions = extractReactionsFromDOM(element);
  
  // Check for quote/reference blocks first
  if (fullText.includes('Reference,')) {
    const lines = fullText.split('\n');
    let referenceTopic = '';
    let messageDate = '';
    let quotedInfo = '';
    let remainingContent = '';
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      if (line.startsWith('Reference,')) {
        referenceTopic = line.substring(10).trim(); // Remove "Reference, "
      } else if (line.match(/^[A-Za-z]+\s+\d{1,2},?\s+\d{1,2}:\d{2}\s*(?:AM|PM)$/)) {
        messageDate = line;
      } else if (line.includes('[NH]') && line.includes('/')) {
        // This line contains the quoted sender, date, and content
        quotedInfo = line;
      } else if (line.length > 0 && !line.match(/^\d+\s*Like reaction/)) {
        // This is the new message content
        remainingContent += (remainingContent ? '\n' : '') + line;
      }
    }
    
    if (quotedInfo) {
      // Parse the quoted info line: "Sender7/7/2025, 3:46 PMQuoted content"
      const quotedPattern = /([A-Za-z]+,\s*[A-Za-z]+\s*\[[^\]]+\])(\d{1,2}\/\d{1,2}\/\d{4}),?\s*(\d{1,2}:\d{2}\s*(?:AM|PM))(.+)/;
      const quotedMatch = quotedInfo.match(quotedPattern);
      
      if (quotedMatch) {
        const quotedSender = quotedMatch[1];
        const quotedDate = quotedMatch[2];
        const quotedTime = quotedMatch[3];
        const quotedContent = quotedMatch[4];
        
        // Clean up content - remove like reactions
        remainingContent = remainingContent.replace(/\d+\s*Like reaction\.?/g, '').trim();
        
        // Format the content with quote block
        const formattedContent = `> **Quoted from ${quotedSender} (${quotedDate}, ${quotedTime}):**
> ${quotedContent.trim()}

${remainingContent}`;
        
        // The actual sender is the person making the reply, not the quoted person
        // Since we can't easily determine this from the quote block alone, mark as Reply
        const { cleanedContent, reactions } = extractReactions(remainingContent);
        const finalContent = `> **Quoted from ${quotedSender} (${quotedDate}, ${quotedTime}):**
> ${quotedContent.trim()}

${cleanedContent}`;
        
        return {
          sender: 'Reply (quoting ' + quotedSender + ')',
          content: finalContent,
          reactions: domReactions.length > 0 ? domReactions : reactions,
          media: media,
          timestamp: messageDate,
          type: 'message'
        };
      }
    }
  }
  
  // Updated Teams format: "Preview... by Sender NameSender Name (repeated)TimestampFull Message"
  // More flexible pattern to handle the actual Teams format
  const messagePattern = /^(.+?)\s+by\s+([A-Za-z]+,\s*[A-Za-z]+\s*(?:\[[^\]]+\])?)\s*\1?\s*(\d{1,2}:\d{2}\s*(?:AM|PM))\s*(.+)$/;
  
  let match = fullText.match(messagePattern);
  
  if (match) {
    const preview = match[1];
    const sender = match[2];
    const timestamp = match[3];
    const fullMessage = match[4];
    
    const { cleanedContent, reactions } = extractReactions(fullMessage || preview);
    return {
      sender: sender,
      content: cleanedContent,
      reactions: domReactions.length > 0 ? domReactions : reactions,
      media: media,
      timestamp: timestamp,
      type: 'message'
    };
  }
  
  // Alternative pattern for simpler cases
  const simplePattern = /^(.+?)\s+by\s+([A-Za-z]+,\s*[A-Za-z]+\s*(?:\[[^\]]+\])?)\s*([A-Za-z]+,\s*[A-Za-z]+\s*(?:\[[^\]]+\])?)\s*(\d{1,2}:\d{2}\s*(?:AM|PM))\s*(.+)$/;
  match = fullText.match(simplePattern);
  
  if (match) {
    const preview = match[1];
    const sender = match[2]; // Use first sender occurrence
    const timestamp = match[4];
    const fullMessage = match[5];
    
    const { cleanedContent, reactions } = extractReactions(fullMessage || preview);
    return {
      sender: sender,
      content: cleanedContent,
      reactions: domReactions.length > 0 ? domReactions : reactions,
      media: media,
      timestamp: timestamp,
      type: 'message'
    };
  }
  
  // Manual parsing approach for complex cases
  let sender = '';
  let timestamp = '';
  let content = '';
  
  // Find sender pattern
  const senderPattern = /([A-Za-z]+,\s*[A-Za-z]+\s*(?:\[[^\]]+\])?)/g;
  const senderMatches = fullText.match(senderPattern);
  if (senderMatches && senderMatches.length > 0) {
    sender = senderMatches[0]; // Use first occurrence
  }
  
  // Find timestamp pattern
  const timePattern = /(\d{1,2}:\d{2}\s*(?:AM|PM))/;
  const timeMatch = fullText.match(timePattern);
  if (timeMatch) {
    timestamp = timeMatch[1];
  }
  
  // Extract content after the timestamp
  if (timestamp) {
    const timeIndex = fullText.indexOf(timestamp);
    if (timeIndex !== -1) {
      content = fullText.substring(timeIndex + timestamp.length).trim();
    }
  }
  
  // If no content found after timestamp, try to extract from the beginning
  if (!content) {
    const byIndex = fullText.indexOf(' by ');
    if (byIndex !== -1) {
      content = fullText.substring(0, byIndex).trim();
    } else {
      content = fullText;
    }
  }
  
  // Clean up content - remove any remaining sender references and like reactions
  if (sender && content.includes(sender)) {
    content = content.replace(new RegExp(sender.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), '').trim();
  }
  
  // Extract and format reactions
  const { cleanedContent, reactions } = extractReactions(content);
  
  return {
    sender: sender || 'Unknown sender',
    content: cleanedContent || 'No content',
    reactions: domReactions.length > 0 ? domReactions : reactions,
    media: media,
    timestamp: timestamp || 'Unknown time',
    type: 'message'
  };
}

function extractMessagesFromText(text) {
  if (!text) {
    console.log('No text content provided');
    return [];
  }

  const messages = [];
  const lines = text.split('\n') || [];
  
  // Improved patterns for Teams message extraction
  const senderNamePattern = /^([A-Za-z]+,\s*[A-Za-z]+\s*(?:\[[^\]]+\])?)/; // "Lastname, Firstname [NH]"
  const timePattern = /(\d{1,2}:\d{2}\s*(?:AM|PM))/i;
  const datePattern = /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:,?\s+\d{4})?$|^\d{1,2}\/\d{1,2}\/\d{4}$|^(Today|Yesterday)$/i;
  
  let currentMessage = '';
  let currentSender = '';
  let currentTimestamp = '';
  let currentDate = '';
  let inMessage = false;
  
  // UI elements to skip
  const skipPatterns = [
    'unread', 'meeting', 'recording', 'like reaction', 'user added', 'edited',
    'begin reference', 'has context menu', 'calendar', 'training session',
    'link party', 'component', 'sharepoint', 'https://', 'you:'
  ];
  
  try {
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i]?.trim() || '';
      
      // Skip empty lines or very short lines
      if (!line || line.length < 2) continue;
      
      // Check for date pattern first
      const dateMatch = line.match(datePattern);
      const weekdayPattern = /^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$/i;
      const weekdayMatch = line.match(weekdayPattern);
      
      if (dateMatch) {
        currentDate = normalizeDate(line);
        continue;
      } else if (weekdayMatch) {
        currentDate = convertWeekdayToDate(line);
        continue;
      }
      
      // Skip obvious UI elements and system messages
      if (skipPatterns.some(pattern => line.toLowerCase().includes(pattern))) {
        continue;
      }
      
      // Check for sender name pattern (Teams format: "Lastname, Firstname [Organization]")
      const senderMatch = line.match(senderNamePattern);
      if (senderMatch && line.length < 50) { // Sender lines are typically short
        // Save previous message if exists
        if (currentMessage && currentSender) {
          let timestamp = currentTimestamp || 'Unknown time';
          if (currentDate && timestamp !== 'Unknown time') {
            timestamp = `${currentDate}, ${timestamp}`;
          }
          
          const { cleanedContent, reactions } = extractReactions(currentMessage.trim());
          messages.push({
            content: cleanedContent,
            sender: currentSender,
            reactions: reactions,
            media: [], // Text extraction doesn't have access to DOM media
            timestamp: timestamp,
            type: 'message'
          });
        }
        
        // Start new message
        currentSender = senderMatch[1];
        currentMessage = '';
        currentTimestamp = '';
        inMessage = true;
        continue;
      }
      
      // Check for timestamp
      const timeMatch = line.match(timePattern);
      if (timeMatch && line.length < 20) { // Timestamp lines are short
        currentTimestamp = timeMatch[1];
        continue;
      }
      
      // Skip lines that look like partial UI elements or navigation
      if (line.includes('…') || line.length > 200 || 
          line.toLowerCase().includes('all the coding font') ||
          line.includes('has context menu')) {
        continue;
      }
      
      // If we're in a message and this looks like message content
      if (inMessage && currentSender) {
        // Skip if line is just the sender name again
        if (line !== currentSender && !line.includes(currentSender)) {
          if (currentMessage) {
            currentMessage += '\n' + line;
          } else {
            currentMessage = line;
          }
        }
      }
    }
    
    // Add the last message if exists
    if (currentMessage && currentSender) {
      let timestamp = currentTimestamp || 'Unknown time';
      if (currentDate && timestamp !== 'Unknown time') {
        timestamp = `${currentDate}, ${timestamp}`;
      }
      
      const { cleanedContent, reactions } = extractReactions(currentMessage.trim());
      messages.push({
        content: cleanedContent,
        sender: currentSender,
        reactions: reactions,
        media: [], // Text extraction doesn't have access to DOM media
        timestamp: timestamp,
        type: 'message'
      });
    }
  } catch (error) {
    console.error('Error processing messages:', error);
    return [];
  }
  
  // Filter and clean messages
  return messages.filter(msg => {
    if (!msg || !msg.content || !msg.sender) return false;
    
    const content = msg.content.toLowerCase();
    const sender = msg.sender.toLowerCase();
    
    // Filter out very short messages, UI elements, and invalid senders
    return msg.content.length > 5 && 
           msg.sender.length > 3 &&
           !skipPatterns.some(pattern => content.includes(pattern) || sender.includes(pattern)) &&
           !sender.includes('all the coding') &&
           !content.includes('has context menu');
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