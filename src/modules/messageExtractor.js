/**
 * Message Extractor Module
 * Handles message extraction and data processing
 */

import { TeamsVariantDetector } from './teamsVariantDetector.js';

export class MessageExtractor {
  constructor() {
    // No settings needed for pure extraction
  }

  /**
   * Extracts detailed information for all visible messages in the chat pane.
   */
  extractVisibleMessages() {
    const extractedMessages = [];
    
    // Try multiple message container selectors for different Teams variants
    const messageSelectors = [
      '[data-tid="chat-pane-item"]',      // Classic/current
      '.fui-unstable-ChatItem',          // New Teams
      '.fui-ChatMessage',                // Fluent UI received messages
      '.fui-ChatMyMessage',              // Fluent UI sent messages
      '[data-tid="message"]',            // Generic
      '[role="article"]',                // ARIA role for messages
      '.message-container'               // Generic message container
    ];
    
    let messageContainers = [];
    
    for (const selector of messageSelectors) {
      const containers = document.querySelectorAll(selector);
      if (containers.length > 0) {
        messageContainers = containers;
        console.log(`📨 Using message selector "${selector}" - found ${containers.length} messages`);
        break;
      }
    }
    
    if (messageContainers.length === 0) {
      console.log('❌ No message containers found');
      return [];
    }
    
    messageContainers.forEach((messageContainer, index) => {
      try {
        const messageData = this.extractMessageData(messageContainer);
        if (messageData) {
          extractedMessages.push(messageData);
        }
      } catch (error) {
        console.log(`⚠️ Error extracting message ${index}:`, error);
      }
    });
    
    console.log(`✅ Extracted ${extractedMessages.length} messages`);
    return extractedMessages;
  }

  /**
   * Extracts data from a single message container
   */
  extractMessageData(messageContainer) {
    // Extract author
    const authorSelectors = [
      '[data-tid="message-author-name"]',
      '.fui-ChatMessage__author',
      '.message-author',
      '.sender-name',
      '[aria-label*="said"]',
      '.author-name'
    ];
    
    let author = 'Unknown';
    for (const selector of authorSelectors) {
      const authorElement = messageContainer.querySelector(selector);
      if (authorElement && authorElement.textContent.trim()) {
        author = authorElement.textContent.trim();
        break;
      }
    }
    
    // Extract timestamp
    const timestampSelectors = [
      '[data-tid="message-timestamp"]',
      '.fui-ChatMessage__timestamp',
      '.message-timestamp',
      'time',
      '[title*=":"]',
      '.timestamp'
    ];
    
    let timestamp = '';
    for (const selector of timestampSelectors) {
      const timestampElement = messageContainer.querySelector(selector);
      if (timestampElement) {
        timestamp = timestampElement.textContent?.trim() || timestampElement.title || '';
        if (timestamp) break;
      }
    }
    
    // Convert relative timestamps to actual dates
    timestamp = this.convertRelativeTimestamp(timestamp);
    
    // Extract message content
    const contentSelectors = [
      '[data-tid="message-body"]',
      '.fui-ChatMessage__content',
      '.message-content',
      '.message-body-content',
      '[data-tid="chat-pane-message"]',
      '.content'
    ];
    
    let content = '';
    for (const selector of contentSelectors) {
      const contentElement = messageContainer.querySelector(selector);
      if (contentElement) {
        // Get both text and HTML content for rich formatting
        content = this.extractRichContent(contentElement);
        if (content.trim()) break;
      }
    }
    
    // If no content found, try getting all text content as fallback
    if (!content.trim()) {
      content = messageContainer.textContent?.trim() || '';
      // Remove author and timestamp from content if they're included
      if (author !== 'Unknown') {
        content = content.replace(author, '').trim();
      }
      if (timestamp) {
        content = content.replace(timestamp, '').trim();
      }
    }
    
    // Extract attachments
    const attachments = this.extractAttachments(messageContainer);
    
    return {
      author,
      timestamp,
      content,
      attachments,
      element: messageContainer
    };
  }

  /**
   * Extracts rich content including formatting
   */
  extractRichContent(contentElement) {
    if (!contentElement) return '';
    
    // Clone the element to avoid modifying the original
    const clone = contentElement.cloneNode(true);
    
    // Remove script tags and other unwanted elements
    const unwantedElements = clone.querySelectorAll('script, style, .timestamp, .author-name');
    unwantedElements.forEach(el => el.remove());
    
    // Get HTML content with basic formatting preserved
    let htmlContent = clone.innerHTML;
    
    // Clean up the HTML
    htmlContent = htmlContent
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/div>/gi, '\n')
      .replace(/<\/p>/gi, '\n\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    
    return htmlContent;
  }

  /**
   * Extracts attachments from a message
   */
  extractAttachments(messageContainer) {
    const attachments = [];
    
    // Look for various attachment types
    const attachmentSelectors = [
      'img[src]',
      'a[href]',
      '[data-tid="file-attachment"]',
      '.attachment',
      '.file-attachment'
    ];
    
    for (const selector of attachmentSelectors) {
      const elements = messageContainer.querySelectorAll(selector);
      elements.forEach(element => {
        if (element.tagName === 'IMG') {
          attachments.push({
            type: 'image',
            src: element.src,
            alt: element.alt || 'Image'
          });
        } else if (element.tagName === 'A') {
          attachments.push({
            type: 'link',
            href: element.href,
            text: element.textContent?.trim() || 'Link'
          });
        } else {
          attachments.push({
            type: 'file',
            text: element.textContent?.trim() || 'File attachment'
          });
        }
      });
    }
    
    return attachments;
  }

  /**
   * Converts relative timestamps to actual dates
   */
  convertRelativeTimestamp(timestamp) {
    if (!timestamp || timestamp.trim() === '') return timestamp;
    
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);
    
    let cleanTimestamp = timestamp.trim();
    
    // Handle "Today" timestamps
    if (cleanTimestamp.toLowerCase().startsWith('today')) {
      const timeMatch = cleanTimestamp.match(/(\d{1,2}:\d{2}(?:\s*[ap]m)?)/i);
      if (timeMatch) {
        const timeStr = timeMatch[1];
        const dateStr = today.toLocaleDateString('en-US', { 
          year: 'numeric', 
          month: 'long', 
          day: 'numeric' 
        });
        return `${dateStr} ${timeStr}`;
      }
      return today.toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      });
    }
    
    // Handle "Yesterday" timestamps
    if (cleanTimestamp.toLowerCase().startsWith('yesterday')) {
      const timeMatch = cleanTimestamp.match(/(\d{1,2}:\d{2}(?:\s*[ap]m)?)/i);
      if (timeMatch) {
        const timeStr = timeMatch[1];
        const dateStr = yesterday.toLocaleDateString('en-US', { 
          year: 'numeric', 
          month: 'long', 
          day: 'numeric' 
        });
        return `${dateStr} ${timeStr}`;
      }
      return yesterday.toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      });
    }
    
    // Handle relative time like "2 hours ago", "5 minutes ago"
    const relativeMatch = cleanTimestamp.match(/(\d+)\s*(minute|hour|day)s?\s*ago/i);
    if (relativeMatch) {
      const amount = parseInt(relativeMatch[1]);
      const unit = relativeMatch[2].toLowerCase();
      
      let targetDate = new Date(now);
      if (unit === 'minute') {
        targetDate.setMinutes(targetDate.getMinutes() - amount);
      } else if (unit === 'hour') {
        targetDate.setHours(targetDate.getHours() - amount);
      } else if (unit === 'day') {
        targetDate.setDate(targetDate.getDate() - amount);
      }
      
      return targetDate.toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    }
    
    // If it's already a formatted date, return as-is
    return cleanTimestamp;
  }
}