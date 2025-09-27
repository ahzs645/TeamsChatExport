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
    
    let dividerCount = 0;
    let systemCount = 0;

    messageContainers.forEach((messageContainer, index) => {
      try {
        const messageData = this.extractMessageData(messageContainer);
        if (messageData) {
          if (messageData.type === 'divider') {
            dividerCount += 1;
          } else if (messageData.type === 'system') {
            systemCount += 1;
          }
          messageData.__sequence = window.__teamsExtractorMessageSequence = (window.__teamsExtractorMessageSequence || 0) + 1;
          extractedMessages.push(messageData);
        }
      } catch (error) {
        console.log(`⚠️ Error extracting message ${index}:`, error);
      }
    });

    if (dividerCount > 0 || systemCount > 0) {
      console.log(`✅ Extracted ${extractedMessages.length} messages (including ${dividerCount} dividers, ${systemCount} system notifications)`);
    } else {
      console.log(`✅ Extracted ${extractedMessages.length} messages`);
    }

    return this.prepareMessages(extractedMessages);
  }

  /**
   * Extracts data from a single message container
  */
  extractMessageData(messageContainer) {
    const messageId = messageContainer.getAttribute('data-mid') || messageContainer.dataset?.mid || messageContainer.getAttribute('id');

    const dividerSelectors = [
      '.fui-Divider__wrapper',
      '[data-tid="message-date-separator"]',
      '[data-tid="message-day-divider"]',
      '.message-date-divider',
      '.message-day-divider',
      '.ui-chat__message__date'
    ];

    for (const selector of dividerSelectors) {
      const dividerElement = messageContainer.querySelector(selector);
      const dividerText = dividerElement?.textContent?.trim();
      if (dividerText && this.isLikelyDateDivider(dividerText)) {
        return {
          author: '',
          timestamp: '',
          message: dividerText,
          content: dividerText,
          attachments: [],
          type: 'divider',
          id: messageId
        };
      }
    }

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
    
    // Extract message content
    const contentSelectors = [
      '[id^="content-"]',
      '.fui-ChatMessageCompact__content',
      '.fui-ChatMessage__content',
      '.fui-ChatMessageContent',
      '.ui-chat__message__content',
      '[data-tid="message-body"]',
      '.message-content',
      '.message-body-content',
      '[data-tid="chat-pane-message"]',
      '.content'
    ];
    
    let content = '';
    let contentElement = null;
    for (const selector of contentSelectors) {
      const candidate = messageContainer.querySelector(selector);
      if (candidate) {
        contentElement = candidate;
        content = this.extractRichContent(candidate);
        if (content.trim()) break;
      }
    }

    // If no content found, try getting all text content as fallback
    if (!content.trim()) {
      content = messageContainer.textContent || '';
    }

    content = this.cleanMessageText(content, author, timestamp);

    // Extract attachments
    const attachmentContext = contentElement || messageContainer;
    const attachments = this.extractAttachments(attachmentContext);

    const message = content;

    const isDayDivider = this.isLikelyDateDivider(message);

    if (isDayDivider) {
      return {
        author: '',
        timestamp: '',
        message,
        content: message,
        attachments,
        type: 'divider',
        id: messageId
      };
    }

    if (!message && attachments.length === 0) {
      return null;
    }

    let derivedType = null;

    if (author === 'Unknown' && !timestamp) {
      derivedType = 'system';
    }

    if (derivedType === 'system' && !message) {
      return null;
    }

    return {
      author,
      timestamp,
      message,
      content,
      attachments,
      type: derivedType,
      id: messageId
    };
  }

  isLikelyDateDivider(label) {
    if (!label) return false;
    const text = label.trim();
    if (!text) return false;
    const lower = text.toLowerCase();
    if (/^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)/i.test(text)) {
      return true;
    }
    if (/^(today|yesterday|tomorrow)\b/i.test(text)) {
      return true;
    }
    const monthPattern = '(january|february|march|april|may|june|july|august|september|october|november|december)';
    if (new RegExp(`^${monthPattern}\\s+\\d{1,2}(?:,\\s*\\d{4})?$`, 'i').test(text)) {
      return true;
    }
    if (new RegExp(`^\\d{1,2}\\s+${monthPattern}(?:,\\s*\\d{4})?$`, 'i').test(text)) {
      return true;
    }
    if (/^\\d{4}-\\d{1,2}-\\d{1,2}$/.test(lower)) {
      return true;
    }
    if (/^\\d{1,2}[\/.-]\\d{1,2}[\/.-]\\d{2,4}$/.test(lower)) {
      return true;
    }
    return false;
  }

  prepareMessages(messages) {
    const prepared = [];
    let currentDateContext = null;

    messages.forEach((msg) => {
      if (!msg) return;

      if (msg.type === 'divider') {
        const dividerDate = this.parseDividerDate(msg.message || msg.content || '', currentDateContext);
        if (dividerDate) {
          currentDateContext = dividerDate;
          msg.isoTimestamp = dividerDate.toISOString();
          msg.timestamp = this.formatDateForDisplay(dividerDate, { includeTime: false });
        }
        prepared.push(msg);
        return;
      }

      const { display, iso } = this.normalizeTimestamp(msg.timestamp, currentDateContext);
      if (display) {
        msg.timestamp = display;
      }
      if (iso) {
        msg.isoTimestamp = iso;
        currentDateContext = new Date(iso);
      }

      prepared.push(msg);
    });

    return prepared;
  }

  normalizeTimestamp(rawTimestamp, contextDate) {
    const referenceDate = contextDate ? new Date(contextDate) : null;
    if (!rawTimestamp || !rawTimestamp.trim()) {
      if (referenceDate && !Number.isNaN(referenceDate.getTime())) {
        return {
          display: this.formatDateForDisplay(referenceDate),
          iso: referenceDate.toISOString()
        };
      }
      return { display: rawTimestamp || '', iso: null };
    }

    const timestampDate = this.convertRelativeTimestamp(rawTimestamp, referenceDate);
    if (timestampDate && !Number.isNaN(timestampDate.getTime())) {
      return {
        display: this.formatDateForDisplay(timestampDate),
        iso: timestampDate.toISOString()
      };
    }

    return { display: rawTimestamp, iso: null };
  }

  parseDividerDate(label, previousDate) {
    if (!label) return previousDate || null;
    const cleaned = label.trim();

    let candidate = cleaned;
    if (!/\d{4}/.test(cleaned)) {
      const inferredYear = this.inferYearFromContext(cleaned, previousDate);
      if (inferredYear) {
        candidate = `${cleaned}, ${inferredYear}`;
      }
    }

    let parsed = Date.parse(candidate);
    if (Number.isNaN(parsed) && /\sat\s/i.test(candidate)) {
      parsed = Date.parse(candidate.replace(/\sat\s/i, ' '));
    }

    if (!Number.isNaN(parsed)) {
      return new Date(parsed);
    }

    return previousDate ? new Date(previousDate) : null;
  }

  inferYearFromContext(label, previousDate) {
    const monthMatch = label.toLowerCase().match(/(january|february|march|april|may|june|july|august|september|october|november|december)/);
    const monthIndex = monthMatch ? this.getMonthIndex(monthMatch[1]) : null;

    if (previousDate && monthIndex !== null && monthIndex >= 0) {
      const prevMonth = previousDate.getMonth();
      const prevYear = previousDate.getFullYear();

      if (monthIndex < prevMonth - 6) {
        return prevYear + 1;
      }
      if (monthIndex > prevMonth + 6) {
        return prevYear - 1;
      }
      return prevYear;
    }

    const currentYear = previousDate ? new Date(previousDate).getFullYear() : new Date().getFullYear();
    return currentYear;
  }

  getMonthIndex(name) {
    if (!name) return null;
    const months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
    return months.indexOf(name.toLowerCase());
  }

  formatDateForDisplay(dateObj, options = { includeTime: true }) {
    if (!(dateObj instanceof Date) || Number.isNaN(dateObj.getTime())) return '';
    const withTime = {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      hour: 'numeric',
      minute: '2-digit'
    };
    const withoutTime = {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    };
    return dateObj.toLocaleString(undefined, options.includeTime === false ? withoutTime : withTime);
  }

  /**
   * Extracts rich content including formatting
   */
  extractRichContent(contentElement) {
    if (!contentElement) return '';

    // Prefer inner content nodes so we don't capture author/timestamp wrappers
    const preferredNode = contentElement.querySelector('[id^="content-"]') ||
      contentElement.querySelector('.fui-ChatMessageCompact__content') ||
      contentElement.querySelector('.fui-ChatMessageContent') ||
      contentElement.querySelector('.ui-chat__message__content');

    const clone = (preferredNode || contentElement).cloneNode(true);

    const selectorsToRemove = [
      'script',
      'style',
      '.timestamp',
      '.author-name',
      '[data-tid="message-author-name"]',
      '.fui-ChatMessageCompact__author',
      '.fui-ChatMessage__author',
      '.fui-ChatMessageCompact__timestamp',
      '.fui-ChatMessage__timestamp',
      'time',
      'button',
      '.fui-ChatMessageCompact__informationRow',
      '.fui-ChatMessageCompact__status',
      '.fui-ChatMessageCompact__header',
      '.fui-ChatMessageCompact__reactions',
      '.fui-Tooltip'
    ];

    clone.querySelectorAll(selectorsToRemove.join(',')).forEach((el) => el.remove());

    return this.cleanMessageText(clone.textContent || '', '', '');
  }

  /**
   * Cleans message text by decoding HTML entities, stripping authors/timestamps and normalising whitespace
   */
  cleanMessageText(text, author, timestamp) {
    if (!text) return '';

    let cleaned = this.decodeHtmlEntities(text)
      .replace(/\u00a0/g, ' ');

    if (author && author !== 'Unknown') {
      const authorPattern = new RegExp(`^\s*${this.escapeRegExp(author)}\s*[:\-]?`, 'i');
      cleaned = cleaned.replace(authorPattern, ' ');
    }

    if (timestamp) {
      const timestampPattern = new RegExp(this.escapeRegExp(timestamp), 'gi');
      cleaned = cleaned.replace(timestampPattern, ' ');
    }

    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    return cleaned;
  }

  /**
   * Decodes basic HTML entities using a temporary textarea
   */
  decodeHtmlEntities(text) {
    if (!text) return '';
    const textarea = document.createElement('textarea');
    textarea.innerHTML = text;
    return textarea.value;
  }

  /**
   * Escapes text for safe use in a RegExp
   */
  escapeRegExp(value) {
    return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /**
   * Extracts attachments from a message
   */
  extractAttachments(contextElement) {
    const attachments = [];

    // Look for various attachment types
    const attachmentSelectors = [
      'img[src]',
      'a[href]',
      '[data-tid="file-attachment"]',
      '.attachment',
      '.file-attachment'
    ];

    if (!contextElement) {
      return attachments;
    }

    for (const selector of attachmentSelectors) {
      const elements = contextElement.querySelectorAll(selector);
      elements.forEach(element => {
        if (element.closest('[data-tid="message-avatar"]') || element.closest('.fui-Avatar')) {
          return;
        }
        if (element.tagName === 'IMG') {
          if (element.classList.contains('fui-Avatar__image')) {
            return;
          }
          attachments.push({
            type: 'image',
            src: element.src,
            alt: element.alt || 'Image'
          });
        } else if (element.tagName === 'A') {
          if (!element.href || element.href.startsWith('javascript:') || element.href === '#') {
            return;
          }
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

  extractFromHTML(htmlString) {
    if (!htmlString) return null;
    const wrapper = document.createElement('div');
    wrapper.innerHTML = htmlString;
    const node = wrapper.firstElementChild;
    if (!node) return null;
    return this.extractMessageData(node);
  }

  mergeMessages(messages) {
    const seen = new Map();
    messages.forEach((msg) => {
      if (!msg) return;
      const key = this.getMessageKey(msg);
      const existing = seen.get(key);
      if (!existing || this.compareMessages(msg, existing) < 0) {
        seen.set(key, msg);
      }
    });
    const merged = Array.from(seen.values()).sort((a, b) => this.compareMessages(a, b));
    merged.forEach((msg) => delete msg.__sequence);
    return merged;
  }

  getMessageKey(msg) {
    if (msg.id) {
      return `id:${msg.id}`;
    }
    return `${msg.isoTimestamp || msg.timestamp || ''}::${msg.author || ''}::${msg.message || msg.content || ''}`;
  }

  compareMessages(a, b) {
    const aIso = a.isoTimestamp;
    const bIso = b.isoTimestamp;
    if (aIso && bIso) {
      const diff = new Date(aIso) - new Date(bIso);
      if (diff !== 0) {
        return diff;
      }
    }
    if (aIso && !bIso) return -1;
    if (!aIso && bIso) return 1;
    const seqA = a.__sequence ?? Infinity;
    const seqB = b.__sequence ?? Infinity;
    if (seqA !== seqB) {
      return seqA - seqB;
    }
    return (a.timestamp || '').localeCompare(b.timestamp || '');
  }

  /**
   * Converts relative timestamps to actual Date objects
   */
  convertRelativeTimestamp(timestamp, referenceDate = null) {
    if (!timestamp || !timestamp.trim()) {
      return referenceDate ? new Date(referenceDate) : null;
    }

    const cleanTimestamp = timestamp.trim();
    const variations = [cleanTimestamp];

    if (/\sat\s/i.test(cleanTimestamp)) {
      variations.push(cleanTimestamp.replace(/\sat\s/i, ' '));
    }

    const slashMatch = cleanTimestamp.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}:\d{2}(?:\s*[ap]m)?))?/i);
    if (slashMatch) {
      const month = parseInt(slashMatch[1], 10) - 1;
      const day = parseInt(slashMatch[2], 10);
      let year = parseInt(slashMatch[3], 10);
      if (year < 100) {
        year += year >= 70 ? 1900 : 2000;
      }
      let hours = 0;
      let minutes = 0;
      if (slashMatch[4]) {
        const timeMatch = slashMatch[4].match(/(\d{1,2}):(\d{2})(?:\s*([ap]m))?/i);
        if (timeMatch) {
          hours = parseInt(timeMatch[1], 10);
          minutes = parseInt(timeMatch[2], 10);
          const meridiem = timeMatch[3];
          if (meridiem) {
            const lowerMeridiem = meridiem.toLowerCase();
            if (lowerMeridiem === 'pm' && hours < 12) {
              hours += 12;
            }
            if (lowerMeridiem === 'am' && hours === 12) {
              hours = 0;
            }
          }
        }
      }
      return new Date(year, month, day, hours, minutes, 0, 0);
    }

    for (const candidate of variations) {
      const parsed = Date.parse(candidate);
      if (!Number.isNaN(parsed)) {
        return new Date(parsed);
      }
    }

    const monthDayMatch = cleanTimestamp.match(/(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2})(?:,?\s*(\d{4}))?(?:\s+(\d{1,2}:\d{2}(?:\s*[ap]m)?))?/i);
    if (monthDayMatch) {
      const monthIndex = this.getMonthIndex(monthDayMatch[1]);
      const day = parseInt(monthDayMatch[2], 10);
      let year = monthDayMatch[3] ? parseInt(monthDayMatch[3], 10) : null;
      if (!year && referenceDate) {
        year = this.inferYearFromContext(monthDayMatch[0], referenceDate);
      }
      if (!year) {
        year = new Date().getFullYear();
      }
      if (monthIndex !== null && day) {
        const candidateDate = new Date(year, monthIndex, day);
        return this.applyTimeToDate(candidateDate, monthDayMatch[4]);
      }
    }

    const baseNow = new Date();
    const baseReference = referenceDate ? new Date(referenceDate) : new Date(baseNow);
    const lower = cleanTimestamp.toLowerCase();

    if (lower.startsWith('today')) {
      return this.applyTimeToDate(new Date(baseReference), cleanTimestamp);
    }

    if (lower.startsWith('yesterday')) {
      const target = new Date(baseReference);
      target.setDate(target.getDate() - 1);
      return this.applyTimeToDate(target, cleanTimestamp);
    }

    const relativeMatch = lower.match(/(\d+)\s*(minute|hour|day)s?\s*ago/);
    if (relativeMatch) {
      const amount = parseInt(relativeMatch[1], 10);
      const unit = relativeMatch[2];
      const target = new Date(baseNow);
      if (unit === 'minute') {
        target.setMinutes(target.getMinutes() - amount);
      } else if (unit === 'hour') {
        target.setHours(target.getHours() - amount);
      } else if (unit === 'day') {
        target.setDate(target.getDate() - amount);
      }
      return target;
    }

    const weekdayMatch = cleanTimestamp.match(/^(Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday)(?:,?\s+(\d{1,2}:\d{2}(?:\s*[ap]m)?))?/i);
    if (weekdayMatch) {
      const targetIndex = this.getWeekdayIndex(weekdayMatch[1]);
      const base = referenceDate ? new Date(referenceDate) : new Date(baseNow);
      let diff = (base.getDay() - targetIndex + 7) % 7;
      if (!referenceDate && diff === 0 && this.compareTimes(weekdayMatch[2], base) > 0) {
        diff = 7;
      }
      const targetDate = new Date(base);
      targetDate.setDate(targetDate.getDate() - diff);
      return this.applyTimeToDate(targetDate, weekdayMatch[2] || cleanTimestamp);
    }

    const timeOnlyMatch = cleanTimestamp.match(/^(\d{1,2}:\d{2}(?:\s*[ap]m)?)$/i);
    if (timeOnlyMatch) {
      const target = referenceDate ? new Date(referenceDate) : new Date(baseReference);
      return this.applyTimeToDate(target, timeOnlyMatch[1]);
    }

    return referenceDate ? new Date(referenceDate) : null;
  }

  applyTimeToDate(dateObj, timeString) {
    const result = new Date(dateObj);
    if (!timeString) return result;
    const timeMatch = timeString.match(/(\d{1,2}):(\d{2})(?:\s*([ap]m))?/i);
    if (timeMatch) {
      let hours = parseInt(timeMatch[1], 10);
      const minutes = parseInt(timeMatch[2], 10);
      const meridiem = timeMatch[3];
      if (meridiem) {
        const lowerMeridiem = meridiem.toLowerCase();
        if (lowerMeridiem === 'pm' && hours < 12) {
          hours += 12;
        }
        if (lowerMeridiem === 'am' && hours === 12) {
          hours = 0;
        }
      }
      result.setHours(hours, minutes, 0, 0);
    }
    return result;
  }

  getWeekdayIndex(name) {
    const weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    return weekdays.findIndex((day) => day.toLowerCase() === name.toLowerCase());
  }

  compareTimes(timeString, baseDate) {
    if (!timeString) return -1;
    const target = this.applyTimeToDate(new Date(baseDate), timeString);
    return target - baseDate;
  }
}
