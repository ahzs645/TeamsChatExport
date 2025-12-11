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
   * Extracts reactions from a message container
   */
  extractReactions(messageContainer) {
    const reactions = [];
    const pills = messageContainer.querySelectorAll('[data-tid="diverse-reaction-pill-button"]');

    pills.forEach(pill => {
      const emojiImg = pill.querySelector('[data-tid="emoticon-renderer"] img');
      const emoji = emojiImg?.getAttribute('alt') || '';

      // Parse aria-label for count
      const label = pill.getAttribute('aria-label') || '';
      const countMatch = label.match(/(\d+)/);
      const count = countMatch ? parseInt(countMatch[1], 10) : 1;

      // Extract reactor names
      const reactors = this.extractReactors(pill);

      if (emoji || count > 0) {
        reactions.push({ emoji, count, reactors });
      }
    });

    return reactions;
  }

  /**
   * Extracts reactor names from a reaction pill
   */
  extractReactors(pill) {
    const labelledBy = pill.getAttribute('aria-labelledby') || '';
    const ids = labelledBy.split(/\s+/).filter(Boolean);
    const names = ids.map(id => {
      const el = document.getElementById(id);
      return el?.innerText?.trim() || '';
    }).filter(Boolean);

    // Parse "User1, User2, and 3 others reacted" format
    const text = names.join(' ');
    return text.split(/react/i)[0]
      .split(/,\s*|\s+and\s+/)
      .map(n => n.trim())
      .filter(n => n && !/^\d+\s+others?$/i.test(n) && !/others$/i.test(n))
      .slice(0, 20);
  }

  /**
   * Extracts reply/quote information from a message
   */
  extractReplyTo(messageContainer) {
    // Method 1: quoted-reply-card (most common)
    const replyCard = messageContainer.querySelector('[data-tid="quoted-reply-card"]');
    if (replyCard) {
      const timestampEl = replyCard.querySelector('[data-tid="quoted-reply-timestamp"]');
      const authorEl = timestampEl?.previousElementSibling;
      const contentEl = replyCard.querySelector('[data-tid="quoted-reply-preview-content"]');

      const author = authorEl?.textContent?.trim() || '';
      const timestamp = timestampEl?.textContent?.trim() || '';
      const text = this.extractTextContent(contentEl) || '';

      if (author || text) {
        return { author, timestamp, text };
      }
    }

    // Method 2: Begin Reference aria-label
    const refGroup = messageContainer.querySelector('[role="group"][aria-label^="Begin Reference"]');
    if (refGroup) {
      const label = refGroup.getAttribute('aria-label') || '';
      const match = label.match(/Begin Reference,\s*([\s\S]*),\s*([^,]+),\s*([^,]+),\s*End reference/i);
      if (match) {
        return { text: match[1].trim(), author: match[2].trim(), timestamp: match[3].trim() };
      }
    }

    // Method 3: div[role="heading"] with Begin Reference pattern
    const heading = messageContainer.querySelector('div[role="heading"]');
    if (heading) {
      const headingText = heading.textContent?.trim() || '';
      const match = headingText.match(/Begin Reference,\s*(.*?)\s*by\s*(.+)$/i);
      if (match) {
        return { text: match[1].trim(), author: match[2].trim(), timestamp: '' };
      }
    }

    return null;
  }

  /**
   * Extracts text content from an element, handling nested elements
   */
  extractTextContent(element) {
    if (!element) return '';

    let text = '';
    const walk = (node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        text += node.nodeValue;
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;

      const tag = node.tagName;
      if (tag === 'BR') {
        text += '\n';
        return;
      }
      if (tag === 'IMG') {
        text += node.getAttribute('alt') || node.getAttribute('aria-label') || '';
        return;
      }

      const blockTags = /^(DIV|P|LI|BLOCKQUOTE)$/;
      const startLen = text.length;

      for (const child of node.childNodes) {
        walk(child);
      }

      if (blockTags.test(tag) && text.length > startLen) {
        text += '\n';
      }
    };

    walk(element);
    return text.replace(/\n{3,}/g, '\n\n').trim();
  }

  /**
   * Extracts edited status from a message
   */
  extractEditedStatus(messageContainer) {
    // Check for edited indicator by ID pattern
    const editedEl = messageContainer.querySelector('[id^="edited-"]');
    if (editedEl) {
      const text = (editedEl.textContent || editedEl.getAttribute('title') || '').trim();
      if (/^edited\b/i.test(text)) {
        return { edited: true, editedTimestamp: null };
      }
    }

    // Check aria-labelledby for edited reference
    const labelledBy = messageContainer.getAttribute('aria-labelledby') || '';
    const ids = labelledBy.split(/\s+/);
    for (const id of ids) {
      if (id.startsWith('edited-')) {
        const el = document.getElementById(id);
        if (el && /^edited\b/i.test(el.textContent || '')) {
          return { edited: true, editedTimestamp: null };
        }
      }
    }

    // Check for nested elements with edited pattern
    const allEdited = messageContainer.querySelectorAll('[id^="edited-"], [class*="edited"]');
    for (const el of allEdited) {
      const text = (el.textContent || el.getAttribute('title') || '').trim().toLowerCase();
      if (text.includes('edited')) {
        return { edited: true, editedTimestamp: null };
      }
    }

    return { edited: false, editedTimestamp: null };
  }

  /**
   * Extracts avatar URL from a message
   */
  extractAvatar(messageContainer) {
    const selectors = [
      '[data-tid="message-avatar"] img',
      '[data-tid="avatar"] img',
      '.fui-Avatar img',
      '.fui-Avatar__image',
      'img[src*="profilepicture"]',
      'img[src*="profilepicturev2"]'
    ];

    for (const sel of selectors) {
      const img = messageContainer.querySelector(sel);
      if (img?.src && (img.src.includes('/profilepicture') || img.src.includes('/profilepicturev2'))) {
        return img.src;
      }
    }
    return null;
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

    // Extract reactions
    const reactions = this.extractReactions(messageContainer);

    // Extract reply/quote info
    const replyTo = this.extractReplyTo(messageContainer);

    // Extract edited status
    const { edited, editedTimestamp } = this.extractEditedStatus(messageContainer);

    // Extract avatar
    const avatar = this.extractAvatar(messageContainer);

    const message = content;

    const isDayDivider = this.isLikelyDateDivider(message);

    if (isDayDivider) {
      return {
        author: '',
        timestamp: '',
        message,
        content: message,
        attachments,
        reactions: [],
        replyTo: null,
        edited: false,
        editedTimestamp: null,
        avatar: null,
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
      reactions,
      replyTo,
      edited,
      editedTimestamp,
      avatar,
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
   * Extracts attachments from a message with enhanced metadata
   */
  extractAttachments(contextElement) {
    if (!contextElement) {
      return [];
    }

    const seen = new Map();

    // File attachment containers
    const containerSelectors = [
      '[data-tid="file-attachment-grid"]',
      '[data-tid="file-preview-root"]',
      '[data-tid="attachments"]'
    ];

    const containers = containerSelectors
      .map(s => contextElement.querySelector(s))
      .filter(Boolean);

    // Also check the context element itself
    containers.push(contextElement);

    containers.forEach(container => {
      // File chiclets and attachment elements
      container.querySelectorAll('[data-testid="file-attachment"], [data-tid^="file-chiclet-"], [data-tid="file-attachment"]').forEach(el => {
        const title = el.getAttribute('title') || el.getAttribute('aria-label') || '';
        const parsed = this.parseAttachmentTitle(title);
        if (parsed) this.addAttachment(seen, parsed);

        // Also check for links within the attachment
        el.querySelectorAll('a[href^="http"]').forEach(link => {
          this.addAttachment(seen, {
            href: link.href,
            label: link.textContent?.trim() || link.href,
            type: 'link'
          });
        });
      });

      // Rich file preview buttons
      container.querySelectorAll('button[data-testid="rich-file-preview-button"][title]').forEach(btn => {
        const parsed = this.parseAttachmentTitle(btn.getAttribute('title'));
        if (parsed) this.addAttachment(seen, parsed);
      });
    });

    // Inline images (lazy-loaded)
    contextElement.querySelectorAll('[data-testid="lazy-image-wrapper"] img').forEach(img => {
      if (img.src?.startsWith('http')) {
        this.addAttachment(seen, {
          href: img.src,
          label: img.alt || 'image',
          type: 'IMAGE'
        });
      }
    });

    // Regular images (excluding avatars)
    contextElement.querySelectorAll('img[src^="http"]').forEach(img => {
      if (img.closest('[data-tid="message-avatar"]') || img.closest('.fui-Avatar')) return;
      if (img.classList.contains('fui-Avatar__image')) return;
      if (img.src.includes('/profilepicture')) return;

      // Skip emoticons
      if (img.closest('[data-tid="emoticon-renderer"]')) return;

      this.addAttachment(seen, {
        href: img.src,
        label: img.alt || 'Image',
        type: 'IMAGE'
      });
    });

    // Safe links and regular links
    contextElement.querySelectorAll('a[data-testid="atp-safelink"], a[href^="http"]').forEach(link => {
      // Skip avatar links
      if (link.closest('[data-tid="message-avatar"]')) return;
      if (link.closest('.fui-Avatar')) return;

      // Skip javascript and hash links
      if (!link.href || link.href.startsWith('javascript:') || link.href === '#') return;

      this.addAttachment(seen, {
        href: link.href,
        label: link.textContent?.trim() || link.href,
        type: 'link'
      });
    });

    return Array.from(seen.values());
  }

  /**
   * Parses attachment metadata from title/aria-label attributes
   */
  parseAttachmentTitle(title) {
    if (!title) return null;

    const lines = title.split(/\n+/).map(l => l.trim()).filter(Boolean);
    const hrefMatch = title.match(/https?:\/\/\S+/);

    const attachment = {
      label: lines[0] || '',
      href: hrefMatch?.[0] || '',
      metaText: lines.slice(1).join(' • ')
    };

    // Extract file type from extension
    const extMatch = attachment.label.match(/\.([A-Za-z0-9]{1,6})$/);
    if (extMatch) {
      attachment.type = extMatch[1].toUpperCase();
    }

    // Extract size from metadata
    const sizeMatch = attachment.metaText.match(/\b\d+(?:[.,]\d+)?\s*(?:bytes?|KB|MB|GB|TB)\b/i);
    if (sizeMatch) {
      attachment.size = sizeMatch[0].replace(',', '.').trim();
    }

    // Extract owner from metadata
    const ownerMatch = attachment.metaText.match(/(?:Shared by|Uploaded by|Sent by|From|Owner)\s*:?\s*([^•,]+)/i);
    if (ownerMatch) {
      attachment.owner = ownerMatch[1].trim();
    }

    return attachment;
  }

  /**
   * Adds an attachment to the seen map, merging duplicates
   */
  addAttachment(seen, attachment) {
    const key = `${attachment.href || ''}@@${attachment.label || ''}`;
    const existing = seen.get(key);

    if (existing) {
      // Merge metadata, preferring existing values
      if (!existing.type && attachment.type) existing.type = attachment.type;
      if (!existing.size && attachment.size) existing.size = attachment.size;
      if (!existing.owner && attachment.owner) existing.owner = attachment.owner;
      if (!existing.metaText && attachment.metaText) existing.metaText = attachment.metaText;
      if (!existing.href && attachment.href) existing.href = attachment.href;
    } else {
      seen.set(key, attachment);
    }
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
