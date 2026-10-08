/*
 * Teams Chat Exporter — results viewer.
 *
 * Storage contract (chrome.storage.local):
 *   savedExtractions      { "[<saved time>] <chat name>": Message[] }   (key format kept for older exports)
 *   savedExtractionsMeta  { <same key>: { name, source: 'extract'|'upload', savedAt, fileName?, originalExtractedAt? } }
 *   teamsChatData         last extraction only; read as a fallback when savedExtractions is empty
 *   viewerSelf            { byConversation: { <chat name>: <author> }, last: <author> }
 *
 * The background page opens results.html?select=<key> after storing an extraction, so the
 * viewer only ever reads conversations from storage and never adds entries of its own.
 */

/**
 * Rendering helpers shared by the viewer and the exported HTML page.
 * This function is serialised with Function#toString into exports, so it must stay
 * self-contained (no references to anything outside its own body).
 */
function createViewerCore() {
  const HUE_COUNT = 8;

  const escapeHtml = (value) => String(value == null ? '' : value).replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  }[c]));

  // A saved key looks like "[10/8/2026, 11:28:21 AM] Chat name". Only bracket groups that contain a
  // digit count as our timestamp prefix, so a real chat called "[EXT] Vendor" keeps its name.
  const PREFIX_RE = /^\[([^\]]*\d[^\]]*)\]\s*([\s\S]+)$/;
  const parseConversationName = (fullName) => {
    const text = String(fullName == null ? '' : fullName);
    const match = text.match(PREFIX_RE);
    if (match) {
      return { cleanName: match[2], extractionTime: match[1], isTimestamped: true };
    }
    return { cleanName: text, extractionTime: null, isTimestamped: false };
  };

  // Removes every leading timestamp prefix ("[time] [time] Name" -> "Name") for re-uploaded exports.
  const stripPrefixes = (fullName) => {
    let current = parseConversationName(fullName);
    let firstTime = current.extractionTime;
    while (current.isTimestamped) {
      const next = parseConversationName(current.cleanName);
      if (!next.isTimestamped) break;
      firstTime = firstTime || next.extractionTime;
      current = next;
    }
    return { cleanName: current.cleanName, extractionTime: firstTime };
  };

  const hashString = (text) => {
    let hash = 5381;
    const value = String(text || '');
    for (let i = 0; i < value.length; i += 1) {
      hash = ((hash << 5) + hash + value.charCodeAt(i)) | 0;
    }
    return Math.abs(hash);
  };

  const hueFor = (name) => hashString(name) % HUE_COUNT;

  const initialsFor = (name) => {
    const words = String(name || '')
      .replace(/[^\p{L}\p{N}\s]/gu, ' ')
      .split(/\s+/)
      .filter(Boolean);
    if (words.length === 0) return '?';
    return words.slice(0, 2).map((word) => Array.from(word)[0]).join('').toUpperCase();
  };

  const pluralize = (count, singular, plural) => `${count} ${count === 1 ? singular : (plural || `${singular}s`)}`;

  const isRealMessage = (msg) => msg && msg.type !== 'divider' && msg.type !== 'system';

  const messageText = (msg) => String((msg && (msg.message || msg.content)) || '').trim();

  const getParticipants = (messages) => {
    const counts = new Map();
    (messages || []).forEach((msg) => {
      if (isRealMessage(msg) && msg.author && msg.author !== 'Unknown') {
        counts.set(msg.author, (counts.get(msg.author) || 0) + 1);
      }
    });
    return Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .map(([name, count]) => ({ name, count }));
  };

  const countMessages = (messages) => (messages || []).filter(isRealMessage).length;

  // Only let through link/image URLs that cannot run script.
  const safeUrl = (url, { image = false } = {}) => {
    if (!url) return null;
    try {
      const parsed = new URL(String(url), document.baseURI);
      const allowed = image ? ['http:', 'https:', 'data:', 'blob:'] : ['http:', 'https:', 'mailto:'];
      if (!allowed.includes(parsed.protocol)) return null;
      if (parsed.protocol === 'data:' && !/^data:image\//i.test(String(url))) return null;
      return parsed.href;
    } catch (_err) {
      return null;
    }
  };

  const el = (tag, className, text) => {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text != null) node.textContent = text;
    return node;
  };

  const fillAvatar = (node, name) => {
    node.textContent = initialsFor(name);
    node.dataset.hue = String(hueFor(name));
  };

  // Appends text to parent, wrapping case-insensitive matches of term in <mark>.
  const appendHighlighted = (parent, text, term) => {
    const value = String(text || '');
    if (!term) {
      parent.appendChild(document.createTextNode(value));
      return;
    }
    const lower = value.toLowerCase();
    const needle = term.toLowerCase();
    let index = 0;
    let hit = lower.indexOf(needle, index);
    while (hit !== -1) {
      if (hit > index) parent.appendChild(document.createTextNode(value.slice(index, hit)));
      parent.appendChild(el('mark', null, value.slice(hit, hit + needle.length)));
      index = hit + needle.length;
      hit = lower.indexOf(needle, index);
    }
    if (index < value.length) parent.appendChild(document.createTextNode(value.slice(index)));
  };

  // Full-size preview in an in-page dialog. Opening images in a new tab is avoided on purpose:
  // browsers block data: navigations, and a blob: copy of an SVG would run its scripts in our origin.
  const openImage = (src, alt) => {
    let dialog = document.getElementById('image-lightbox');
    if (!dialog) {
      dialog = el('dialog', 'lightbox');
      dialog.id = 'image-lightbox';
      dialog.setAttribute('aria-label', 'Image preview');
      const close = el('button', 'btn lightbox-close', 'Close');
      close.type = 'button';
      close.addEventListener('click', () => dialog.close());
      dialog.append(close, el('img', 'lightbox-img'));
      dialog.addEventListener('click', (event) => { if (event.target === dialog) dialog.close(); });
      document.body.appendChild(dialog);
    }
    const img = dialog.querySelector('img');
    img.src = src;
    img.alt = alt || '';
    dialog.showModal();
  };

  const buildAttachments = (attachments, term) => {
    const list = el('ul', 'message-attachments');
    list.setAttribute('aria-label', 'Attachments');
    attachments.forEach((att) => {
      if (!att) return;
      const item = el('li', 'attachment-item');
      const label = att.label || att.name || att.text || att.title || 'Attachment';
      const href = safeUrl(att.href || att.url);
      const icon = el('span', 'attachment-icon');
      icon.setAttribute('aria-hidden', 'true');
      icon.textContent = '📎';
      item.appendChild(icon);
      if (href) {
        const link = el('a');
        link.href = href;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        appendHighlighted(link, label, term);
        item.appendChild(link);
      } else {
        const span = el('span', 'attachment-name');
        appendHighlighted(span, label, term);
        item.appendChild(span);
      }
      const extra = [att.type && String(att.type).length <= 24 ? att.type : null, att.size].filter(Boolean).join(' · ');
      if (extra) item.appendChild(el('span', 'attachment-meta', extra));
      list.appendChild(item);
    });
    return list;
  };

  const buildImages = (images) => {
    const wrap = el('div', 'message-images');
    images.forEach((img, index) => {
      if (!img || img.isEmoji) return;
      const src = safeUrl(img.src, { image: true });
      if (!src) return;
      const label = img.alt || img.title || `Image ${index + 1}`;
      const figure = el('div', 'message-image');
      const open = el('button', 'message-image-open');
      open.type = 'button';
      open.setAttribute('aria-label', `Open ${label} full size`);
      const imgEl = el('img');
      imgEl.src = src;
      imgEl.alt = label;
      imgEl.loading = 'lazy';
      open.appendChild(imgEl);
      open.addEventListener('click', () => openImage(src, label));
      const download = el('a', 'message-image-download', 'Download');
      download.href = src;
      download.download = img.alt || 'image';
      download.target = '_blank';
      download.rel = 'noopener noreferrer';
      download.setAttribute('aria-label', `Download ${label}`);
      figure.append(open, download);
      wrap.appendChild(figure);
    });
    return wrap.children.length ? wrap : null;
  };

  const buildReactions = (reactions) => {
    const wrap = el('ul', 'message-reactions');
    wrap.setAttribute('aria-label', 'Reactions');
    reactions.forEach((r) => {
      if (!r) return;
      const count = Number(r.count) || 1;
      const chip = el('li', 'reaction-chip', `${r.emoji || ''} ${count}`);
      const users = Array.isArray(r.users) ? r.users.join(', ') : '';
      chip.title = users;
      chip.setAttribute('aria-label', `${r.emoji || ''} ${pluralize(count, 'reaction')}${users ? `: ${users}` : ''}`);
      wrap.appendChild(chip);
    });
    return wrap;
  };

  const matchesSearch = (msg, term) => {
    const needle = term.toLowerCase();
    if (String(msg.author || '').toLowerCase().includes(needle)) return true;
    if (messageText(msg).toLowerCase().includes(needle)) return true;
    return Array.isArray(msg.attachments) && msg.attachments.some((att) =>
      att && String(att.label || att.name || att.text || '').toLowerCase().includes(needle));
  };

  /**
   * Renders a conversation into container.
   * options: { currentUser, searchTerm, authorIndex: Map(author -> palette index) }
   * Returns { matches } (number of messages matching searchTerm, or total when not searching).
   */
  const renderMessages = (container, messages, options = {}) => {
    const { currentUser = null, authorIndex = new Map() } = options;
    const term = String(options.searchTerm || '').trim();
    container.replaceChildren();

    const all = Array.isArray(messages) ? messages.filter((m) => m && typeof m === 'object') : [];
    if (countMessages(all) === 0 && all.length === 0) {
      container.appendChild(el('p', 'message-list-note', 'No messages in this conversation.'));
      return { matches: 0 };
    }

    const visible = term ? all.filter((msg) => isRealMessage(msg) && matchesSearch(msg, term)) : all;
    if (term && visible.length === 0) {
      container.appendChild(el('p', 'message-list-note', `No messages match “${term}” in this chat.`));
      return { matches: 0 };
    }

    const TIME_THRESHOLD_MS = 3 * 60 * 1000;
    let lastAuthor = null;
    let lastTimestamp = null;
    const fragment = document.createDocumentFragment();

    visible.forEach((msg) => {
      if (msg.type === 'divider') {
        fragment.appendChild(el('div', 'message-divider', messageText(msg)));
        lastAuthor = null;
        lastTimestamp = null;
        return;
      }

      const attachments = Array.isArray(msg.attachments) ? msg.attachments.filter(Boolean) : [];
      const images = Array.isArray(msg.embeddedImages) ? msg.embeddedImages.filter((i) => i && !i.isEmoji) : [];
      const body = messageText(msg);
      if (!body && attachments.length === 0 && images.length === 0) return;

      const date = msg.isoTimestamp ? new Date(msg.isoTimestamp) : (msg.timestamp ? new Date(msg.timestamp) : null);
      const millis = date && !Number.isNaN(date.getTime()) ? date.getTime() : null;
      const isSystem = msg.type === 'system';
      const isSent = !isSystem && !!currentUser && msg.author === currentUser;

      const row = el('div', `message-container ${isSystem ? 'system' : (isSent ? 'sent' : 'received')}`);

      if (!isSystem) {
        const startsGroup = msg.author !== lastAuthor ||
          (lastTimestamp !== null && millis !== null && (millis - lastTimestamp) > TIME_THRESHOLD_MS);
        if (startsGroup || term) {
          const details = el('div', 'message-details');
          const author = el('span', 'message-author');
          appendHighlighted(author, msg.author || 'Unknown', term);
          details.appendChild(author);
          if (msg.timestamp) {
            const time = el('time', 'message-time', msg.timestamp);
            if (msg.isoTimestamp) time.dateTime = msg.isoTimestamp;
            details.appendChild(time);
          }
          if (msg.edited) details.appendChild(el('span', 'message-edited', 'Edited'));
          row.appendChild(details);
        } else {
          row.classList.add('consecutive-message');
        }
      }

      const bubble = el('div', 'message-bubble');
      if (isSent) bubble.classList.add('sent-message');
      const index = authorIndex.get(msg.author);
      if (!isSystem && !isSent && index !== undefined) {
        bubble.dataset.authorIndex = String(index % HUE_COUNT);
      }

      if (msg.replyTo && msg.replyTo.text) {
        const reply = el('blockquote', 'reply-preview');
        reply.appendChild(el('div', 'reply-author', msg.replyTo.author || 'Unknown'));
        const replyText = String(msg.replyTo.text);
        reply.appendChild(el('div', 'reply-text', replyText.length > 100 ? `${replyText.slice(0, 100)}…` : replyText));
        bubble.appendChild(reply);
      }

      if (body) {
        const text = el('div', 'message-text');
        appendHighlighted(text, body, term);
        bubble.appendChild(text);
      }
      if (attachments.length) bubble.appendChild(buildAttachments(attachments, term));
      if (images.length) {
        const imagesEl = buildImages(images);
        if (imagesEl) bubble.appendChild(imagesEl);
      }
      if (Array.isArray(msg.reactions) && msg.reactions.length) bubble.appendChild(buildReactions(msg.reactions));

      row.appendChild(bubble);
      fragment.appendChild(row);

      lastAuthor = isSystem ? null : msg.author;
      if (millis !== null) lastTimestamp = millis;
    });

    container.appendChild(fragment);
    return { matches: term ? visible.length : countMessages(all) };
  };

  const renderParticipants = (list, participants, { currentUser = null, authorIndex = new Map() } = {}) => {
    list.replaceChildren();
    participants.forEach((p) => {
      const chip = el('li', 'participant-chip');
      const index = authorIndex.get(p.name);
      if (index !== undefined) chip.dataset.authorIndex = String(index % HUE_COUNT);
      if (currentUser === p.name) chip.classList.add('is-you');
      const avatar = el('span', 'avatar avatar-xs');
      avatar.setAttribute('aria-hidden', 'true');
      avatar.textContent = initialsFor(p.name);
      if (index !== undefined) avatar.dataset.hue = String(index % HUE_COUNT);
      const name = el('span', 'participant-name', p.name);
      const count = el('span', 'participant-count', String(p.count));
      count.setAttribute('aria-label', pluralize(p.count, 'message'));
      chip.append(avatar, name, count);
      if (currentUser === p.name) chip.appendChild(el('span', 'participant-you', '(you)'));
      list.appendChild(chip);
    });
  };

  const buildAuthorIndex = (participants) => new Map(participants.map((p, i) => [p.name, i]));

  /** Builds a sidebar entry; returns { item, button }. */
  const buildChatListItem = ({ key, title, subtitle, subtitleTitle }) => {
    const item = el('li', 'chat-list-item');
    const button = el('button', 'chat-list-item-main');
    button.type = 'button';
    button.title = title;
    button.dataset.key = key;
    const avatar = el('span', 'avatar avatar-sm');
    avatar.setAttribute('aria-hidden', 'true');
    fillAvatar(avatar, title);
    const content = el('span', 'chat-list-item-content');
    content.appendChild(el('span', 'chat-list-item-title', title));
    if (subtitle) {
      const sub = el('span', 'chat-list-item-subtitle', subtitle);
      if (subtitleTitle) sub.title = subtitleTitle;
      content.appendChild(sub);
    }
    button.append(avatar, content);
    item.appendChild(button);
    return { item, button };
  };

  return {
    HUE_COUNT,
    escapeHtml,
    parseConversationName,
    stripPrefixes,
    hashString,
    hueFor,
    initialsFor,
    pluralize,
    isRealMessage,
    getParticipants,
    countMessages,
    safeUrl,
    fillAvatar,
    renderMessages,
    renderParticipants,
    buildAuthorIndex,
    buildChatListItem
  };
}

const Core = createViewerCore();
const { escapeHtml, parseConversationName, pluralize } = Core;

const STORAGE_KEYS = {
  data: 'savedExtractions',
  meta: 'savedExtractionsMeta',
  legacy: 'teamsChatData',
  self: 'viewerSelf'
};

/** Human-readable info about a stored conversation key. */
// "10/8/2026, 11:28:21 AM" -> "10/8/2026, 11:28 AM" for compact sidebar/header labels.
const shortTime = (text) => String(text || '').replace(/(\d{1,2}:\d{2}):\d{2}/, '$1');

const describeConversation = (key, meta) => {
  const info = meta && meta[key] ? meta[key] : {};
  const parsed = parseConversationName(key);
  const title = info.name || parsed.cleanName;
  let subtitle = '';
  let subtitleTitle = '';
  if (info.source === 'upload') {
    const from = info.fileName ? `from ${info.fileName}` : 'from file';
    const when = info.originalExtractedAt ? ` · extracted ${shortTime(info.originalExtractedAt)}` : '';
    subtitle = `Imported ${from}${when}`;
    subtitleTitle = `Imported ${parsed.extractionTime || ''} ${from}${when}`.replace(/\s+/g, ' ').trim();
  } else if (parsed.extractionTime) {
    subtitle = `Extracted ${shortTime(parsed.extractionTime)}`;
    subtitleTitle = `Extracted ${parsed.extractionTime}`;
  }
  return { title, subtitle, subtitleTitle: subtitleTitle || subtitle, savedAt: info.savedAt || null };
};

const timestampSlug = () => new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);

/**
 * Generates the JSON export: { meta, conversations }. The upload handler accepts this shape and
 * the plain { name: messages } shape.
 */
const generateEnhancedJSONExport = (conversations, metaMap = {}) => {
  const conversationInfo = {};
  Object.keys(conversations).forEach((key) => {
    const { title } = describeConversation(key, metaMap);
    conversationInfo[key] = { ...(metaMap[key] || {}), name: title };
  });
  return JSON.stringify({
    meta: {
      exportedAt: new Date().toISOString(),
      exporter: 'Teams Chat Exporter',
      version: '2',
      conversationCount: Object.keys(conversations).length,
      messageCount: Object.values(conversations).reduce((n, msgs) => n + Core.countMessages(msgs), 0),
      conversationInfo
    },
    conversations
  }, null, 2);
};

/** Generates CSV export of conversations (one row per message). */
const generateCSVExport = (conversations) => {
  const headers = ['conversation', 'id', 'author', 'timestamp', 'text', 'edited', 'reactions_json', 'attachments_json', 'images_json', 'reply_to_json'];
  const rows = [headers.join(',')];

  Object.entries(conversations).forEach(([name, messages]) => {
    messages.forEach((msg) => {
      if (msg.type === 'divider' || msg.type === 'system') return;
      const row = [
        name,
        msg.id || '',
        msg.author || '',
        msg.isoTimestamp || msg.timestamp || '',
        (msg.message || msg.content || '').replace(/\n/g, '\\n').replace(/\r/g, ''),
        msg.edited ? 'true' : 'false',
        JSON.stringify(msg.reactions || []),
        JSON.stringify(msg.attachments || []),
        JSON.stringify(msg.embeddedImages || []),
        JSON.stringify(msg.replyTo || null)
      ].map((field) => `"${String(field).replace(/"/g, '""')}"`);
      rows.push(row.join(','));
    });
  });

  return rows.join('\n');
};

/** Generates TXT export of conversations. */
const generateTXTExport = (conversations) => {
  const lines = [];

  Object.entries(conversations).forEach(([name, messages]) => {
    const { cleanName, extractionTime } = parseConversationName(name);
    lines.push('='.repeat(50));
    lines.push(cleanName);
    if (extractionTime) lines.push(`Extracted ${extractionTime}`);
    lines.push('='.repeat(50));
    lines.push('');

    messages.forEach((msg) => {
      if (msg.type === 'divider') {
        lines.push('');
        lines.push(`--- ${msg.message || msg.content || ''} ---`);
        lines.push('');
        return;
      }
      if (msg.type === 'system') {
        lines.push(`[SYSTEM] ${msg.message || msg.content || ''}`);
        return;
      }

      const ts = msg.isoTimestamp || msg.timestamp || '';
      const author = msg.author || 'Unknown';
      const text = msg.message || msg.content || '';
      const editedMarker = msg.edited ? ' (edited)' : '';

      lines.push(`[${ts}] ${author}${editedMarker}:`);
      lines.push(text);

      if (Array.isArray(msg.attachments) && msg.attachments.length > 0) {
        msg.attachments.forEach((att) => {
          if (!att) return;
          const label = att.label || att.name || att.text || 'Attachment';
          lines.push(`  Attachment: ${label}${att.href ? ` <${att.href}>` : ''}`);
        });
      }

      if (msg.embeddedImages && msg.embeddedImages.length > 0) {
        const nonEmoji = msg.embeddedImages.filter((img) => !img.isEmoji);
        if (nonEmoji.length > 0) {
          lines.push(`  Images: ${pluralize(nonEmoji.length, 'image')}`);
          nonEmoji.forEach((img) => lines.push(`    - ${img.src}`));
        }
      }

      if (msg.reactions && msg.reactions.length > 0) {
        lines.push(`  Reactions: ${msg.reactions.map((r) => `${r.emoji} ${r.count}`).join(' ')}`);
      }

      if (msg.replyTo) {
        const replyText = (msg.replyTo.text || '').substring(0, 50);
        lines.push(`  -> Replying to ${msg.replyTo.author}: "${replyText}..."`);
      }

      lines.push('');
    });

    lines.push('');
    lines.push('');
  });

  return lines.join('\n');
};

/**
 * Script for the exported HTML page. Serialised with Function#toString, so self-contained:
 * it receives the shared core, the conversations, and display info as arguments.
 */
function runExportedPage(Core, conversations, info) {
  const chatList = document.getElementById('chat-list');
  const messageList = document.getElementById('message-list');
  const title = document.getElementById('chat-title');
  const metaLine = document.getElementById('chat-meta');
  const avatar = document.getElementById('chat-avatar');
  const userSelect = document.getElementById('current-user-select');
  const participantsList = document.getElementById('participants-list');
  const STORE_KEY = 'teams-chat-export-viewing-as';

  let viewingAs = {};
  try { viewingAs = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (_e) { viewingAs = {}; }
  const saveViewingAs = () => { try { localStorage.setItem(STORE_KEY, JSON.stringify(viewingAs)); } catch (_e) { /* file:// may block storage */ } };

  let currentKey = null;

  const show = (key) => {
    currentKey = key;
    const entry = info.entries[key];
    const messages = conversations[key] || [];
    const participants = Core.getParticipants(messages);
    const authorIndex = Core.buildAuthorIndex(participants);
    const names = participants.map((p) => p.name);
    let me = viewingAs[entry.title];
    if (!names.includes(me)) me = names.includes(viewingAs.__last) ? viewingAs.__last : null;

    title.textContent = entry.title;
    title.title = entry.title;
    Core.fillAvatar(avatar, entry.title);
    metaLine.textContent = [Core.pluralize(Core.countMessages(messages), 'message'), entry.subtitle].filter(Boolean).join(' · ');

    userSelect.replaceChildren(new Option('Nobody (read only)', ''));
    names.forEach((n) => userSelect.appendChild(new Option(n, n, false, n === me)));
    userSelect.value = me || '';

    Core.renderParticipants(participantsList, participants, { currentUser: me, authorIndex });
    Core.renderMessages(messageList, messages, { currentUser: me, authorIndex });
    messageList.scrollTop = messageList.scrollHeight;

    chatList.querySelectorAll('.chat-list-item-main').forEach((b) => {
      const active = b.dataset.key === key;
      b.closest('.chat-list-item').classList.toggle('active', active);
      if (active) b.setAttribute('aria-current', 'true'); else b.removeAttribute('aria-current');
    });
  };

  userSelect.addEventListener('change', () => {
    const entry = info.entries[currentKey];
    if (userSelect.value) {
      viewingAs[entry.title] = userSelect.value;
      viewingAs.__last = userSelect.value;
    } else {
      delete viewingAs[entry.title];
    }
    saveViewingAs();
    show(currentKey);
  });

  info.order.forEach((key) => {
    const entry = info.entries[key];
    const { item, button } = Core.buildChatListItem({
      key,
      title: entry.title,
      subtitle: entry.listSubtitle,
      subtitleTitle: entry.subtitle
    });
    button.addEventListener('click', () => show(key));
    chatList.appendChild(item);
  });

  chatList.addEventListener('keydown', (event) => {
    const buttons = Array.from(chatList.querySelectorAll('.chat-list-item-main'));
    const index = buttons.indexOf(document.activeElement);
    if (index === -1) return;
    const next = { ArrowDown: index + 1, ArrowUp: index - 1, Home: 0, End: buttons.length - 1 }[event.key];
    if (next === undefined) return;
    event.preventDefault();
    const target = buttons[Math.max(0, Math.min(buttons.length - 1, next))];
    if (target) target.focus();
  });

  if (info.order.length) show(info.order[0]);
}

/** Collects the viewer stylesheets (tokens.css + style.css) as text for the exported page. */
const loadViewerCss = async () => {
  const parts = [];
  for (const file of ['tokens.css', 'style.css']) {
    let text = '';
    try {
      const response = await fetch(file);
      if (response.ok) text = await response.text();
    } catch (_err) {
      text = '';
    }
    if (!text) {
      const sheet = Array.from(document.styleSheets).find((s) => s.href && s.href.split(/[?#]/)[0].endsWith(`/${file}`));
      try {
        text = sheet ? Array.from(sheet.cssRules).map((rule) => rule.cssText).join('\n') : '';
      } catch (_err) {
        text = '';
      }
    }
    parts.push(text);
  }
  return parts.join('\n').replace(/<\/style/gi, '<\\/style');
};

const sha256Base64 = async (text) => {
  if (!(window.crypto && crypto.subtle)) return null;
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  let binary = '';
  new Uint8Array(digest).forEach((b) => { binary += String.fromCharCode(b); });
  return btoa(binary);
};

// JSON that is safe to place inside <script>…</script>.
const scriptSafeJson = (value) => JSON.stringify(value)
  .replace(/</g, '\\u003c')
  .replace(/\u2028/g, '\\u2028')
  .replace(/\u2029/g, '\\u2029');

/**
 * Generates a standalone HTML page that looks like the viewer. All dynamic text is set with
 * textContent at runtime, and a CSP with a script hash blocks anything that is not our script.
 * Returns a Promise<string>.
 */
const generateHTMLExport = async (conversations, metaMap = {}, order = Object.keys(conversations)) => {
  const css = await loadViewerCss();
  const entries = {};
  order.forEach((key) => {
    const { title, subtitle } = describeConversation(key, metaMap);
    const when = shortTime((metaMap[key] && metaMap[key].originalExtractedAt) || parseConversationName(key).extractionTime || '');
    const listSubtitle = [pluralize(Core.countMessages(conversations[key]), 'message'), when].filter(Boolean).join(' · ');
    entries[key] = { title, subtitle, listSubtitle };
  });
  const exportedAt = new Date().toLocaleString();
  const pageTitle = order.length === 1
    ? `${entries[order[0]].title} — Teams chat export`
    : `Teams chat export — ${exportedAt}`;

  const script = `
const conversations = ${scriptSafeJson(conversations)};
const info = ${scriptSafeJson({ order, entries })};
const Core = (${createViewerCore.toString()})();
(${runExportedPage.toString()})(Core, conversations, info);
`;
  const hash = await sha256Base64(script);
  const scriptSrc = hash ? `'sha256-${hash}'` : "'unsafe-inline'";
  const csp = `default-src 'none'; img-src http: https: data: blob:; style-src 'unsafe-inline'; script-src ${scriptSrc}; base-uri 'none'; form-action 'none'`;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(pageTitle)}</title>
<style>
${css}
</style>
</head>
<body class="viewer is-export">
<div id="main-content">
  <nav id="sidebar" aria-labelledby="sidebar-heading">
    <h2 id="sidebar-heading">Conversations <span class="count-badge">${order.length}</span></h2>
    <ul id="chat-list" aria-labelledby="sidebar-heading"></ul>
    <p class="sidebar-hint">Exported ${escapeHtml(exportedAt)} with Teams Chat Exporter</p>
  </nav>
  <main id="chat-area">
    <header id="chat-header">
      <div class="chat-header-row">
        <span id="chat-avatar" class="avatar avatar-md" aria-hidden="true"></span>
        <div class="chat-heading">
          <h1 id="chat-title" class="chat-title"></h1>
          <span id="chat-meta" class="chat-meta"></span>
        </div>
        <div class="you-control">
          <label for="current-user-select">Viewing as:</label>
          <select id="current-user-select"></select>
        </div>
      </div>
      <div class="chat-header-row participants-row">
        <span id="participants-label" class="participants-label">Participants</span>
        <ul id="participants-list" class="participants-list" aria-labelledby="participants-label"></ul>
      </div>
    </header>
    <div id="message-list" role="region" aria-label="Messages" tabindex="0"></div>
  </main>
</div>
<noscript><p class="message-list-note">This export needs JavaScript to display messages.</p></noscript>
<script>${script}</script>
</body>
</html>`;
};

/** Helper to download a file. */
const downloadFile = (content, filename, mimeType) => {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

/**
 * Normalises an uploaded JSON file into [{ name, messages, info }].
 * Accepts { meta, conversations } (our enhanced export), { name: messages }, or a bare messages array.
 */
const normalizeUpload = (parsed, fileName) => {
  let conversations = parsed;
  let infoMap = {};
  if (Array.isArray(parsed)) {
    conversations = { [fileName.replace(/\.json$/i, '') || 'Imported chat']: parsed };
  } else if (parsed && typeof parsed === 'object' && parsed.conversations && typeof parsed.conversations === 'object' && !Array.isArray(parsed.conversations)) {
    conversations = parsed.conversations;
    infoMap = (parsed.meta && parsed.meta.conversationInfo) || {};
  }
  if (!conversations || typeof conversations !== 'object') return [];

  return Object.entries(conversations)
    .filter(([, messages]) => Array.isArray(messages))
    .map(([rawName, messages]) => {
      const info = infoMap[rawName] || {};
      const stripped = Core.stripPrefixes(rawName);
      const name = (typeof info.name === 'string' && info.name.trim()) || stripped.cleanName || 'Untitled chat';
      return {
        name,
        messages: messages.filter((m) => m && typeof m === 'object'),
        originalExtractedAt: stripped.extractionTime || (info.savedAt ? new Date(info.savedAt).toLocaleString() : null)
      };
    });
};

document.addEventListener('DOMContentLoaded', () => {
  const $ = (id) => document.getElementById(id);
  const chatList = $('chat-list');
  const chatCount = $('chat-count');
  const chatHeader = $('chat-header');
  const chatTitle = $('chat-title');
  const chatMeta = $('chat-meta');
  const chatAvatar = $('chat-avatar');
  const messageList = $('message-list');
  const emptyState = $('empty-state');
  const sidebarHint = $('sidebar-hint');
  const fileUpload = $('file-upload');
  const searchInput = $('global-search-input');
  const searchCount = $('search-count');
  const searchClear = $('search-clear');
  const userSelect = $('current-user-select');
  const participantsList = $('participants-list');
  const exportButton = $('export-menu-button');
  const toastRegion = $('toast-region');
  const confirmDialogEl = $('confirm-dialog');

  let allConversations = {};
  let metaMap = {};
  let selfPrefs = { byConversation: {}, last: null };
  let currentKey = null;
  let currentUser = null;
  let usingLegacyData = false;

  // ---------- Toasts ----------
  const showToast = (text, tone = 'info') => {
    const toast = document.createElement('div');
    toast.className = `toast toast-${tone}`;
    toast.textContent = text;
    toastRegion.appendChild(toast);
    setTimeout(() => {
      toast.classList.add('toast-leaving');
      setTimeout(() => toast.remove(), 250);
    }, tone === 'error' ? 7000 : 4000);
  };

  // ---------- Confirm dialog ----------
  const confirmDialog = ({ title, body, confirmLabel }) => new Promise((resolve) => {
    $('confirm-title').textContent = title;
    $('confirm-body').textContent = body;
    $('confirm-ok').textContent = confirmLabel;
    const previousFocus = document.activeElement;
    const onClose = () => {
      confirmDialogEl.removeEventListener('close', onClose);
      if (previousFocus && document.contains(previousFocus)) previousFocus.focus();
      resolve(confirmDialogEl.returnValue === 'confirm');
    };
    confirmDialogEl.returnValue = 'cancel';
    confirmDialogEl.addEventListener('close', onClose);
    confirmDialogEl.showModal();
  });

  // ---------- Menus (menu button pattern) ----------
  const setupMenu = (button, menu) => {
    const items = () => Array.from(menu.querySelectorAll('[role="menuitem"]:not([disabled])'));
    const close = (returnFocus) => {
      if (menu.hidden) return;
      menu.hidden = true;
      button.setAttribute('aria-expanded', 'false');
      if (returnFocus) button.focus();
    };
    const open = (focusLast) => {
      document.querySelectorAll('.menu-list:not([hidden])').forEach((m) => { if (m !== menu) m.dispatchEvent(new CustomEvent('menu-close')); });
      menu.hidden = false;
      button.setAttribute('aria-expanded', 'true');
      const list = items();
      const target = focusLast ? list[list.length - 1] : list[0];
      if (target) target.focus();
    };
    menu.addEventListener('menu-close', () => close(false));
    button.addEventListener('click', () => (menu.hidden ? open(false) : close(false)));
    button.addEventListener('keydown', (event) => {
      if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
        event.preventDefault();
        open(event.key === 'ArrowUp');
      }
    });
    menu.addEventListener('keydown', (event) => {
      const list = items();
      const index = list.indexOf(document.activeElement);
      const move = (i) => { event.preventDefault(); list[(i + list.length) % list.length].focus(); };
      switch (event.key) {
        case 'ArrowDown': move(index + 1); break;
        case 'ArrowUp': move(index - 1); break;
        case 'Home': move(0); break;
        case 'End': move(list.length - 1); break;
        case 'Escape': event.preventDefault(); close(true); break;
        case 'Tab': close(false); break;
        default: break;
      }
    });
    menu.addEventListener('click', (event) => {
      if (event.target.closest('[role="menuitem"]')) close(true);
    });
    document.addEventListener('pointerdown', (event) => {
      if (!menu.hidden && !menu.contains(event.target) && !button.contains(event.target)) close(false);
    });
  };
  setupMenu(exportButton, $('export-menu'));
  setupMenu($('more-menu-button'), $('more-menu'));

  // ---------- Storage helpers ----------
  const storageGet = (keys) => new Promise((resolve) => chrome.storage.local.get(keys, (r) => resolve(r || {})));
  const storageSet = (values) => new Promise((resolve) => chrome.storage.local.set(values, () => resolve()));
  const storageRemove = (keys) => new Promise((resolve) => chrome.storage.local.remove(keys, () => resolve()));

  // Read-modify-write so a stale viewer tab never drops entries another tab just saved.
  const mutateConversations = async (mutator) => {
    const result = await storageGet([STORAGE_KEYS.data, STORAGE_KEYS.meta]);
    // When only the legacy teamsChatData was shown, promote it into savedExtractions first.
    const data = { ...(usingLegacyData ? allConversations : {}), ...(result[STORAGE_KEYS.data] || {}) };
    const meta = { ...(result[STORAGE_KEYS.meta] || {}) };
    const outcome = mutator(data, meta);
    await storageSet({ [STORAGE_KEYS.data]: data, [STORAGE_KEYS.meta]: meta });
    allConversations = data;
    metaMap = meta;
    usingLegacyData = false;
    return outcome;
  };

  const orderedKeys = () => Object.keys(allConversations).reverse(); // newest saved first

  // ---------- Sidebar ----------
  const setRovingItem = (button) => {
    chatList.querySelectorAll('.chat-list-item-main, .chat-list-item-delete-button').forEach((b) => { b.tabIndex = -1; });
    if (!button) return;
    button.tabIndex = 0;
    const del = button.parentElement.querySelector('.chat-list-item-delete-button');
    if (del) del.tabIndex = 0;
  };

  const findItemButton = (key) => Array.from(chatList.querySelectorAll('.chat-list-item-main')).find((b) => b.dataset.key === key) || null;

  const renderChatList = () => {
    chatList.replaceChildren();
    const keys = orderedKeys();
    chatCount.textContent = keys.length ? String(keys.length) : '';
    keys.forEach((key) => {
      const { title, subtitle, subtitleTitle } = describeConversation(key, metaMap);
      const { item, button } = Core.buildChatListItem({ key, title, subtitle, subtitleTitle });
      button.addEventListener('click', () => selectConversation(key));
      button.addEventListener('focus', () => setRovingItem(button));

      const del = document.createElement('button');
      del.type = 'button';
      del.className = 'chat-list-item-delete-button icon-button';
      del.setAttribute('aria-label', `Delete ${title}`);
      del.title = 'Delete conversation';
      del.innerHTML = '<svg fill="currentColor" aria-hidden="true" width="16" height="16" viewBox="0 0 20 20"><path d="M8.5 4h3a1.5 1.5 0 0 0-3 0Zm-1 0a2.5 2.5 0 0 1 5 0h5a.5.5 0 0 1 0 1h-1.05l-1.2 10.34A3 3 0 0 1 12.27 18H7.73a3 3 0 0 1-2.98-2.66L3.55 5H2.5a.5.5 0 0 1 0-1h5ZM5.74 15.23A2 2 0 0 0 7.73 17h4.54a2 2 0 0 0 1.99-1.77L15.44 5H4.56l1.18 10.23ZM8.5 7.5c.28 0 .5.22.5.5v6a.5.5 0 0 1-1 0V8c0-.28.22-.5.5-.5Zm3.5.5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0V8Z"/></svg>';
      del.addEventListener('click', () => requestDelete(key));
      item.appendChild(del);
      chatList.appendChild(item);
    });
    const activeButton = currentKey ? findItemButton(currentKey) : null;
    setRovingItem(activeButton || chatList.querySelector('.chat-list-item-main'));
    markActive();
  };

  const markActive = () => {
    chatList.querySelectorAll('.chat-list-item-main').forEach((b) => {
      const active = b.dataset.key === currentKey;
      b.parentElement.classList.toggle('active', active);
      if (active) b.setAttribute('aria-current', 'true'); else b.removeAttribute('aria-current');
    });
  };

  chatList.addEventListener('keydown', (event) => {
    const buttons = Array.from(chatList.querySelectorAll('.chat-list-item-main'));
    const onMain = event.target.classList.contains('chat-list-item-main');
    const index = buttons.indexOf(event.target.closest('.chat-list-item')?.querySelector('.chat-list-item-main'));
    if (index === -1) return;
    const targetIndex = { ArrowDown: index + 1, ArrowUp: index - 1, Home: 0, End: buttons.length - 1 }[event.key];
    if (targetIndex !== undefined) {
      event.preventDefault();
      const target = buttons[Math.max(0, Math.min(buttons.length - 1, targetIndex))];
      if (target) target.focus();
      return;
    }
    if (onMain && (event.key === 'Delete' || event.key === 'Backspace')) {
      event.preventDefault();
      requestDelete(buttons[index].dataset.key);
    }
  });

  const requestDelete = async (key) => {
    const { title } = describeConversation(key, metaMap);
    const ok = await confirmDialog({
      title: `Delete “${title}”?`,
      body: 'This removes the saved copy from the viewer. It does not affect the chat in Teams.',
      confirmLabel: 'Delete'
    });
    if (!ok) return;
    const keysBefore = orderedKeys();
    const position = keysBefore.indexOf(key);
    await mutateConversations((data, meta) => {
      delete data[key];
      delete meta[key];
    });
    const keysAfter = orderedKeys();
    if (currentKey === key) {
      currentKey = keysAfter[Math.min(position, keysAfter.length - 1)] || null;
    }
    refresh();
    const next = currentKey ? findItemButton(currentKey) : null;
    if (next) next.focus();
    showToast(`Deleted “${title}”.`);
  };

  // ---------- Header & messages ----------
  const selfFor = (title, participants) => {
    const names = participants.map((p) => p.name);
    const stored = selfPrefs.byConversation[title];
    if (stored && names.includes(stored)) return stored;
    if (selfPrefs.last && names.includes(selfPrefs.last)) return selfPrefs.last;
    return null;
  };

  const renderCurrent = ({ keepScroll = false } = {}) => {
    if (!currentKey || !allConversations[currentKey]) return;
    const messages = allConversations[currentKey];
    const { title, subtitle } = describeConversation(currentKey, metaMap);
    const participants = Core.getParticipants(messages);
    const authorIndex = Core.buildAuthorIndex(participants);
    currentUser = selfFor(title, participants);

    chatTitle.textContent = title;
    chatTitle.title = title;
    Core.fillAvatar(chatAvatar, title);
    chatMeta.textContent = [pluralize(Core.countMessages(messages), 'message'), subtitle].filter(Boolean).join(' · ');

    userSelect.replaceChildren(new Option('Not set', ''));
    participants.forEach((p) => userSelect.appendChild(new Option(p.name, p.name, false, p.name === currentUser)));
    userSelect.value = currentUser || '';
    Core.renderParticipants(participantsList, participants, { currentUser, authorIndex });

    const term = searchInput.value.trim();
    const previousScroll = messageList.scrollTop;
    const { matches } = Core.renderMessages(messageList, messages, { currentUser, searchTerm: term, authorIndex });
    searchCount.textContent = term ? (matches ? pluralize(matches, 'match', 'matches') : 'No matches') : '';
    searchClear.hidden = !searchInput.value;
    if (keepScroll) messageList.scrollTop = previousScroll;
    else messageList.scrollTop = term ? 0 : messageList.scrollHeight;
  };

  const selectConversation = (key) => {
    if (!allConversations[key]) return;
    currentKey = key;
    markActive();
    renderCurrent();
  };

  const refresh = () => {
    const keys = orderedKeys();
    const empty = keys.length === 0;
    if (currentKey && !allConversations[currentKey]) currentKey = null;
    if (!currentKey && !empty) currentKey = keys[0];

    renderChatList();
    emptyState.hidden = !empty;
    chatHeader.hidden = empty;
    messageList.hidden = empty;
    sidebarHint.hidden = keys.length < 2;
    exportButton.disabled = empty;
    $('clear-data-button').disabled = empty;
    searchInput.disabled = empty;
    if (empty) {
      messageList.replaceChildren();
      searchCount.textContent = '';
      chatList.appendChild(Object.assign(document.createElement('li'), { className: 'chat-list-empty', textContent: 'Nothing saved yet.' }));
    } else {
      renderCurrent();
    }
  };

  userSelect.addEventListener('change', () => {
    const { title } = describeConversation(currentKey, metaMap);
    const value = userSelect.value || null;
    const byConversation = { ...selfPrefs.byConversation };
    if (value) byConversation[title] = value; else delete byConversation[title];
    selfPrefs = { byConversation, last: value || selfPrefs.last };
    // Clearing the choice for this chat should not immediately re-apply the global fallback.
    if (!value) selfPrefs.last = null;
    storageSet({ [STORAGE_KEYS.self]: selfPrefs });
    renderCurrent({ keepScroll: true });
  });

  // ---------- Search ----------
  searchInput.addEventListener('input', () => renderCurrent());
  searchInput.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && searchInput.value) {
      event.preventDefault();
      searchInput.value = '';
      renderCurrent();
    }
  });
  searchClear.addEventListener('click', () => {
    searchInput.value = '';
    renderCurrent();
    searchInput.focus();
  });

  // ---------- Upload ----------
  const openFilePicker = () => fileUpload.click();
  $('upload-button').addEventListener('click', openFilePicker);
  $('empty-upload-button').addEventListener('click', openFilePicker);

  fileUpload.addEventListener('change', () => {
    const file = fileUpload.files && fileUpload.files[0];
    fileUpload.value = '';
    if (!file) return;
    const reader = new FileReader();
    reader.onerror = () => showToast(`Couldn’t read ${file.name}.`, 'error');
    reader.onload = async () => {
      let entries;
      try {
        entries = normalizeUpload(JSON.parse(reader.result), file.name);
      } catch (error) {
        console.error('Error parsing JSON:', error);
        showToast(`${file.name} isn’t valid JSON.`, 'error');
        return;
      }
      if (!entries.length) {
        showToast(`${file.name} doesn’t contain any chats this viewer understands.`, 'error');
        return;
      }
      const now = new Date();
      const stamp = now.toLocaleString();
      const newKeys = await mutateConversations((data, meta) => {
        const keys = [];
        entries.forEach((entry) => {
          let key = `[${stamp}] ${entry.name}`;
          for (let n = 2; key in data; n += 1) key = `[${stamp} #${n}] ${entry.name}`;
          data[key] = entry.messages;
          meta[key] = {
            name: entry.name,
            source: 'upload',
            fileName: file.name,
            savedAt: now.toISOString(),
            ...(entry.originalExtractedAt ? { originalExtractedAt: entry.originalExtractedAt } : {})
          };
          keys.push(key);
        });
        return keys;
      });
      searchInput.value = '';
      currentKey = newKeys[0];
      refresh();
      showToast(`Imported ${pluralize(newKeys.length, 'chat')} from ${file.name}.`, 'success');
    };
    reader.readAsText(file);
  });

  // ---------- Export ----------
  const exportAll = async (format) => {
    const keys = orderedKeys();
    if (!keys.length) {
      showToast('Nothing to export yet. Extract a chat or upload a JSON export first.', 'error');
      return;
    }
    const ordered = {};
    keys.forEach((k) => { ordered[k] = allConversations[k]; });
    const slug = timestampSlug();
    try {
      if (format === 'json') {
        downloadFile(generateEnhancedJSONExport(ordered, metaMap), `teams-chat-export-${slug}.json`, 'application/json');
      } else if (format === 'csv') {
        downloadFile(generateCSVExport(ordered), `teams-chat-export-${slug}.csv`, 'text/csv');
      } else if (format === 'txt') {
        downloadFile(generateTXTExport(ordered), `teams-chat-export-${slug}.txt`, 'text/plain');
      } else if (format === 'html') {
        downloadFile(await generateHTMLExport(ordered, metaMap, keys), `teams-chat-export-${slug}.html`, 'text/html');
      }
      showToast(`Exported ${pluralize(keys.length, 'chat')} as ${format.toUpperCase()}.`, 'success');
    } catch (error) {
      console.error('Export failed:', error);
      showToast(`Export failed: ${error.message}`, 'error');
    }
  };
  $('export-menu').addEventListener('click', (event) => {
    const item = event.target.closest('[data-export]');
    if (item) exportAll(item.dataset.export);
  });

  // ---------- Clear all ----------
  $('clear-data-button').addEventListener('click', async () => {
    const count = Object.keys(allConversations).length;
    const ok = await confirmDialog({
      title: 'Clear all data?',
      body: `This permanently removes ${pluralize(count, 'saved chat')} and your “You” choices from this browser. Export first if you want to keep a copy.`,
      confirmLabel: 'Clear all data'
    });
    if (!ok) return;
    await storageRemove([STORAGE_KEYS.data, STORAGE_KEYS.meta, STORAGE_KEYS.legacy, STORAGE_KEYS.self]);
    allConversations = {};
    metaMap = {};
    selfPrefs = { byConversation: {}, last: null };
    currentKey = null;
    searchInput.value = '';
    refresh();
    showToast('All saved chats were cleared.');
  });

  // ---------- Loading ----------
  const load = async () => {
    const result = await storageGet([STORAGE_KEYS.data, STORAGE_KEYS.meta, STORAGE_KEYS.legacy, STORAGE_KEYS.self]);
    const saved = result[STORAGE_KEYS.data] || {};
    const legacy = result[STORAGE_KEYS.legacy] || {};
    usingLegacyData = Object.keys(saved).length === 0 && Object.keys(legacy).length > 0;
    allConversations = usingLegacyData ? legacy : saved;
    metaMap = result[STORAGE_KEYS.meta] || {};
    const self = result[STORAGE_KEYS.self] || {};
    selfPrefs = { byConversation: self.byConversation || {}, last: self.last || null };
  };

  // Picks the newest stored entry for a chat name (used by the legacy displayData message).
  const newestKeyFor = (name) => orderedKeys().find((k) => describeConversation(k, metaMap).title === name) || null;

  chrome.runtime.onMessage.addListener((request) => {
    if (request && request.action === 'displayData') {
      // Older background pages pushed the extraction here; it is already in storage, so just reload.
      load().then(() => {
        const wanted = (Array.isArray(request.keys) && request.keys.find((k) => allConversations[k])) ||
          (request.data && newestKeyFor(Object.keys(request.data)[0]));
        if (wanted) currentKey = wanted;
        refresh();
      });
    }
  });

  if (chrome.storage.onChanged) {
    chrome.storage.onChanged.addListener((changes, area) => {
      if (area !== 'local' || !(changes[STORAGE_KEYS.data] || changes[STORAGE_KEYS.meta])) return;
      const nextData = changes[STORAGE_KEYS.data] ? (changes[STORAGE_KEYS.data].newValue || {}) : allConversations;
      if (changes[STORAGE_KEYS.meta]) metaMap = changes[STORAGE_KEYS.meta].newValue || {};
      const before = Object.keys(allConversations).join('\n');
      allConversations = nextData;
      usingLegacyData = false;
      if (Object.keys(allConversations).join('\n') !== before) refresh();
    });
  }

  const params = new URLSearchParams(location.search);
  load().then(() => {
    const requested = params.get('select');
    if (requested && allConversations[requested]) currentKey = requested;
    refresh();
  });
});
