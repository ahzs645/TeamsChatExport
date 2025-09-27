/**
 * Generates HTML export of conversations using the existing viewer interface
 */
const generateHTMLExport = (conversations) => {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
  
  // Get the current CSS from the page
  const styleContent = `
body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    margin: 0;
    display: flex;
    flex-direction: column;
    height: 100vh;
    background-color: #f4f4f9;
}

#main-content {
    display: flex;
    flex-grow: 1;
}

#sidebar {
    background-color: #f9f9f9;
    width: 300px;
    padding: 20px;
    border-right: 1px solid #e0e0e0;
    overflow-y: auto;
}

#sidebar h2 {
    margin-top: 0;
    margin-bottom: 15px;
    font-size: 20px;
    color: #333;
}

#chat-list {
    list-style: none;
    padding: 0;
    margin: 0;
}

.chat-list-item {
    display: flex;
    align-items: center;
    padding: 10px;
    margin-bottom: 5px;
    cursor: pointer;
    border-radius: 5px;
    transition: background-color 0.2s;
}

.chat-list-item:hover,
.chat-list-item.active {
    background-color: #e0e0e0;
}

.chat-list-item-avatar-wrapper {
    margin-right: 10px;
}

.chat-list-item-avatar {
    width: 40px;
    height: 40px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: bold;
    color: white;
}

.chat-list-item-initials {
    font-size: 14px;
}

.chat-list-item-content {
    flex-grow: 1;
    min-width: 0;
}

.chat-list-item-header {
    display: flex;
    justify-content: space-between;
    margin-bottom: 2px;
}

.chat-list-item-title {
    font-weight: 600;
    font-size: 14px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

.chat-list-item-timestamp {
    font-size: 12px;
    color: #666;
}

.chat-list-item-preview {
    font-size: 12px;
    color: #666;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

#chat-area {
    flex-grow: 1;
    display: flex;
    flex-direction: column;
    background-color: white;
}

#chat-header {
    background-color: #f5f5f5;
    padding: 15px 20px;
    border-bottom: 1px solid #e0e0e0;
}

.header-content {
    display: flex;
    align-items: center;
    gap: 15px;
}

.avatar-container {
    width: 40px;
    height: 40px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    background-color: #8B5FBF;
}

.chat-avatar-initials {
    color: white;
    font-weight: bold;
    font-size: 16px;
}

#chat-title {
    font-size: 18px;
    font-weight: 600;
    color: #333;
}

#message-list {
    flex-grow: 1;
    padding: 20px;
    overflow-y: auto;
    background-color: #fafafa;
}

.message-container {
    margin-bottom: 15px;
}

.message-container.sent {
    text-align: right;
}

.message-divider {
    text-align: center;
    color: #666;
    font-weight: 600;
    margin: 20px 0 15px;
}

.message-container.system {
    text-align: center;
}

.message-details {
    font-size: 12px;
    color: #666;
    margin-bottom: 5px;
}

.message-bubble {
    display: inline-block;
    padding: 10px 15px;
    border-radius: 15px;
    background-color: #e0e0e0;
    max-width: 70%;
    text-align: left;
}

.message-bubble.sent-message {
    background-color: #8B5FBF !important;
    color: white !important;
    margin-left: auto !important;
    margin-right: 0 !important;
}

.message-container.system .message-bubble {
    background-color: #f0f0f5;
    color: #333;
    margin: 0 auto;
}

.consecutive-message .message-bubble {
    margin-top: 2px;
}
`;

  let html = `<!DOCTYPE html>
<html>
<head>
  <title>Teams Chat Export - ${timestamp}</title>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    ${styleContent}
    
    /* User Selection Modal */
    .modal-overlay {
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0,0,0,0.5);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 1000;
    }
    
    .modal {
      background: white;
      border-radius: 8px;
      padding: 30px;
      max-width: 500px;
      width: 90%;
      max-height: 80vh;
      overflow-y: auto;
    }
    
    .modal h2 {
      margin-top: 0;
      margin-bottom: 20px;
      color: #333;
    }
    
    .user-selection-list {
      max-height: 300px;
      overflow-y: auto;
      margin-bottom: 20px;
    }
    
    .user-option {
      padding: 12px 16px;
      margin: 8px 0;
      border: 2px solid #e0e0e0;
      border-radius: 8px;
      cursor: pointer;
      transition: all 0.2s;
      display: flex;
      align-items: center;
      gap: 12px;
    }
    
    .user-option:hover {
      border-color: #5B5FC5;
      background: #f8f9ff;
    }
    
    .user-option.selected {
      border-color: #5B5FC5;
      background: #5B5FC5;
      color: white;
    }
    
    .user-avatar {
      width: 32px;
      height: 32px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: bold;
      color: white;
      font-size: 12px;
      flex-shrink: 0;
    }
    
    .user-info {
      flex-grow: 1;
    }
    
    .user-name {
      font-weight: 600;
      margin-bottom: 2px;
    }
    
    .user-message-count {
      font-size: 12px;
      opacity: 0.7;
    }
    
    .modal-buttons {
      display: flex;
      gap: 12px;
      justify-content: flex-end;
    }
    
    .btn {
      padding: 8px 16px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-weight: 600;
    }
    
    .btn-primary {
      background: #5B5FC5;
      color: white;
    }
    
    .btn-secondary {
      background: #e0e0e0;
      color: #333;
    }
    
    .btn:hover {
      opacity: 0.9;
    }
    
    .header-actions {
      display: flex;
      gap: 10px;
      align-items: center;
    }
    
    .change-user-btn {
      padding: 6px 12px;
      background: #5B5FC5;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 12px;
    }
    
    .change-user-btn:hover {
      background: #4a4d9e;
    }
    
    .current-user-indicator {
      font-size: 12px;
      color: #666;
      margin-left: 10px;
    }
  </style>
</head>
<body>
  <div id="main-content">
    <div id="sidebar">
      <h2>Conversations</h2>
      <ul id="chat-list"></ul>
    </div>
    <div id="chat-area">
      <div id="chat-header">
        <div class="header-content">
          <div class="avatar-container">
            <span class="chat-avatar-initials"></span>
          </div>
          <div class="title-container">
            <span id="chat-title">Select a conversation</span>
          </div>
          <div class="header-actions">
            <button class="change-user-btn" onclick="showUserSelection()">Change User</button>
            <span class="current-user-indicator" id="current-user-display">No user selected</span>
          </div>
        </div>
      </div>
      <div id="message-list"></div>
    </div>
  </div>
  
  <!-- User Selection Modal -->
  <div id="user-selection-modal" class="modal-overlay" style="display: none;">
    <div class="modal">
      <h2>Who are you in this conversation?</h2>
      <p style="color: #666; margin-bottom: 20px;">Select your name to properly align sent/received messages:</p>
      <div class="user-selection-list" id="user-list"></div>
      <div class="modal-buttons">
        <button class="btn btn-secondary" onclick="closeUserSelection()">Cancel</button>
        <button class="btn btn-primary" onclick="confirmUserSelection()" id="confirm-btn" disabled>Confirm</button>
      </div>
    </div>
  </div>
  
  <script>
    // Embed the conversation data
    const allConversations = ${JSON.stringify(conversations)};
    let currentConversationName = null;
    let currentUser = null;
    let selectedUserOption = null;
    
    // Function to generate initials and random background color
    const generateAvatar = (name) => {
      const words = name.split(' ').filter(word => word.length > 0);
      let initials = '';
      if (words.length >= 2) {
        initials = words[0][0] + words[1][0];
      } else if (words.length === 1) {
        initials = words[0][0];
      }
      
      const pastelColors = [
        '#FFB3BA', '#FFDFBA', '#FFFFBA', '#BAFFC9', '#BAE1FF', 
        '#E0BBE4', '#957DAD', '#D291BC', '#FFC72C', '#DA2C38'
      ];
      const randomColor = pastelColors[Math.floor(Math.random() * pastelColors.length)];
      
      return { initials: initials.toUpperCase(), backgroundColor: randomColor };
    };
    
    // Function to render messages
    const renderMessages = (conversationName) => {
      currentConversationName = conversationName;
      
      const chatTitle = document.getElementById('chat-title');
      const messageList = document.getElementById('message-list');
      const avatarContainer = document.querySelector('.avatar-container');
      const avatarInitials = document.querySelector('.chat-avatar-initials');
      
      const avatarData = generateAvatar(conversationName);
      chatTitle.textContent = conversationName;
      avatarInitials.textContent = avatarData.initials;
      avatarContainer.style.backgroundColor = avatarData.backgroundColor;
      
      messageList.innerHTML = '';
      const messages = allConversations[conversationName];
      
      if (!messages || messages.length === 0) {
        messageList.innerHTML = '<p style="text-align: center; color: #666;">No messages in this conversation.</p>';
        return;
      }
      
      let lastAuthor = null;
      let lastTimestamp = null;
      const TIME_THRESHOLD_MS = 3 * 60 * 1000;
      
      messages.forEach(msg => {
        if (msg.type === 'divider') {
          const dividerEl = document.createElement('div');
          dividerEl.classList.add('message-divider');
          dividerEl.textContent = msg.message;
          messageList.appendChild(dividerEl);
          lastAuthor = null;
          lastTimestamp = null;
          return;
        }

        const messageContainer = document.createElement('div');
        messageContainer.classList.add('message-container');

        const timestampDate = msg.isoTimestamp ? new Date(msg.isoTimestamp) : (msg.timestamp ? new Date(msg.timestamp) : null);
        const timestampMillis = timestampDate && !Number.isNaN(timestampDate.getTime())
          ? timestampDate.getTime()
          : null;

        const authorLabel = msg.author || 'Unknown';
        const messageType = msg.type || null;
        const isSystemMessage = messageType === 'system';

        if (!isSystemMessage) {
          if (
            msg.author !== lastAuthor ||
            (lastTimestamp !== null && timestampMillis !== null && (timestampMillis - lastTimestamp) > TIME_THRESHOLD_MS)
          ) {
            const messageDetails = document.createElement('div');
            messageDetails.classList.add('message-details');
            messageDetails.textContent = msg.timestamp ? authorLabel + ' - ' + msg.timestamp : authorLabel;
            messageContainer.appendChild(messageDetails);
          } else {
            messageContainer.classList.add('consecutive-message');
          }
        }

        const messageBubble = document.createElement('div');
        messageBubble.classList.add('message-bubble');

        const isFromCurrentUser = currentUser && msg.author === currentUser;
        const bubbleType = messageType || (isFromCurrentUser ? 'sent' : 'received');
        const isSent = bubbleType === 'sent';

        messageContainer.classList.add(bubbleType);

        if (isSystemMessage) {
          messageContainer.style.textAlign = 'center';
        } else if (isSent) {
          messageBubble.classList.add('sent-message');
          messageContainer.style.textAlign = 'right';
        }

        const messageText = document.createElement('div');
        const messageBody = (msg.message || msg.content || '').trim();
        if (!messageBody && (!Array.isArray(msg.attachments) || msg.attachments.length === 0)) {
          return;
        }
        messageText.textContent = messageBody;

        messageBubble.appendChild(messageText);
        messageContainer.appendChild(messageBubble);
        messageList.appendChild(messageContainer);

        if (isSystemMessage) {
          lastAuthor = null;
          if (timestampMillis !== null) {
            lastTimestamp = timestampMillis;
          }
        } else {
          lastAuthor = msg.author;
          if (timestampMillis !== null) {
            lastTimestamp = timestampMillis;
          }
        }
      });
      
      messageList.scrollTop = messageList.scrollHeight;
    };
    
    // Function to render chat list
    const renderChatList = () => {
      const chatList = document.getElementById('chat-list');
      chatList.innerHTML = '';
      
      for (const name in allConversations) {
        const listItem = document.createElement('li');
        listItem.classList.add('chat-list-item');
        
        const avatarData = generateAvatar(name);
        const messages = allConversations[name] || [];
        const latestMessage = [...messages].reverse().find(msg => msg.type !== 'divider' && msg.type !== 'system');
        const previewAuthor = latestMessage?.author || '';
        const previewBody = latestMessage?.message || latestMessage?.content || '';
        const previewText = latestMessage
          ? (previewAuthor ? previewAuthor + ': ' : '') + previewBody
          : 'No messages';
        const timestampText = latestMessage && latestMessage.timestamp ? latestMessage.timestamp : '';
        
        listItem.innerHTML = \`
          <div class="chat-list-item-avatar-wrapper">
            <div class="chat-list-item-avatar" style="background-color: \${avatarData.backgroundColor};">
              <span class="chat-list-item-initials">\${avatarData.initials}</span>
            </div>
          </div>
          <div class="chat-list-item-content">
            <div class="chat-list-item-header">
              <span class="chat-list-item-title">\${name}</span>
              <span class="chat-list-item-timestamp">\${timestampText}</span>
            </div>
            <div class="chat-list-item-preview">\${previewText}</div>
          </div>
        \`;
        
        listItem.addEventListener('click', () => {
          document.querySelectorAll('.chat-list-item').forEach(item => {
            item.classList.remove('active');
          });
          listItem.classList.add('active');
          renderMessages(name);
        });
        
        chatList.appendChild(listItem);
      }
      
      // Auto-select first conversation
      if (Object.keys(allConversations).length > 0) {
        const firstName = Object.keys(allConversations)[0];
        renderMessages(firstName);
        chatList.firstChild.classList.add('active');
      }
    };
    
    // User Selection Functions
    const getAllUsers = () => {
      const users = new Set();
      Object.values(allConversations).forEach(messages => {
        messages.forEach(msg => {
          if (msg.author && msg.author.trim()) {
            users.add(msg.author);
          }
        });
      });
      return Array.from(users).sort();
    };
    
    const getUserMessageCount = (userName) => {
      let count = 0;
      Object.values(allConversations).forEach(messages => {
        messages.forEach(msg => {
          if (msg.author === userName) count++;
        });
      });
      return count;
    };
    
    const showUserSelection = () => {
      const modal = document.getElementById('user-selection-modal');
      const userList = document.getElementById('user-list');
      const users = getAllUsers();
      
      userList.innerHTML = '';
      
      users.forEach(userName => {
        const messageCount = getUserMessageCount(userName);
        const avatarData = generateAvatar(userName);
        
        const userOption = document.createElement('div');
        userOption.className = 'user-option';
        userOption.dataset.userName = userName;
        
        if (currentUser === userName) {
          userOption.classList.add('selected');
          selectedUserOption = userOption;
        }
        
        userOption.innerHTML = \`
          <div class="user-avatar" style="background-color: \${avatarData.backgroundColor};">
            \${avatarData.initials}
          </div>
          <div class="user-info">
            <div class="user-name">\${userName}</div>
            <div class="user-message-count">\${messageCount} messages</div>
          </div>
        \`;
        
        userOption.addEventListener('click', () => {
          if (selectedUserOption) {
            selectedUserOption.classList.remove('selected');
          }
          userOption.classList.add('selected');
          selectedUserOption = userOption;
          document.getElementById('confirm-btn').disabled = false;
        });
        
        userList.appendChild(userOption);
      });
      
      modal.style.display = 'flex';
      document.getElementById('confirm-btn').disabled = !selectedUserOption;
    };
    
    const closeUserSelection = () => {
      document.getElementById('user-selection-modal').style.display = 'none';
      selectedUserOption = null;
    };
    
    const confirmUserSelection = () => {
      if (selectedUserOption) {
        currentUser = selectedUserOption.dataset.userName;
        document.getElementById('current-user-display').textContent = 'You: ' + currentUser;
        
        // Re-render current conversation to update message alignment
        if (currentConversationName) {
          renderMessages(currentConversationName);
        }
      }
      closeUserSelection();
    };
    
    // Initialize
    renderChatList();
    
    // Show user selection on first load
    setTimeout(() => {
      if (!currentUser) {
        showUserSelection();
      }
    }, 1000);
  </script>
</body>
</html>`;
  
  return html;
};

/**
 * Escapes HTML special characters to prevent XSS
 */
const escapeHtml = (text) => {
  if (!text) return '';
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, m => map[m]);
};

document.addEventListener('DOMContentLoaded', () => {
  const chatList = document.getElementById('chat-list');
  const chatTitle = document.getElementById('chat-title');
  const messageList = document.getElementById('message-list');
  const fileUpload = document.getElementById('file-upload');
  const globalSearchInput = document.getElementById('global-search-input');
  const conversationTab = document.querySelector('.fui-TabList .fui-Tab:nth-child(1)');
  const sharedTab = document.querySelector('.fui-TabList .fui-Tab:nth-child(2)');
  const chatAvatarContainer = document.querySelector('#chat-header .avatar-container');
  const chatAvatarInitials = document.querySelector('#chat-header .chat-avatar-initials');
  const downloadJsonButton = document.getElementById('download-json-button');
  const downloadHtmlButton = document.getElementById('download-html-button');
  const clearDataButton = document.getElementById('clear-data-button');
  const currentUserDisplay = document.getElementById('current-user-display');

  let allConversations = {};
  let currentConversationName = null;
  let currentUser = null;
  let selectedUserOption = null;

  // Function to generate initials and a random pastel background color
  const generateAvatar = (name) => {
    const words = name.split(' ').filter(word => word.length > 0);
    let initials = '';
    if (words.length >= 2) {
      initials = words[0][0] + words[1][0];
    } else if (words.length === 1) {
      initials = words[0][0];
    }

    const pastelColors = [
      '#FFB3BA', '#FFDFBA', '#FFFFBA', '#BAFFC9', '#BAE1FF', 
      '#E0BBE4', '#957DAD', '#D291BC', '#FFC72C', '#DA2C38'
    ];
    const randomColor = pastelColors[Math.floor(Math.random() * pastelColors.length)];

    return { initials: initials.toUpperCase(), backgroundColor: randomColor };
  };

  // Function to extract clean name and extraction info from timestamped names
  const parseConversationName = (fullName) => {
    // Check if name has timestamp prefix like "[8/29/2025, 11:28:21 AM] Name"
    const timestampMatch = fullName.match(/^\[([^\]]+)\]\s*(.+)$/);
    if (timestampMatch) {
      const extractionTime = timestampMatch[1];
      const cleanName = timestampMatch[2];
      return { cleanName, extractionTime, isTimestamped: true };
    }
    return { cleanName: fullName, extractionTime: null, isTimestamped: false };
  };

  // Function to render the chat list
  const renderChatList = () => {
    chatList.innerHTML = '';
    for (const name in allConversations) {
      const listItem = document.createElement('li');
      listItem.classList.add('chat-list-item');
      listItem.dataset.conversationName = name;

      const { cleanName, extractionTime } = parseConversationName(name);
      const avatarData = generateAvatar(cleanName);

      listItem.innerHTML = `
        <div class="chat-list-item-avatar-wrapper">
          <div class="chat-list-item-avatar" style="background-color: ${avatarData.backgroundColor};">
            <span class="chat-list-item-initials">${avatarData.initials}</span>
          </div>
        </div>
        <div class="chat-list-item-content">
          <div class="chat-list-item-header">
            <span class="chat-list-item-title">${cleanName}</span>
          </div>
          <div class="chat-list-item-extraction-date">${extractionTime || ''}</div>
        </div>
        <div class="chat-list-item-actions">
          <button type="button" class="fui-Button chat-list-item-more-button">
            <span class="fui-Button__icon">
              <svg class="fui-Icon-regular" fill="currentColor" aria-hidden="true" width="1em" height="1em" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg"><path d="M6.25 10a1.25 1.25 0 1 1-2.5 0 1.25 1.25 0 0 1 2.5 0Zm5 0a1.25 1.25 0 1 1-2.5 0 1.25 1.25 0 0 1 2.5 0ZM15 11.25a1.25 1.25 0 1 0 0-2.5 1.25 1.25 0 0 0 0 2.5Z" fill="currentColor"></path></svg>
            </span>
          </button>
        </div>
      `;

      listItem.addEventListener('click', () => {
        // Remove active class from previous item
        const currentActive = document.querySelector('.chat-list-item.active');
        if (currentActive) {
          currentActive.classList.remove('active');
        }
        // Add active class to clicked item
        listItem.classList.add('active');
        currentConversationName = name; // Set current conversation (use full name for data lookup)
        renderMessages(name, globalSearchInput.value, avatarData.initials, avatarData.backgroundColor); // Render all messages for the selected conversation
      });
      chatList.appendChild(listItem);
    }
  };

  // Function to render messages for a selected conversation, with optional search term and avatar data
  const renderMessages = (conversationName, searchTerm = '', initials = '', backgroundColor = '') => {
    const { cleanName } = parseConversationName(conversationName);
    chatTitle.textContent = cleanName;
    
    // Update avatar
    chatAvatarInitials.textContent = initials;
    chatAvatarContainer.style.backgroundColor = backgroundColor;

    messageList.innerHTML = '';
    let messages = allConversations[conversationName];

    if (!messages || messages.length === 0) {
      messageList.innerHTML = '<p style="text-align: center; color: #666;">No messages in this conversation.</p>';
      return;
    }

    // Filter messages if a search term is provided
    if (searchTerm) {
      const lowerCaseSearchTerm = searchTerm.toLowerCase();
      messages = messages.filter((msg) => {
        const authorText = (msg.author || '').toLowerCase();
        const bodyText = (msg.message || msg.content || '').toLowerCase();
        return authorText.includes(lowerCaseSearchTerm) || bodyText.includes(lowerCaseSearchTerm);
      });
      if (messages.length === 0) {
        messageList.innerHTML = '<p style="text-align: center; color: #666;">No messages found matching your search.</p>';
        return;
      }
    }

    let lastAuthor = null;
    let lastTimestamp = null;
    const TIME_THRESHOLD_MS = 3 * 60 * 1000; // 3 minutes in milliseconds

    messages.forEach((msg) => {
      if (msg.type === 'divider') {
        const dividerEl = document.createElement('div');
        dividerEl.classList.add('message-divider');
        dividerEl.textContent = msg.message;
        messageList.appendChild(dividerEl);
        lastAuthor = null;
        lastTimestamp = null;
        return;
      }

      const messageContainer = document.createElement('div');
      messageContainer.classList.add('message-container');

      const timestampDate = msg.isoTimestamp ? new Date(msg.isoTimestamp) : (msg.timestamp ? new Date(msg.timestamp) : null);
      const timestampMillis = timestampDate && !Number.isNaN(timestampDate.getTime())
        ? timestampDate.getTime()
        : null;

      const authorLabel = msg.author || 'Unknown';
      const messageType = msg.type || null;
      const isSystemMessage = messageType === 'system';

      if (!isSystemMessage) {
        if (
          msg.author !== lastAuthor ||
          (lastTimestamp !== null && timestampMillis !== null && (timestampMillis - lastTimestamp) > TIME_THRESHOLD_MS)
        ) {
          const messageDetails = document.createElement('div');
          messageDetails.classList.add('message-details');
          messageDetails.textContent = msg.timestamp ? authorLabel + ' - ' + msg.timestamp : authorLabel;
          messageContainer.appendChild(messageDetails);
        } else {
          messageContainer.classList.add('consecutive-message');
        }
      }

      const messageBubble = document.createElement('div');
      messageBubble.classList.add('message-bubble');

      const isFromCurrentUser = currentUser && msg.author === currentUser;
      const bubbleType = messageType || (isFromCurrentUser ? 'sent' : 'received');
      const isSent = bubbleType === 'sent';

      messageContainer.classList.add(bubbleType);

      if (isSystemMessage) {
        messageContainer.style.textAlign = 'center';
      } else if (isSent) {
        messageBubble.classList.add('sent-message');
        messageContainer.style.textAlign = 'right';
      }

      const messageText = document.createElement('div');
      const messageBody = (msg.message || msg.content || '').trim();
      if (!messageBody && (!Array.isArray(msg.attachments) || msg.attachments.length === 0)) {
        return;
      }
      messageText.textContent = messageBody;

      messageBubble.appendChild(messageText);
      messageContainer.appendChild(messageBubble);
      messageList.appendChild(messageContainer);

      if (isSystemMessage) {
        lastAuthor = null;
        if (timestampMillis !== null) {
          lastTimestamp = timestampMillis;
        }
      } else {
        lastAuthor = msg.author;
        if (timestampMillis !== null) {
          lastTimestamp = timestampMillis;
        }
      }
    });
    messageList.scrollTop = messageList.scrollHeight; // Scroll to bottom
  };

  const getAllUsers = () => {
    const users = new Set();
    Object.values(allConversations).forEach((messages) => {
      messages.forEach((msg) => {
        if (msg.author && msg.author.trim()) {
          users.add(msg.author);
        }
      });
    });
    return Array.from(users).sort();
  };

  const getUserMessageCount = (userName) => {
    let count = 0;
    Object.values(allConversations).forEach((messages) => {
      messages.forEach((msg) => {
        if (msg.author === userName) {
          count += 1;
        }
      });
    });
    return count;
  };

  const showUserSelection = () => {
    const modal = document.getElementById('user-selection-modal');
    const userList = document.getElementById('user-list');
    const confirmButton = document.getElementById('confirm-btn');

    if (!modal || !userList || !confirmButton) {
      return;
    }

    const users = getAllUsers();

    userList.innerHTML = '';
    selectedUserOption = null;

    users.forEach((userName) => {
      const messageCount = getUserMessageCount(userName);
      const avatarData = generateAvatar(userName);

      const userOption = document.createElement('div');
      userOption.className = 'user-option';
      userOption.dataset.userName = userName;

      if (currentUser === userName) {
        userOption.classList.add('selected');
        selectedUserOption = userOption;
      }

      userOption.innerHTML = `
        <div class="user-avatar" style="background-color: ${avatarData.backgroundColor};">
          ${avatarData.initials}
        </div>
        <div class="user-info">
          <div class="user-name">${userName}</div>
          <div class="user-message-count">${messageCount} messages</div>
        </div>
      `;

      userOption.addEventListener('click', () => {
        if (selectedUserOption) {
          selectedUserOption.classList.remove('selected');
        }
        userOption.classList.add('selected');
        selectedUserOption = userOption;
        confirmButton.disabled = false;
      });

      userList.appendChild(userOption);
    });

    modal.style.display = 'flex';
    confirmButton.disabled = !selectedUserOption;
  };

  const closeUserSelection = () => {
    const modal = document.getElementById('user-selection-modal');
    if (modal) {
      modal.style.display = 'none';
    }
    selectedUserOption = null;
  };

  const confirmUserSelection = () => {
    if (selectedUserOption) {
      currentUser = selectedUserOption.dataset.userName;
      if (currentUserDisplay) {
        currentUserDisplay.textContent = 'You: ' + currentUser;
      }
      if (currentConversationName) {
        renderMessages(currentConversationName);
      }
    }
    closeUserSelection();
  };

  window.showUserSelection = showUserSelection;
  window.closeUserSelection = closeUserSelection;
  window.confirmUserSelection = confirmUserSelection;

  // Handle file upload
  fileUpload.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const uploadedData = JSON.parse(e.target.result);
          
          // Add uploaded data with timestamp prefix instead of replacing
          const timestamp = new Date().toLocaleString();
          const fileName = file.name.replace('.json', '');
          
          Object.keys(uploadedData).forEach(conversationName => {
            const newName = `[${fileName}] ${conversationName}`;
            allConversations[newName] = uploadedData[conversationName];
          });
          
          renderChatList();
          
          // Automatically select the first uploaded conversation
          const firstUploadedName = `[${fileName}] ${Object.keys(uploadedData)[0]}`;
          if (Object.keys(uploadedData).length > 0) {
            const firstListItem = document.querySelector(`[data-conversation-name="${firstUploadedName}"]`);
            if (firstListItem) {
              // Remove active from previous items
              document.querySelectorAll('.chat-list-item').forEach(item => {
                item.classList.remove('active');
              });
              firstListItem.classList.add('active');
              
              currentConversationName = firstUploadedName;
              const avatarData = generateAvatar(firstUploadedName);
              renderMessages(firstUploadedName, '', avatarData.initials, avatarData.backgroundColor);
            }
          }
          
          globalSearchInput.value = ''; // Clear search on new upload
        } catch (error) {
          alert('Invalid JSON file. Please upload a valid JSON.');
          console.error('Error parsing JSON:', error);
        }
      };
      reader.readAsText(file);
    }
  });

  // Handle JSON download
  downloadJsonButton.addEventListener('click', () => {
    if (Object.keys(allConversations).length === 0) {
      alert('No data to download. Please upload a JSON file first.');
      return;
    }
    const dataStr = JSON.stringify(allConversations, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'teams_chat_data.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  // Handle HTML download
  downloadHtmlButton.addEventListener('click', () => {
    if (Object.keys(allConversations).length === 0) {
      alert('No data to export. Please upload a JSON file or extract conversations first.');
      return;
    }
    const htmlContent = generateHTMLExport(allConversations);
    const blob = new Blob([htmlContent], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    a.download = `teams-chat-export-${timestamp}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  // Global search functionality
  globalSearchInput.addEventListener('input', () => {
    if (currentConversationName) {
      renderMessages(currentConversationName, globalSearchInput.value);
    }
  });

  // Tab functionality
  if (conversationTab) {
    conversationTab.addEventListener('click', () => {
      conversationTab.classList.add('active');
      sharedTab.classList.remove('active');
      // In a real scenario, you'd load conversation-specific content here
    });
  }

  if (sharedTab) {
    sharedTab.addEventListener('click', () => {
      sharedTab.classList.add('active');
      conversationTab.classList.remove('active');
      // In a real scenario, you'd load shared content here
    });
  }

  // Initial load: check if data is passed from background script
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "displayData" && request.data) {
      // Merge new conversations with existing ones instead of replacing
      const timestamp = new Date().toLocaleString();
      Object.keys(request.data).forEach(conversationName => {
        const newName = `[${timestamp}] ${conversationName}`;
        allConversations[newName] = request.data[conversationName];
      });
      
      renderChatList();
      // Optionally select the first NEW conversation by default
      const newConversationNames = Object.keys(request.data).map(name => `[${timestamp}] ${name}`);
      if (newConversationNames.length > 0) {
        const firstNewName = newConversationNames[0];
        const firstListItem = document.querySelector(`[data-conversation-name="${firstNewName}"]`);
        if (firstListItem) {
          // Remove active from previous items
          document.querySelectorAll('.chat-list-item').forEach(item => {
            item.classList.remove('active');
          });
          firstListItem.classList.add('active');
        }
        currentConversationName = firstNewName; // Set current conversation
        const firstAvatarData = generateAvatar(firstNewName);
        renderMessages(firstNewName, '', firstAvatarData.initials, firstAvatarData.backgroundColor);
      }
    }
  });

  // Handle clear data button
  clearDataButton.addEventListener('click', () => {
    if (confirm('Are you sure you want to clear all conversation data? This cannot be undone.')) {
      allConversations = {};
      currentConversationName = null;
      currentUser = null;
      if (currentUserDisplay) {
        currentUserDisplay.textContent = 'No user selected';
      }
      chrome.storage.local.remove(['teamsChatData', 'savedExtractions']);
      renderChatList();
      document.getElementById('chat-title').textContent = 'Select a conversation';
      document.getElementById('message-list').innerHTML = '';
    }
  });

  // Load both current extraction and all saved extractions
  chrome.storage.local.get(['teamsChatData', 'savedExtractions'], (result) => {
    const savedExtractions = result.savedExtractions || {};
    const teamsChatData = result.teamsChatData || {};

    const hasSavedExtracts = Object.keys(savedExtractions).length > 0;
    const hasCurrentExtraction = Object.keys(teamsChatData).length > 0;

    if (hasSavedExtracts) {
      allConversations = savedExtractions;
    } else if (hasCurrentExtraction) {
      allConversations = teamsChatData;
    }
    
    renderChatList();
    if (Object.keys(allConversations).length > 0) {
      const firstConversationName = Object.keys(allConversations)[0];
      const firstListItem = document.querySelector(`[data-conversation-name="${firstConversationName}"]`);
      if (firstListItem) {
        firstListItem.classList.add('active');
      }
      currentConversationName = firstConversationName; // Set current conversation
      const firstAvatarData = generateAvatar(firstConversationName);
      renderMessages(firstConversationName, '', firstAvatarData.initials, firstAvatarData.backgroundColor);
    }
    if (!currentUser) {
      const users = getAllUsers();
      if (users.length === 1) {
        currentUser = users[0];
        document.getElementById('current-user-display').textContent = 'You: ' + currentUser;
      } else if (users.length > 1) {
        setTimeout(() => {
          if (!currentUser) {
            showUserSelection();
          }
        }, 300);
      }
    }
  });
});
