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

  let allConversations = {};
  let currentConversationName = null;

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

  // Function to render the chat list
  const renderChatList = () => {
    chatList.innerHTML = '';
    for (const name in allConversations) {
      const listItem = document.createElement('li');
      listItem.classList.add('chat-list-item');
      listItem.dataset.conversationName = name;

      const avatarData = generateAvatar(name);
      const latestMessage = allConversations[name][allConversations[name].length - 1];
      const previewText = latestMessage ? `${latestMessage.author}: ${latestMessage.message}` : 'No messages';
      const timestampText = latestMessage ? latestMessage.timestamp : '';

      listItem.innerHTML = `
        <div class="chat-list-item-avatar-wrapper">
          <div class="chat-list-item-avatar" style="background-color: ${avatarData.backgroundColor};">
            <span class="chat-list-item-initials">${avatarData.initials}</span>
          </div>
        </div>
        <div class="chat-list-item-content">
          <div class="chat-list-item-header">
            <span class="chat-list-item-title">${name}</span>
            <span class="chat-list-item-timestamp">${timestampText}</span>
          </div>
          <div class="chat-list-item-preview">${previewText}</div>
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
        currentConversationName = name; // Set current conversation
        renderMessages(name, globalSearchInput.value, avatarData.initials, avatarData.backgroundColor); // Render all messages for the selected conversation
      });
      chatList.appendChild(listItem);
    }
  };

  // Function to render messages for a selected conversation, with optional search term and avatar data
  const renderMessages = (conversationName, searchTerm = '', initials = '', backgroundColor = '') => {
    chatTitle.textContent = conversationName;
    
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
      messages = messages.filter(msg => 
        msg.author.toLowerCase().includes(lowerCaseSearchTerm) ||
        msg.message.toLowerCase().includes(lowerCaseSearchTerm)
      );
      if (messages.length === 0) {
        messageList.innerHTML = '<p style="text-align: center; color: #666;">No messages found matching your search.</p>';
        return;
      }
    }

    let lastAuthor = null;
    let lastTimestamp = null;
    const TIME_THRESHOLD_MS = 3 * 60 * 1000; // 3 minutes in milliseconds

    messages.forEach(msg => {
      const messageContainer = document.createElement('div');
      messageContainer.classList.add('message-container');
      messageContainer.classList.add(msg.type);

      const currentTimestamp = new Date(msg.timestamp).getTime();

      // Only show message details if author is different from previous message
      // OR if the same author sends a message after a significant time gap
      if (msg.author !== lastAuthor || (lastTimestamp && (currentTimestamp - lastTimestamp) > TIME_THRESHOLD_MS)) {
        const messageDetails = document.createElement('div');
        messageDetails.classList.add('message-details');
        messageDetails.textContent = `${msg.author} - ${msg.timestamp}`;
        messageContainer.appendChild(messageDetails);
      } else {
        messageContainer.classList.add('consecutive-message');
      }

      const messageBubble = document.createElement('div');
      messageBubble.classList.add('message-bubble');

      const messageText = document.createElement('div');
      messageText.textContent = msg.message;

      messageBubble.appendChild(messageText);
      messageContainer.appendChild(messageBubble);
      messageList.appendChild(messageContainer);

      lastAuthor = msg.author;
      lastTimestamp = currentTimestamp;
    });
    messageList.scrollTop = messageList.scrollHeight; // Scroll to bottom
  };

  // Handle file upload
  fileUpload.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const uploadedData = JSON.parse(e.target.result);
          allConversations = uploadedData; // Overwrite with uploaded data
          renderChatList();
          chatTitle.textContent = 'Select a conversation';
          messageList.innerHTML = '';
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
      allConversations = request.data;
      renderChatList();
      // Optionally select the first conversation by default
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
    }
  });

  // If the page is opened directly (e.g., for testing), try to load from storage
  // This part might not be strictly necessary if always opened via background script
  // but can be useful for development.
  chrome.storage.local.get(['teamsChatData'], (result) => {
    if (result.teamsChatData) {
      allConversations = result.teamsChatData;
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
    }
  });
});