(() => {
  const LOG_PREFIX = '[Teams API Debug]';
  const CANDIDATE_KEYWORDS = ['messages', 'conversations', 'chatsvc', 'api/csa', 'skype', 'graph'];

  console.log(LOG_PREFIX, 'Starting network capture...');

  // Helper to check URL
  const isInteresting = (url) => {
    const lower = (url || '').toLowerCase();
    return CANDIDATE_KEYWORDS.some(k => lower.includes(k));
  };

  // 1. Hook window.fetch
  const originalFetch = window.fetch;
  window.fetch = async (...args) => {
    const [resource] = args;
    const url = typeof resource === 'string' ? resource : resource?.url || '';
    
    if (isInteresting(url)) {
      console.log(LOG_PREFIX, 'FETCH Request:', url);
    }

    try {
      const response = await originalFetch(...args);
      if (isInteresting(url)) {
        const clone = response.clone();
        clone.text().then(text => {
          try {
            const json = JSON.parse(text);
            console.log(LOG_PREFIX, 'FETCH Response (JSON):', url, json);
            // Check if it looks like a message list
            if (json.value && Array.isArray(json.value) && json.value.length > 0 && json.value[0].messageType) {
                 console.log(LOG_PREFIX, '!!! POTENTIAL MATCH !!!', url);
            }
             if (json.messages && Array.isArray(json.messages)) {
                 console.log(LOG_PREFIX, '!!! POTENTIAL MATCH !!!', url);
            }
          } catch (e) {
            console.log(LOG_PREFIX, 'FETCH Response (Text):', url, text.substring(0, 200) + '...');
          }
        });
      }
      return response;
    } catch (err) {
      console.error(LOG_PREFIX, 'FETCH Error:', url, err);
      throw err;
    }
  };

  // 2. Hook XMLHttpRequest
  const originalOpen = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function(...args) {
    this._debugUrl = args[1];
    return originalOpen.apply(this, args);
  };

  const originalSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.send = function(...args) {
    if (isInteresting(this._debugUrl)) {
      console.log(LOG_PREFIX, 'XHR Request:', this._debugUrl);
      this.addEventListener('load', () => {
        console.log(LOG_PREFIX, 'XHR Response:', this._debugUrl, this.status);
        try {
            const text = this.responseText;
            const json = JSON.parse(text);
             console.log(LOG_PREFIX, 'XHR Response Body:', this._debugUrl, json);
        } catch(e) {
            // ignore
        }
      });
    }
    return originalSend.apply(this, args);
  };

  // 3. Attempt to hook Worker creation (experimental)
  // This might allow us to inject this same script into workers if they are created via new Worker()
  const originalWorker = window.Worker;
  window.Worker = function(scriptURL, options) {
    console.log(LOG_PREFIX, 'New Worker created:', scriptURL);
    return new originalWorker(scriptURL, options);
  };

  console.log(LOG_PREFIX, 'Hooks installed. Please click a chat.');
})();
