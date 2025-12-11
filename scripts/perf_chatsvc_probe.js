// Run this in the Teams page console (top frame) to find chatsvc URLs and pull one conversation.
(() => {
  // Collect recent resource URLs that look like chat message calls
  const urls = performance.getEntriesByType('resource')
    .map((e) => e.name)
    .filter((u) => u.includes('/conversations/') && u.includes('/messages'));

  const unique = Array.from(new Set(urls));
  console.log('[Probe] Found chatsvc message URLs:', unique);

  // Try to extract the first conversation id (19:...).
  const firstUrl = unique[0];
  if (!firstUrl) {
    console.warn('[Probe] No chatsvc message URLs found. Click a chat and let it load, then rerun.');
    return;
  }

  const match = firstUrl.match(/(19:[^/?#]+)/);
  if (!match) {
    console.warn('[Probe] Could not parse conversation id from:', firstUrl);
    return;
  }

  const convId = match[1];
  console.log('[Probe] Using conversation id:', convId);

  const pageSize = 200;
  const fetchUrl = `https://teams.microsoft.com/api/chatsvc/ca/v1/users/ME/conversations/${convId}/messages?startTime=1&pageSize=${pageSize}&view=msnp24Equivalent`;

  fetch(fetchUrl, { credentials: 'include' })
    .then((resp) => resp.json())
    .then((data) => {
      console.log('[Probe] Sample payload keys:', Object.keys(data || {}));
      console.log('[Probe] Messages length:', Array.isArray(data.messages) ? data.messages.length : 'n/a');
      console.log('[Probe] Payload sample:', data.messages ? data.messages.slice(0, 3) : data);
    })
    .catch((err) => console.error('[Probe] Fetch error', err));
})();
