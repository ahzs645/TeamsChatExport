// Run this in the Teams page console while a chat is open
(() => {
  const results = {};
  const activeItem = document.querySelector('[data-tid="chat-list-item"][aria-selected="true"], [aria-selected="true"][data-qa-chat-id]');
  results.activeItemDataset = activeItem ? { ...activeItem.dataset } : null;
  results.activeItemAttrs = activeItem ? Array.from(activeItem.attributes).reduce((acc, attr) => ({ ...acc, [attr.name]: attr.value }), {}) : null;
  results.linkHref = (document.querySelector('a[href*="19:"]') || {}).href || null;
  results.locationHref = window.location.href;
  results.hiddenCacheSample = (() => {
    const div = document.getElementById('teams-extractor-message-api-cache');
    if (!div || !div.textContent) return null;
    try {
      const parsed = JSON.parse(div.textContent);
      return Array.isArray(parsed) ? parsed.slice(-2) : null;
    } catch (_err) {
      return 'parse-error';
    }
  })();
  results.memoryCacheSize = Array.isArray(window.__teamsExtractorMessageApiCache) ? window.__teamsExtractorMessageApiCache.length : 0;
  console.log('[Extractor debug]', results);
  return results;
})();
