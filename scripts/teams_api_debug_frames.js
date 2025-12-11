// Run this in the top-level Teams page console (it will probe all same-origin frames)
(function () {
  const debugFn = () => {
    const results = {};
    const activeItem = document.querySelector('[data-tid="chat-list-item"][aria-selected="true"], [aria-selected="true"][data-qa-chat-id], [role="treeitem"][aria-selected="true"]');
    results.activeItemDataset = activeItem ? { ...activeItem.dataset } : null;
    results.activeItemAttrs = activeItem ? Array.from(activeItem.attributes).reduce((acc, attr) => ({ ...acc, [attr.name]: attr.value }), {}) : null;
    results.activeItemText = activeItem ? (activeItem.textContent || '').slice(0, 200) : null;

    const anchorWith19 = Array.from(document.querySelectorAll('a')).find((a) => (a.href || '').includes('19:'));
    results.linkHref = anchorWith19?.href || null;

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

    console.log('[Extractor frame debug]', results);
    return results;
  };

  Array.from(window.frames).forEach((frame, idx) => {
    try {
      console.log(`-- Probing frame ${idx} href:`, frame.location?.href);
      frame.eval(`(${debugFn.toString()})()`);
    } catch (err) {
      console.warn(`Frame ${idx} not accessible`, err?.message || err);
    }
  });

  // Also run in the current frame
  try {
    debugFn();
  } catch (err) {
    console.warn('Self debug failed', err);
  }
})();
