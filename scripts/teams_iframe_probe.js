// Step 1: List iframes and same-origin reachability
(() => {
  const framesList = Array.from(document.querySelectorAll('iframe')).map((f, i) => {
    const sameOrigin = (() => {
      try {
        return !!f.contentDocument;
      } catch (_e) {
        return false;
      }
    })();
    return { i, src: f.src, sameOrigin };
  });
  console.log('[Extractor] Frames:', framesList);
})();

// Step 2: To probe a specific same-origin frame, replace INDEX below and run this separately:
/*
(() => {
  const INDEX = 0; // <-- change this to the frame index you want to probe
  const frame = window.frames[INDEX];
  if (!frame) {
    console.warn('No frame at index', INDEX);
    return;
  }
  frame.eval(`(() => {
    const results = {};
    const activeItem = document.querySelector('[data-tid="chat-list-item"][aria-selected="true"], [aria-selected="true"][data-qa-chat-id], [role="treeitem"][aria-selected="true"]');
    results.activeItemDataset = activeItem ? { ...activeItem.dataset } : null;
    results.activeItemAttrs = activeItem ? Array.from(activeItem.attributes).reduce((acc, a) => ({ ...acc, [a.name]: a.value }), {}) : null;
    results.activeItemText = activeItem ? (activeItem.textContent || '').slice(0, 200) : null;
    const anchorWith19 = Array.from(document.querySelectorAll('a')).find(a => (a.href || '').includes('19:'));
    results.linkHref = anchorWith19?.href || null;
    results.locationHref = window.location.href;
    results.hiddenCacheSample = (() => {
      const div = document.getElementById('teams-extractor-message-api-cache');
      if (!div || !div.textContent) return null;
      try { const parsed = JSON.parse(div.textContent); return Array.isArray(parsed) ? parsed.slice(-2) : null; } catch { return 'parse-error'; }
    })();
    results.memoryCacheSize = Array.isArray(window.__teamsExtractorMessageApiCache) ? window.__teamsExtractorMessageApiCache.length : 0;
    console.log('[Extractor iframe debug]', ${'${INDEX}'}, results);
    return results;
  })();`);
})();
*/
