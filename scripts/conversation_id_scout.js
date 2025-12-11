// Run this in the Teams page console (top frame). It hunts for conversation ids (19:...)
// from DOM, iframes, and localStorage, and tries to pick a "best guess" for the active chat.
(() => {
  const log = (...args) => console.log('[IdScout]', ...args);

  const scores = new Map(); // id -> {sources:Set, score:number}
  const add = (id, source) => {
    if (!id || !id.includes('19:')) return;
    const cleaned = id.match(/(19:[^"'\\s<>]+)/)?.[1];
    if (!cleaned) return;
    if (cleaned.length < 15) return; // filter out timestamps like 19:41:42
    const entry = scores.get(cleaned) || { sources: new Set(), score: 0 };
    entry.sources.add(source);
    let bonus = 1;
    if (/active/.test(source)) bonus += 5;
    if (/data-qa-chat-id|data-fui-tree-item-value|id/.test(source)) bonus += 3;
    if (/aria-selected/.test(source)) bonus += 2;
    if (/localStorage/.test(source)) bonus += 1;
    entry.score += bonus;
    scores.set(cleaned, entry);
  };

  const scanAttributes = (root, source) => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT, null);
    let node;
    let hits = 0;
    while ((node = walker.nextNode()) && hits < 80) {
      for (const attr of Array.from(node.attributes || [])) {
        const val = attr.value;
        if (val && val.includes('19:')) {
          add(val, `${source} attr ${attr.name}`);
          hits++;
        }
      }
    }
  };

  const scanTextSnippets = (root, source) => {
    const text = root.innerText || '';
    const matches = text.match(/19:[0-9A-Za-z@.\-_]+/g) || [];
    matches.slice(0, 30).forEach((m) => add(m, `${source} text`));
  };

  const scanLocalStorage = () => {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      const val = localStorage.getItem(key);
      if (!val) continue;
      if (val.includes('19:')) add(val, `localStorage ${key}`);
      if (key && key.includes('19:')) add(key, `localStorage-key ${key}`);
    }
  };

  // Active chat element datasets
  const active = document.querySelector('[aria-selected=\"true\"][data-qa-chat-id], [data-tid=\"chat-list-item\"][aria-selected=\"true\"], [role=\"treeitem\"][aria-selected=\"true\"]');
  if (active) {
    Object.entries(active.dataset || {}).forEach(([k, v]) => add(v, `active.dataset ${k}`));
    scanAttributes(active, 'active subtree');
  }

  // Whole document
  scanAttributes(document.body || document.documentElement, 'top');
  scanTextSnippets(document.body || document.documentElement, 'top');

  // Iframes (same-origin only)
  document.querySelectorAll('iframe').forEach((frame, idx) => {
    try {
      const doc = frame.contentDocument;
      if (!doc) return;
      scanAttributes(doc.body || doc.documentElement, `iframe#${idx}`);
      scanTextSnippets(doc.body || doc.documentElement, `iframe#${idx}`);
    } catch (_err) {
      // cross-origin, ignore
    }
  });

  // localStorage
  scanLocalStorage();

  const all = Array.from(scores.entries()).map(([id, info]) => ({
    id,
    score: info.score,
    sources: Array.from(info.sources)
  })).sort((a, b) => b.score - a.score);

  log('top candidates (sorted):', all.slice(0, 10));
  if (all[0]) {
    log('best guess:', all[0].id, 'sources:', all[0].sources);
  } else {
    log('no ids found');
  }
  return all;
})();
