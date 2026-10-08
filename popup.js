'use strict';

/*
 * Toolbar popup for Teams Chat Exporter.
 *
 * Talks to content.js with chrome.tabs.sendMessage using the existing actions:
 *   getState, getTranscriptStatus, getVideoModules   (polled every 2 s)
 *   extractActiveChat, extractTranscript, startBatchTranscript, downloadVideo
 *
 * Rendering is split into regions. A region is only touched when its data
 * snapshot changes, and the periodic refresh never touches a region that
 * currently holds keyboard focus, so polling cannot steal focus.
 */

document.addEventListener('DOMContentLoaded', () => {
  const $ = (id) => document.getElementById(id);

  const el = {
    chip: $('pageChip'),
    chipLabel: $('pageChipLabel'),
    status: $('statusLine'),
    cards: $('cards'),
    chatCard: $('chatCard'),
    extractBtn: $('extractActiveChatBtn'),
    extractHint: $('extractHint'),
    videoCard: $('videoCard'),
    videoBtn: $('videoPrimaryBtn'),
    videoHint: $('videoPrimaryHint'),
    videoMore: $('videoMore'),
    videoList: $('videoMethodList'),
    transcriptCard: $('transcriptCard'),
    transcriptStatus: $('transcriptStatus'),
    copyBtn: $('copyTranscriptBtn'),
    vttBtn: $('downloadVttBtn'),
    txtBtn: $('downloadTxtBtn'),
    batchBtn: $('batchTranscriptBtn'),
    openResultsBtn: $('openResultsBtn'),
    settingsHeader: $('settingsHeader'),
    settingsContent: $('settingsContent'),
    maxMode: $('maxMode'),
    pageSize: $('pageSize'),
    maxPages: $('maxPages'),
    maxMessagesInfo: $('maxMessagesInfo'),
  };

  // ---------------------------------------------------------------------------
  // Messaging helpers
  // ---------------------------------------------------------------------------

  const classifyUrl = (url = '') => {
    if (url.includes('teams.microsoft.com') || url.includes('teams.cloud.microsoft')) return 'teams';
    if (url.includes('sharepoint.com') || url.includes('stream.aspx')) return 'stream';
    return 'other';
  };

  const getActiveTab = () => new Promise((resolve) => {
    try {
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => resolve((tabs && tabs[0]) || null));
    } catch (e) {
      resolve(null);
    }
  });

  // Resolves to { ok, response } or { ok: false, error }. Never rejects, never hangs.
  const sendToTab = (tabId, message, timeoutMs = 4000) => new Promise((resolve) => {
    let settled = false;
    let timer = null;
    const done = (result) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(result);
    };
    timer = setTimeout(() => done({ ok: false, error: 'timeout' }), timeoutMs);
    try {
      chrome.tabs.sendMessage(tabId, message, (response) => {
        const err = chrome.runtime.lastError;
        if (err) done({ ok: false, error: err.message || 'unreachable' });
        else done({ ok: true, response });
      });
    } catch (e) {
      done({ ok: false, error: e.message });
    }
  });

  // ---------------------------------------------------------------------------
  // Status line (replaces alert())
  // ---------------------------------------------------------------------------

  const setStatus = (text, tone = 'info') => {
    el.status.dataset.tone = tone;
    el.status.textContent = text || '';
  };

  let closeTimer = null;
  const closeSoon = (ms = 1500) => {
    clearTimeout(closeTimer);
    closeTimer = setTimeout(() => window.close(), ms);
  };

  // ---------------------------------------------------------------------------
  // State
  // ---------------------------------------------------------------------------

  let currentTab = null;
  let model = {
    site: 'loading',
    reachable: false,
    chatName: null,
    recording: false,
    hasVideo: false,
    transcriptReady: false,
    modules: null,
  };
  const busy = { extract: false, video: null, transcript: null, batch: false };

  // Plain-language names for the video download modules reported by content.js.
  const METHOD_INFO = {
    directDownload: {
      name: 'Original file',
      desc: 'Downloads the original MP4 as it was uploaded.',
      unavailable: 'Needs download permission on this recording.',
    },
    manifestDownload: {
      name: 'Fast download',
      desc: 'Downloads the video in parallel pieces. Quickest for long recordings.',
      unavailable: 'Play the video for a few seconds to turn this on.',
    },
    mseCaptureDownload: {
      name: 'Save what was played',
      desc: 'Saves the video your browser has already loaded, without downloading it again.',
      unavailable: 'Needs the video player on the page.',
    },
    captureStreamDownload: {
      name: 'Record playback',
      desc: 'Plays the video at high speed and records it. Keep this tab in front.',
      unavailable: 'Needs the video player on the page.',
    },
  };

  const methodInfo = (mod) => {
    const known = METHOD_INFO[mod.name];
    return {
      name: known ? known.name : (mod.label || mod.name),
      desc: known ? known.desc : (mod.description || ''),
      unavailable: known ? known.unavailable : 'Not available on this page right now.',
      tech: mod.label || mod.name,
    };
  };

  // ---------------------------------------------------------------------------
  // Region rendering with snapshot + focus guard
  // ---------------------------------------------------------------------------

  const snapshots = {};
  const renderRegion = (name, container, data, paint, force) => {
    const snap = JSON.stringify(data);
    if (snapshots[name] === snap) return;
    // Polling must never disturb a control the user is on. Leave the snapshot
    // stale so the next refresh after focus leaves picks the change up.
    if (!force && container && container.matches(':focus-within')) return;
    snapshots[name] = snap;
    paint(data);
  };

  const setDisabled = (btn, disabled) => {
    // aria-disabled keeps the button focusable (no focus loss, hint stays reachable).
    if (disabled) btn.setAttribute('aria-disabled', 'true');
    else btn.removeAttribute('aria-disabled');
  };
  const isDisabled = (btn) => btn.getAttribute('aria-disabled') === 'true';

  const setBusy = (btn, isBusy) => {
    if (isBusy) btn.setAttribute('aria-busy', 'true');
    else btn.removeAttribute('aria-busy');
  };

  const setPrimary = (card, btn, primary) => {
    card.classList.toggle('card-primary', primary);
    btn.classList.toggle('btn-primary', primary);
    btn.classList.toggle('btn-secondary', !primary);
  };

  const derive = (m) => {
    const showMedia = m.reachable && (m.recording || m.hasVideo);
    const mode = showMedia && (m.recording || !m.chatName) ? 'recording' : 'chat';
    const chatVisible = mode === 'chat' || (m.site === 'teams' && !!m.chatName);
    return { showMedia, mode, chatVisible };
  };

  const renderChip = (m, d) => {
    let tone = 'neutral';
    let label;
    if (m.site === 'loading') label = 'Checking page…';
    else if (m.site === 'other') label = 'Not a Teams page';
    else if (!m.reachable) { tone = 'warning'; label = 'Reload the page to connect'; }
    else if (d.mode === 'recording') { tone = 'accent'; label = 'Meeting recording'; }
    else if (m.site === 'teams' && m.chatName) { tone = 'accent'; label = `Teams chat: ${m.chatName}`; }
    else if (m.site === 'teams') label = 'Teams: no chat open';
    else label = 'No chat or recording here';

    renderRegion('chip', null, { tone, label }, ({ tone: t, label: l }) => {
      el.chip.dataset.tone = t;
      el.chipLabel.textContent = l;
      el.chip.title = l;
    });
  };

  const renderOrder = (d, force) => {
    renderRegion('order', el.cards, { mode: d.mode }, ({ mode }) => {
      const order = mode === 'recording'
        ? [el.videoCard, el.transcriptCard, el.chatCard]
        : [el.chatCard, el.videoCard, el.transcriptCard];
      order.forEach((card) => el.cards.appendChild(card));
    }, force);
  };

  const renderChat = (m, d, force) => {
    let hint = '';
    if (m.site === 'loading') hint = '';
    else if (m.site === 'other') hint = 'Open a chat in Microsoft Teams on the web to export it.';
    else if (m.site !== 'teams') hint = 'Chat export works on teams.microsoft.com. Open the chat there.';
    else if (!m.reachable) hint = "Can't reach this tab. Reload the Teams page, then open this popup again.";
    else if (!m.chatName) hint = 'Open a chat or channel in Teams first.';

    const data = {
      visible: d.chatVisible,
      primary: d.mode === 'chat',
      enabled: m.site === 'teams' && m.reachable && !!m.chatName,
      busy: busy.extract,
      hint,
    };
    renderRegion('chat', el.chatCard, data, (v) => {
      el.chatCard.hidden = !v.visible;
      setPrimary(el.chatCard, el.extractBtn, v.primary);
      el.extractBtn.textContent = v.busy ? 'Extracting…' : 'Extract this chat';
      setDisabled(el.extractBtn, !v.enabled || v.busy);
      setBusy(el.extractBtn, v.busy);
      el.extractHint.textContent = v.hint;
      el.extractHint.hidden = !v.hint;
    }, force);
  };

  const renderVideo = (m, d, force) => {
    const modules = m.modules;
    const chosen = modules ? modules.find((mod) => mod.available) : null;
    const data = {
      visible: d.showMedia,
      primary: d.mode === 'recording',
      modules,
      chosen: chosen ? chosen.name : null,
      busy: busy.video,
    };
    renderRegion('video', el.videoCard, data, (v) => {
      el.videoCard.hidden = !v.visible;
      setPrimary(el.videoCard, el.videoBtn, v.primary);

      const chosenMod = v.modules && v.modules.find((mod) => mod.name === v.chosen);
      el.videoBtn.textContent = v.busy ? 'Starting download…' : 'Download video';
      setDisabled(el.videoBtn, !chosenMod || !!v.busy);
      setBusy(el.videoBtn, v.busy === v.chosen && !!v.busy);
      el.videoBtn.dataset.method = chosenMod ? chosenMod.name : '';

      if (!v.modules) {
        el.videoHint.textContent = "Video download isn't available on this page.";
      } else if (!chosenMod) {
        el.videoHint.textContent = 'Play the recording for a few seconds, then try again.';
      } else {
        const info = methodInfo(chosenMod);
        el.videoHint.textContent = `Uses “${info.name}”: ${info.desc}`;
      }

      const others = (v.modules || []).filter((mod) => mod.name !== v.chosen);
      el.videoMore.hidden = others.length === 0;
      // Update method rows in place when the set of methods is unchanged, so a
      // focused row survives a forced render (e.g. after it was clicked).
      const existing = Array.from(el.videoList.querySelectorAll('button.method'));
      const sameSet = existing.length === others.length
        && existing.every((btn, i) => btn.dataset.method === others[i].name);
      if (!sameSet) {
        el.videoList.replaceChildren(...others.map((mod) => {
          const li = document.createElement('li');
          const btn = document.createElement('button');
          btn.type = 'button';
          btn.className = 'method';
          btn.dataset.method = mod.name;
          const name = document.createElement('span');
          name.className = 'method-name';
          const desc = document.createElement('span');
          desc.className = 'method-desc';
          const tech = document.createElement('span');
          tech.className = 'method-tech';
          btn.append(name, desc, tech);
          li.appendChild(btn);
          return li;
        }));
      }
      el.videoList.querySelectorAll('button.method').forEach((btn, i) => {
        const mod = others[i];
        const info = methodInfo(mod);
        btn.querySelector('.method-name').textContent = v.busy === mod.name ? 'Starting download…' : info.name;
        btn.querySelector('.method-desc').textContent = mod.available ? info.desc : info.unavailable;
        btn.querySelector('.method-tech').textContent = info.tech;
        setDisabled(btn, !mod.available || !!v.busy);
        setBusy(btn, v.busy === mod.name);
      });
    }, force);
  };

  const renderTranscript = (m, d, force) => {
    const data = {
      visible: d.showMedia,
      ready: m.transcriptReady,
      busy: busy.transcript,
      batchBusy: busy.batch,
    };
    renderRegion('transcript', el.transcriptCard, data, (v) => {
      el.transcriptCard.hidden = !v.visible;
      el.transcriptStatus.dataset.tone = v.ready ? 'success' : 'warning';
      el.transcriptStatus.textContent = v.ready
        ? 'Transcript ready'
        : 'Play the recording for a few seconds to load the transcript.';

      const labels = {
        copy: [el.copyBtn, 'Copy transcript'],
        vtt: [el.vttBtn, 'Download subtitles (.vtt)'],
        txt: [el.txtBtn, 'Download text (.txt)'],
      };
      Object.entries(labels).forEach(([key, [btn, label]]) => {
        btn.textContent = v.busy === key ? 'Getting transcript…' : label;
        setDisabled(btn, !v.ready || !!v.busy);
        setBusy(btn, v.busy === key);
      });

      el.batchBtn.textContent = v.batchBusy ? 'Opening…' : 'Download all meeting transcripts';
      setDisabled(el.batchBtn, v.batchBusy);
      setBusy(el.batchBtn, v.batchBusy);
    }, force);
  };

  const render = (force = false) => {
    const d = derive(model);
    renderChip(model, d);
    renderOrder(d, force);
    renderChat(model, d, force);
    renderVideo(model, d, force);
    renderTranscript(model, d, force);
  };

  // ---------------------------------------------------------------------------
  // Polling
  // ---------------------------------------------------------------------------

  const NO_CHAT = /^no chat (open|selected)$/i;

  let polling = false;
  const poll = async () => {
    if (polling) return;
    polling = true;
    try {
      const tab = await getActiveTab();
      currentTab = tab;
      const site = classifyUrl(tab && tab.url);
      const next = {
        site,
        reachable: false,
        chatName: null,
        recording: false,
        hasVideo: false,
        transcriptReady: false,
        modules: null,
      };

      if (tab && site !== 'other') {
        const [state, transcript, video] = await Promise.all([
          sendToTab(tab.id, { action: 'getState' }, 3000),
          sendToTab(tab.id, { action: 'getTranscriptStatus' }, 3000),
          sendToTab(tab.id, { action: 'getVideoModules' }, 3000),
        ]);
        next.reachable = state.ok || transcript.ok || video.ok;

        const name = state.ok && state.response && typeof state.response.currentChat === 'string'
          ? state.response.currentChat.trim() : '';
        if (site === 'teams' && name && !NO_CHAT.test(name)) next.chatName = name;

        const t = transcript.ok && transcript.response;
        if (t) {
          next.recording = !!t.isVideoPage;
          next.hasVideo = !!t.hasVideo;
          next.transcriptReady = !!t.available;
        }

        const mods = video.ok && video.response && video.response.modules;
        if (Array.isArray(mods)) {
          next.modules = mods.map((mod) => ({
            name: String(mod.name || ''),
            label: String(mod.label || ''),
            description: String(mod.description || ''),
            available: !!mod.available,
          }));
        }
      }

      model = next;
      render();
    } finally {
      polling = false;
    }
  };

  // ---------------------------------------------------------------------------
  // Actions
  // ---------------------------------------------------------------------------

  const ackFailed = (res) => {
    if (!res.ok) return true;
    const r = res.response;
    return !r || r.success === false || r.status === 'error' || !!r.error;
  };

  const unreachableMessage = (res) => (res.ok && res.response && res.response.error)
    ? res.response.error
    : "The page didn't respond. Reload the Teams tab and try again.";

  el.extractBtn.addEventListener('click', async () => {
    if (isDisabled(el.extractBtn) || !currentTab) return;
    busy.extract = true;
    render(true);
    setStatus('Extracting messages…');

    const res = await sendToTab(currentTab.id, { action: 'extractActiveChat' }, 8000);
    if (ackFailed(res)) {
      busy.extract = false;
      render(true);
      setStatus(unreachableMessage(res), 'error');
      return;
    }
    setStatus('Extracting messages… Progress is shown on the page and the viewer opens when it finishes.', 'success');
    closeSoon(2000);
  });

  const startVideoDownload = async (method, btn) => {
    if (!currentTab || !method || isDisabled(btn)) return;
    busy.video = method;
    render(true);
    setStatus('Starting video download…');

    const res = await sendToTab(currentTab.id, { action: 'downloadVideo', data: { method } });
    if (ackFailed(res)) {
      busy.video = null;
      render(true);
      setStatus(unreachableMessage(res), 'error');
      return;
    }
    setStatus('Video download started. Progress is shown on the page; keep the tab open until it finishes.', 'success');
    closeSoon(2500);
  };

  el.videoBtn.addEventListener('click', () => startVideoDownload(el.videoBtn.dataset.method, el.videoBtn));
  el.videoList.addEventListener('click', (event) => {
    const btn = event.target.closest('button.method');
    if (btn) startVideoDownload(btn.dataset.method, btn);
  });

  const saveFile = (content, filename, type) => {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const withTranscript = async (key, btn, onData) => {
    if (isDisabled(btn) || !currentTab) return;
    busy.transcript = key;
    render(true);
    setStatus('Getting transcript…');

    const res = await sendToTab(currentTab.id, { action: 'extractTranscript' }, 30000);
    busy.transcript = null;
    render(true);

    const r = res.ok ? res.response : null;
    if (!r || !r.success) {
      setStatus((r && r.error) || "Couldn't get the transcript. Play the recording for a few seconds and try again.", 'error');
      return;
    }
    try {
      await onData(r);
    } catch (e) {
      setStatus(`Something went wrong: ${e.message || e}`, 'error');
    }
  };

  el.copyBtn.addEventListener('click', () => withTranscript('copy', el.copyBtn, async (r) => {
    try {
      await navigator.clipboard.writeText(r.vtt);
    } catch (e) {
      setStatus("Couldn't copy to the clipboard. Try one of the download buttons instead.", 'error');
      return;
    }
    setStatus('Transcript copied to the clipboard.', 'success');
  }));

  el.vttBtn.addEventListener('click', () => withTranscript('vtt', el.vttBtn, (r) => {
    const filename = `transcript-${r.title}.vtt`;
    saveFile(r.vtt, filename, 'text/vtt;charset=utf-8');
    setStatus(`Saved ${filename}.`, 'success');
  }));

  el.txtBtn.addEventListener('click', () => withTranscript('txt', el.txtBtn, (r) => {
    const filename = `transcript-${r.title}.txt`;
    saveFile(r.txt, filename, 'text/plain;charset=utf-8');
    setStatus(`Saved ${filename}.`, 'success');
  }));

  el.batchBtn.addEventListener('click', async () => {
    if (isDisabled(el.batchBtn) || !currentTab) return;
    if (classifyUrl(currentTab.url) === 'other') {
      setStatus('Open a Teams meeting recap or recording to download transcripts in bulk.', 'error');
      return;
    }
    busy.batch = true;
    render(true);
    const res = await sendToTab(currentTab.id, { action: 'startBatchTranscript' });
    if (ackFailed(res)) {
      busy.batch = false;
      render(true);
      setStatus(unreachableMessage(res), 'error');
      return;
    }
    setStatus('Batch download panel opened on the page.', 'success');
    closeSoon(1500);
  });

  el.openResultsBtn.addEventListener('click', () => {
    chrome.tabs.create({ url: chrome.runtime.getURL('results.html') });
    window.close();
  });

  // ---------------------------------------------------------------------------
  // Advanced settings
  // ---------------------------------------------------------------------------

  const DEFAULTS = { pageSize: 200, maxPages: 15 };
  const FULL = { pageSize: 500, maxPages: 200 };

  el.settingsHeader.addEventListener('click', () => {
    const expanded = el.settingsHeader.getAttribute('aria-expanded') === 'true';
    el.settingsHeader.setAttribute('aria-expanded', String(!expanded));
    el.settingsContent.hidden = expanded;
  });

  const readInputs = () => ({
    pageSize: parseInt(el.pageSize.value, 10) || DEFAULTS.pageSize,
    maxPages: parseInt(el.maxPages.value, 10) || DEFAULTS.maxPages,
  });

  const updateReadout = () => {
    const { pageSize, maxPages } = readInputs();
    el.maxMessagesInfo.textContent = `Up to ${(pageSize * maxPages).toLocaleString()} messages`;
    el.maxMode.checked = pageSize === FULL.pageSize && maxPages === FULL.maxPages;
  };

  const loadSettings = () => {
    chrome.storage.local.get(['teamsChatApiPageSize', 'teamsChatApiMaxPages'], (result) => {
      el.pageSize.value = result.teamsChatApiPageSize || DEFAULTS.pageSize;
      el.maxPages.value = result.teamsChatApiMaxPages || DEFAULTS.maxPages;
      updateReadout();
    });
  };

  const saveSettings = () => {
    const pageSize = Math.max(1, Math.min(500, parseInt(el.pageSize.value, 10) || DEFAULTS.pageSize));
    const maxPages = Math.max(1, Math.min(200, parseInt(el.maxPages.value, 10) || DEFAULTS.maxPages));
    el.pageSize.value = pageSize;
    el.maxPages.value = maxPages;
    chrome.storage.local.set({ teamsChatApiPageSize: pageSize, teamsChatApiMaxPages: maxPages });
    updateReadout();
  };

  el.maxMode.addEventListener('change', () => {
    const preset = el.maxMode.checked ? FULL : DEFAULTS;
    el.pageSize.value = preset.pageSize;
    el.maxPages.value = preset.maxPages;
    saveSettings();
  });

  el.pageSize.addEventListener('change', saveSettings);
  el.maxPages.addEventListener('change', saveSettings);
  el.pageSize.addEventListener('input', updateReadout);
  el.maxPages.addEventListener('input', updateReadout);

  // ---------------------------------------------------------------------------
  // Init
  // ---------------------------------------------------------------------------

  render(true);
  loadSettings();
  poll();

  const refresh = setInterval(poll, 2000);
  window.addEventListener('beforeunload', () => {
    clearInterval(refresh);
    clearTimeout(closeTimer);
  });
});
