/**
 * Video Download Coordinator
 * Orchestrates the download modules, trying them in priority order.
 * Exposes a unified API via window.__videoDownloadCoordinator.
 * Also handles communication with the content script via CustomEvents.
 */
(() => {
  // Priority order: fastest/best quality first, most compatible last
  const METHOD_PRIORITY = ['directDownload', 'mseCaptureDownload', 'manifestDownload', 'captureStreamDownload'];

  /**
   * Get all loaded download modules.
   */
  const getModules = () => window.__videoDownloadModules || {};

  /**
   * Get available modules (that can work on this page).
   */
  const getAvailableModules = () => {
    const modules = getModules();
    const available = [];
    for (const name of METHOD_PRIORITY) {
      const mod = modules[name];
      if (mod) {
        try {
          available.push({
            name: mod.name,
            label: mod.label,
            description: mod.description,
            available: mod.isAvailable()
          });
        } catch (e) {
          available.push({ name: mod.name, label: mod.label, available: false });
        }
      }
    }
    return available;
  };

  /**
   * Try to download using the best available method.
   * @param {Function} onProgress - progress callback
   * @param {string} preferredMethod - optional method name to use
   * @returns {Promise<{success: boolean, method?: string, error?: string}>}
   */
  const download = async (onProgress, preferredMethod) => {
    const modules = getModules();

    // If a preferred method is specified, use it directly
    if (preferredMethod && modules[preferredMethod]) {
      const mod = modules[preferredMethod];
      if (onProgress) onProgress({ stage: 'starting', message: `Using ${mod.label}...`, method: mod.name });
      const result = await mod.download(onProgress);
      return { ...result, method: mod.name };
    }

    // Try each method in priority order
    for (const name of METHOD_PRIORITY) {
      const mod = modules[name];
      if (!mod) continue;

      try {
        if (!mod.isAvailable()) {
          console.log(`[VideoCoordinator] ${name}: not available, skipping`);
          continue;
        }

        console.log(`[VideoCoordinator] Trying ${name}...`);
        if (onProgress) onProgress({ stage: 'trying', message: `Trying ${mod.label}...`, method: name });

        const result = await mod.download(onProgress);
        if (result.success) {
          console.log(`[VideoCoordinator] ${name}: success`);
          return { ...result, method: name };
        }

        console.log(`[VideoCoordinator] ${name}: failed - ${result.error}`);
      } catch (err) {
        console.warn(`[VideoCoordinator] ${name}: error - ${err.message}`);
      }
    }

    return { success: false, error: 'All download methods failed' };
  };

  /**
   * Stop any active download (mainly for captureStream).
   */
  const stop = () => {
    const modules = getModules();
    for (const mod of Object.values(modules)) {
      if (mod.stop) mod.stop();
    }
  };

  // Expose coordinator API
  window.__videoDownloadCoordinator = {
    getModules,
    getAvailableModules,
    download,
    stop,
    METHOD_PRIORITY
  };

  // === CustomEvent interface for content script communication ===
  document.addEventListener('tceVideoDownloadCommand', async (e) => {
    const { command, data } = e.detail || {};
    let result;

    switch (command) {
      case 'getAvailableModules':
        result = getAvailableModules();
        break;

      case 'download':
        result = await download(
          (progress) => {
            document.dispatchEvent(new CustomEvent('tceVideoDownloadProgress', {
              detail: progress
            }));
          },
          data?.method
        );
        break;

      case 'stop':
        stop();
        result = { stopped: true };
        break;

      case 'getStatus':
        const modules = getModules();
        const captureStream = modules.captureStreamDownload;
        result = captureStream?.getStatus?.() || { isRecording: false };
        break;

      case 'getDownloadUrl':
        // For popup: get direct download URL without triggering download
        const directMod = getModules().directDownload;
        if (directMod?.getDownloadUrl) {
          result = await directMod.getDownloadUrl();
        } else {
          result = null;
        }
        break;

      default:
        result = { error: 'Unknown command: ' + command };
    }

    document.dispatchEvent(new CustomEvent('tceVideoDownloadResponse', {
      detail: { command, result }
    }));
  });

  // === Progress panel (page world) ===
  // Uses the shared panel helper (ui/tceUI.js, injected into this world by
  // content.js) so it matches the content-script panels and lives in the same
  // top-right stack. Always has a Close button; failures show a readable error.
  const PROGRESS_PANEL_ID = 'tce-download-progress';

  const describeError = (err) => {
    const msg = typeof err === 'string' ? err : ((err && (err.message || err.error)) || '');
    if (!msg || msg === 'All download methods failed') {
      return 'None of the download methods worked for this video. Play the video for a few seconds, then try again, or use Record Stream.';
    }
    return msg;
  };

  const showProgressPanel = (method) => {
    const ui = window.__tceUI;
    let panel = null;
    let statusEl = null;
    let bar = null;
    let detailEl = null;

    if (ui) {
      document.getElementById(PROGRESS_PANEL_ID)?.remove();
      panel = ui.createPanel({ id: PROGRESS_PANEL_ID, title: 'Downloading video', wide: true, focus: false });
      statusEl = ui.h('div', { class: 'tce-status', role: 'status', 'aria-live': 'polite', text: 'Starting…' });
      bar = ui.progressBar({ label: 'Download progress', value: null });
      detailEl = ui.h('p', { class: 'tce-hint', text: 'Keep this tab open until the download finishes.' });
      const stopBtn = ui.button({ label: 'Stop', icon: 'stop', variant: 'secondary' });
      stopBtn.addEventListener('click', () => {
        stop();
        stopBtn.disabled = true;
        statusEl.textContent = 'Stopping…';
      });
      panel.body.append(bar.el, statusEl, detailEl);
      panel.footer.append(stopBtn);
    } else {
      console.warn('[VideoCoordinator] UI helpers not loaded; progress is logged to the console only');
    }

    const setStatus = (text, kind) => {
      if (!statusEl) return;
      statusEl.textContent = text;
      statusEl.className = 'tce-status' + (kind ? ' tce-status--' + kind : '');
    };

    const fail = (err) => {
      const message = describeError(err);
      console.error('[VideoCoordinator] Download error:', err);
      if (!panel || panel.closed) {
        if (ui) ui.toast({ kind: 'error', title: 'Video download failed', message });
        return;
      }
      panel.titleEl.textContent = 'Video download failed';
      setStatus(message, 'error');
      if (bar) bar.el.remove();
      if (detailEl) detailEl.textContent = 'You can close this panel and try another method from the extension popup.';
      panel.footer.replaceChildren();
      panel.closeBtn.focus();
    };

    download(
      (progress) => {
        setStatus(progress.message || progress.stage || 'Working…');
        if (bar) bar.set(progress.percent !== undefined ? progress.percent : null, 100);
      },
      method
    ).then((result) => {
      console.log('[VideoCoordinator] Download result:', result);
      if (!result || !result.success) {
        fail(result);
        return;
      }
      // manifestDownload shows its own save panel; other methods have already
      // handed the file to the browser.
      if (panel) panel.close('done');
      if (ui && result.method !== 'manifestDownload') {
        ui.toast({ kind: 'success', title: 'Video download complete', message: 'Check your browser downloads.' });
      }
    }).catch(fail);
  };

  // === DOM-based command interface (for content script communication) ===
  // Content script writes commands to a hidden div since inline scripts are blocked by CSP.
  let lastCmdTimestamp = '0';
  setInterval(() => {
    const cmdDiv = document.getElementById('tce-video-cmd');
    if (!cmdDiv) return;
    const ts = cmdDiv.getAttribute('data-timestamp') || '0';
    if (ts === lastCmdTimestamp) return;
    lastCmdTimestamp = ts;

    const command = cmdDiv.getAttribute('data-command');
    const method = cmdDiv.getAttribute('data-method') || '';
    cmdDiv.removeAttribute('data-command');

    if (command === 'download') {
      console.log('[VideoCoordinator] Download command received, method:', method || 'auto');

      showProgressPanel(method || undefined);
    } else if (command === 'stop') {
      stop();
    }
  }, 500);

  console.log('[Teams Chat Exporter] Video download coordinator loaded');
})();
