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

      // Create visible progress panel
      let panel = document.getElementById('tce-download-progress');
      if (!panel) {
        panel = document.createElement('div');
        panel.id = 'tce-download-progress';
        panel.style.cssText = 'position:fixed;top:10px;left:50%;transform:translateX(-50%);z-index:99999;padding:16px 24px;background:rgba(0,0,0,0.92);color:white;border-radius:10px;font-family:monospace;font-size:13px;min-width:500px;text-align:center;';
        panel.innerHTML = '<div id="tce-dl-status">Starting...</div>' +
          '<div style="background:#333;height:20px;border-radius:10px;margin:8px 0;overflow:hidden">' +
          '<div id="tce-dl-bar" style="background:#28a745;height:100%;width:0%;transition:width 0.3s"></div></div>' +
          '<div id="tce-dl-detail"></div>';
        document.body.appendChild(panel);
      }

      download(
        (progress) => {
          const statusEl = document.getElementById('tce-dl-status');
          const barEl = document.getElementById('tce-dl-bar');
          if (statusEl) statusEl.textContent = progress.message || progress.stage || 'working...';
          if (barEl && progress.percent !== undefined) barEl.style.width = progress.percent + '%';
        },
        method || undefined
      ).then((result) => {
        console.log('[VideoCoordinator] Download result:', result);
        // Panel with save buttons is created by manifestDownload itself
        const progressPanel = document.getElementById('tce-download-progress');
        if (progressPanel) progressPanel.remove();
      }).catch((err) => {
        const statusEl = document.getElementById('tce-dl-status');
        if (statusEl) statusEl.textContent = 'ERROR: ' + err.message;
        console.error('[VideoCoordinator] Download error:', err);
      });
    } else if (command === 'stop') {
      stop();
    }
  }, 500);

  console.log('[Teams Chat Exporter] Video download coordinator loaded');
})();
