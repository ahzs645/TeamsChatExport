/**
 * Video Download Coordinator
 * Orchestrates the download modules, trying them in priority order.
 * Exposes a unified API via window.__videoDownloadCoordinator.
 * Also handles communication with the content script via CustomEvents.
 */
(() => {
  // Priority order: fastest/best quality first, most compatible last
  const METHOD_PRIORITY = ['directDownload', 'manifestDownload', 'captureStreamDownload'];

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

  console.log('[Teams Chat Exporter] Video download coordinator loaded');
})();
