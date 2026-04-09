/**
 * Direct Download Module
 * Downloads video via SharePoint Drive API @content.downloadUrl.
 * Works when user has download permission on the file.
 * Fastest approach - instant download of the original MP4.
 */
(() => {
  const MODULE_NAME = 'directDownload';

  /**
   * Check if direct download is available.
   * Reads drive item data captured by transcriptAPIFetcher.js.
   */
  const isAvailable = () => {
    const div = document.getElementById('video-drive-data');
    if (!div) return false;
    try {
      const data = JSON.parse(div.getAttribute('data-drive-item') || '{}');
      return !!(data.driveId && data.itemId && data.apiBase);
    } catch (e) {
      return false;
    }
  };

  /**
   * Get the drive item metadata from the hidden div.
   */
  const getDriveItem = () => {
    const div = document.getElementById('video-drive-data');
    if (!div) return null;
    try {
      const data = JSON.parse(div.getAttribute('data-drive-item') || '{}');
      if (data.driveId && data.itemId) return data;
    } catch (e) {}
    return null;
  };

  /**
   * Get a SharePoint token for the given host.
   */
  const getToken = (host) => {
    const tokens = window.__sharePointTokens || {};
    const entry = tokens[host];
    if (entry && entry.token && (Date.now() - entry.capturedAt < 30 * 60 * 1000)) {
      return entry.token;
    }
    return null;
  };

  /**
   * Attempt direct download. Returns a result object.
   * @param {Function} onProgress - optional progress callback
   * @returns {Promise<{success: boolean, error?: string, downloadUrl?: string, fileName?: string}>}
   */
  const download = async (onProgress) => {
    const driveItem = getDriveItem();
    if (!driveItem || !driveItem.apiBase) {
      return { success: false, error: 'No drive item data available' };
    }

    const token = getToken(driveItem.host);
    if (!token) {
      return { success: false, error: 'No auth token available' };
    }

    if (onProgress) onProgress({ stage: 'fetching', message: 'Getting download URL...' });

    try {
      const resp = await fetch(driveItem.apiBase, {
        headers: { 'Authorization': token, 'Accept': 'application/json' }
      });

      if (!resp.ok) {
        if (resp.status === 403) {
          return { success: false, error: 'Access denied - you may not have download permission' };
        }
        return { success: false, error: `API returned ${resp.status}` };
      }

      const data = await resp.json();
      const downloadUrl = data['@content.downloadUrl'];

      if (!downloadUrl) {
        return { success: false, error: 'No download URL in API response' };
      }

      const fileName = data.name || 'recording.mp4';
      const fileSize = data.size || 0;

      if (onProgress) onProgress({
        stage: 'downloading',
        message: `Downloading ${fileName} (${Math.round(fileSize / 1024 / 1024)}MB)...`,
        fileName,
        fileSize
      });

      // Trigger browser-native download
      window.open(downloadUrl, '_blank');

      return { success: true, downloadUrl, fileName, fileSize };
    } catch (err) {
      return { success: false, error: err.message };
    }
  };

  /**
   * Get the download URL without triggering download.
   * Used by the popup to open via chrome.tabs.create.
   */
  const getDownloadUrl = async () => {
    const driveItem = getDriveItem();
    if (!driveItem || !driveItem.apiBase) return null;

    const token = getToken(driveItem.host);
    if (!token) return null;

    try {
      const resp = await fetch(driveItem.apiBase, {
        headers: { 'Authorization': token, 'Accept': 'application/json' }
      });
      if (!resp.ok) return null;
      const data = await resp.json();
      return {
        downloadUrl: data['@content.downloadUrl'] || null,
        fileName: data.name,
        fileSize: data.size
      };
    } catch (e) {
      return null;
    }
  };

  // Expose module
  window.__videoDownloadModules = window.__videoDownloadModules || {};
  window.__videoDownloadModules[MODULE_NAME] = {
    name: MODULE_NAME,
    label: 'Direct Download',
    description: 'Download original MP4 via SharePoint API (requires download permission)',
    isAvailable,
    download,
    getDownloadUrl
  };

  console.log('[Teams Chat Exporter] Video download module loaded: directDownload');
})();
