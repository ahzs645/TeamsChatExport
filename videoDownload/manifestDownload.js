/**
 * Manifest Download Module
 * Downloads video by fetching encrypted segments in parallel and decrypting them
 * using the captured AES-CBC key from the player's crypto.subtle.decrypt calls.
 *
 * Flow:
 * 1. Player briefly plays → crypto.subtle.decrypt is called → we capture key + IV
 * 2. We fetch ALL encrypted segments in parallel (fast, no playback needed)
 * 3. We decrypt each segment with crypto.subtle.decrypt(algo, key, data)
 * 4. Combine init segments + decrypted moof segments → playable MP4
 *
 * This is the fastest approach for DRM-protected content:
 * - No full playback needed (just 1-2 seconds to capture the key)
 * - Parallel downloads (~6 concurrent)
 * - Original quality (no re-encoding)
 */
(() => {
  const MODULE_NAME = 'manifestDownload';
  const MAX_CONCURRENT = 6;
  const SEGMENT_TIMEOUT = 30000;
  const SEGMENT_DURATION_MS = 2000; // 2 seconds per segment (typical for SharePoint)

  // === Crypto Access ===

  const getCrypto = () => window.__videoCryptoCapture || { ready: false };

  const hasCryptoKey = () => getCrypto().ready === true;

  // === Segment Templates ===

  const findSegmentTemplates = () => {
    if (!window.__teamsVideoCapture) return null;
    const tokens = window.__teamsVideoCapture.getTokens?.();
    if (!tokens?.videoTemplate) return null;
    return {
      videoTemplate: tokens.videoTemplate,
      audioTemplate: tokens.audioTemplate
    };
  };

  // === Init Segment Capture ===
  // We need the ftyp+moov init segments which are only available after the player
  // decrypts and feeds them to appendBuffer. We capture them from MSE.

  const getInitSegments = () => {
    // Check if videoDownloadOverride captured them
    if (window.__teamsVideoCapture) {
      const stats = window.__teamsVideoCapture.getStats?.();
      if (stats?.hasVideoInit && stats?.hasAudioInit) {
        return { fromCapture: true };
      }
    }
    // Check our own stored init segments
    return window.__manifestInitSegments || null;
  };

  // === Segment Fetching + Decryption ===

  const fetchAndDecrypt = async (url, cryptoKey, algo) => {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), SEGMENT_TIMEOUT);

    const resp = await fetch(url, { signal: controller.signal });
    clearTimeout(timeout);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

    const encrypted = await resp.arrayBuffer();

    // Decrypt using the captured AES-CBC key
    // Need to recreate the algo object with a fresh IV copy each time
    const decryptAlgo = { name: algo.name, iv: algo.iv || getCrypto().iv };
    const decrypted = await crypto.subtle.decrypt(decryptAlgo, cryptoKey, encrypted);

    return decrypted;
  };

  const fetchAndDecryptWithRetry = async (url, cryptoKey, algo, retries = 2) => {
    for (let attempt = 0; attempt <= retries; attempt++) {
      try {
        return await fetchAndDecrypt(url, cryptoKey, algo);
      } catch (e) {
        if (attempt === retries) throw e;
        await new Promise(r => setTimeout(r, 1000 * (attempt + 1)));
      }
    }
  };

  // === Parallel Download Engine ===

  const downloadAllSegments = async (template, segmentCount, cryptoKey, algo, onProgress, label) => {
    const results = new Array(segmentCount);
    let completed = 0;
    let failed = 0;

    const queue = Array.from({ length: segmentCount }, (_, i) => i);

    const worker = async () => {
      while (queue.length > 0) {
        const i = queue.shift();
        if (i === undefined) break;

        const time = i * SEGMENT_DURATION_MS;
        const url = template.replace('{TIME}', time).replace('%7BTIME%7D', time);

        try {
          results[i] = await fetchAndDecryptWithRetry(url, cryptoKey, algo);
          completed++;
        } catch (e) {
          console.warn(`[manifestDownload] ${label} segment ${i} failed:`, e.message);
          failed++;
          completed++;
        }

        if (onProgress) {
          onProgress({
            stage: label,
            message: `${label}: ${completed}/${segmentCount} segments (${failed} failed)`,
            percent: Math.round(completed / segmentCount * 100),
            completed,
            total: segmentCount,
            failed
          });
        }
      }
    };

    // Run workers in parallel
    const workers = Array.from({ length: MAX_CONCURRENT }, () => worker());
    await Promise.all(workers);

    return results.filter(Boolean);
  };

  // === Init Segment Capture via Brief Playback ===

  const captureInitSegments = () => {
    return new Promise((resolve) => {
      let videoInit = null, audioInit = null;
      const origAppend = SourceBuffer.prototype.appendBuffer;

      SourceBuffer.prototype.appendBuffer = function(data) {
        const arr = new Uint8Array(data instanceof ArrayBuffer ? data : data.buffer);
        if (arr.length >= 8) {
          const boxType = String.fromCharCode(arr[4], arr[5], arr[6], arr[7]);
          const mime = this._mseTrackType || '';
          if (boxType === 'ftyp') {
            if (mime.includes('audio') && !audioInit) {
              audioInit = arr.slice().buffer;
            } else if (!videoInit) {
              videoInit = arr.slice().buffer;
            }
          }
        }
        const result = origAppend.call(this, data);
        // Once we have both, resolve
        if (videoInit && audioInit) {
          SourceBuffer.prototype.appendBuffer = origAppend;
          window.__manifestInitSegments = { videoInit, audioInit };
          resolve({ videoInit, audioInit });
        }
        return result;
      };

      // If the video was already playing, seek to trigger new init
      const video = document.querySelector('video');
      if (video) {
        video.currentTime = 0;
        // Timeout: if we don't get init in 10 seconds, resolve with what we have
        setTimeout(() => {
          SourceBuffer.prototype.appendBuffer = origAppend;
          if (videoInit || audioInit) {
            window.__manifestInitSegments = { videoInit, audioInit };
            resolve({ videoInit, audioInit });
          } else {
            resolve(null);
          }
        }, 10000);
      } else {
        resolve(null);
      }
    });
  };

  // === Main Download Function ===

  const isAvailable = () => {
    return !!(findSegmentTemplates() && hasCryptoKey());
  };

  /**
   * Download the full video by fetching + decrypting all segments in parallel.
   */
  const download = async (onProgress) => {
    const templates = findSegmentTemplates();
    if (!templates) {
      return { success: false, error: 'No segment templates found. Play the video briefly first.' };
    }

    const cryptoData = getCrypto();
    if (!cryptoData.ready) {
      return { success: false, error: 'No decryption key captured. Play the video for a few seconds first.' };
    }

    const video = document.querySelector('video');
    const duration = video?.duration || 0;
    if (!duration) {
      return { success: false, error: 'Cannot determine video duration.' };
    }

    const segmentCount = Math.ceil(duration / (SEGMENT_DURATION_MS / 1000));

    if (onProgress) onProgress({
      stage: 'init',
      message: `Preparing: ${segmentCount} segments to download (~${Math.round(duration / 60)} min video)`,
      percent: 0
    });

    // Step 1: Capture init segments if we don't have them
    let initSegs = getInitSegments();
    if (!initSegs || !initSegs.videoInit) {
      if (onProgress) onProgress({ stage: 'init', message: 'Capturing init segments (brief play)...', percent: 0 });
      initSegs = await captureInitSegments();
      if (!initSegs || !initSegs.videoInit) {
        return { success: false, error: 'Failed to capture init segments. Try playing the video briefly first.' };
      }
    }

    // Step 2: Download + decrypt all video segments
    if (onProgress) onProgress({
      stage: 'video',
      message: `Downloading ${segmentCount} video segments...`,
      percent: 0
    });

    let videoSegments;
    try {
      videoSegments = await downloadAllSegments(
        templates.videoTemplate,
        segmentCount,
        cryptoData.key,
        cryptoData.algo,
        (p) => {
          if (onProgress) onProgress({
            ...p,
            percent: Math.round(p.percent * 0.6) // 0-60% for video
          });
        },
        'Video'
      );
    } catch (e) {
      return { success: false, error: `Video download failed: ${e.message}` };
    }

    // Step 3: Download + decrypt all audio segments
    if (onProgress) onProgress({ stage: 'audio', message: 'Downloading audio segments...', percent: 60 });

    let audioSegments;
    try {
      audioSegments = await downloadAllSegments(
        templates.audioTemplate,
        segmentCount,
        cryptoData.key,
        cryptoData.algo,
        (p) => {
          if (onProgress) onProgress({
            ...p,
            stage: 'Audio',
            percent: 60 + Math.round(p.percent * 0.3) // 60-90% for audio
          });
        },
        'Audio'
      );
    } catch (e) {
      return { success: false, error: `Audio download failed: ${e.message}` };
    }

    // Step 4: Combine into playable file
    if (onProgress) onProgress({ stage: 'combining', message: 'Combining video + audio...', percent: 90 });

    const fileName = getFileName('mp4');

    // Try MP4 muxer for combined file
    if (window.MP4Muxer && initSegs.videoInit && initSegs.audioInit) {
      try {
        const combined = window.MP4Muxer.mux(
          initSegs.videoInit, videoSegments,
          initSegs.audioInit, audioSegments
        );
        const blob = new Blob([combined], { type: 'video/mp4' });
        triggerDownload(blob, fileName);

        if (onProgress) onProgress({
          stage: 'done',
          message: `Done! ${fileName} (${Math.round(blob.size / 1024 / 1024)}MB)`,
          percent: 100
        });
        return { success: true, fileName, fileSize: blob.size, method: 'muxed' };
      } catch (e) {
        console.warn('[manifestDownload] Mux failed, downloading separately:', e.message);
      }
    }

    // Fallback: download video and audio separately
    const vBlob = new Blob([initSegs.videoInit, ...videoSegments], { type: 'video/mp4' });
    const aBlob = new Blob([initSegs.audioInit, ...audioSegments], { type: 'audio/mp4' });

    triggerDownload(vBlob, getFileName('video.mp4'));
    setTimeout(() => triggerDownload(aBlob, getFileName('audio.mp4')), 500);

    if (onProgress) onProgress({
      stage: 'done',
      message: `Done! Downloaded video (${Math.round(vBlob.size / 1024 / 1024)}MB) + audio (${Math.round(aBlob.size / 1024 / 1024)}MB) separately`,
      percent: 100
    });

    return { success: true, method: 'separate', videoSize: vBlob.size, audioSize: aBlob.size };
  };

  // === Helpers ===

  const getFileName = (ext) => {
    const title = document.querySelector('h1, h2, [class*="videoTitle"] label')
      ?.textContent?.trim()?.replace(/[^a-zA-Z0-9\s-]/g, '')?.trim() || 'recording';
    return `${title}.${ext}`;
  };

  const triggerDownload = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 120000);
  };

  // Expose module
  window.__videoDownloadModules = window.__videoDownloadModules || {};
  window.__videoDownloadModules[MODULE_NAME] = {
    name: MODULE_NAME,
    label: 'Fast Download',
    description: 'Parallel fetch + decrypt (play video briefly first to capture key)',
    isAvailable,
    download,
    // Expose for debugging
    hasCryptoKey,
    findSegmentTemplates,
    getInitSegments
  };

  console.log('[Teams Chat Exporter] Video download module loaded: manifestDownload');
})();
