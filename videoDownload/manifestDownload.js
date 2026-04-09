/**
 * Manifest Download Module
 * Downloads video by fetching the videomanifest URL and downloading all segments in parallel.
 * This is the most popular online method (used by ms-teams-sharepoint-downloader, FFmpeg approach, etc.)
 * The manifest URL contains embedded auth tokens, so no extra auth is needed.
 * Works even when download permission is disabled.
 */
(() => {
  const MODULE_NAME = 'manifestDownload';
  const MAX_CONCURRENT = 6;
  const SEGMENT_TIMEOUT = 30000;

  /**
   * Find the videomanifest URL from performance entries or captured data.
   */
  const findManifestUrl = () => {
    // Check performance entries first
    const entries = performance.getEntriesByType('resource');
    const manifestEntry = entries.find(e =>
      e.name.includes('videomanifest') || e.name.includes('manifest(format=')
    );
    if (manifestEntry) return manifestEntry.name;

    // Check captured URLs from videoDownloadOverride
    if (window.__teamsVideoCapture) {
      const tokens = window.__teamsVideoCapture.getTokens?.();
      if (tokens?.videoTemplate) {
        // The template URL can be converted back to a manifest URL
        try {
          const u = new URL(tokens.videoTemplate);
          // Remove segment-specific params to get manifest
          u.searchParams.delete('segmentTime');
          u.searchParams.delete('part');
          u.searchParams.set('part', 'manifest');
          return u.toString();
        } catch (e) {}
      }
    }

    return null;
  };

  /**
   * Find video/audio segment URL templates from captured data.
   */
  const findSegmentTemplates = () => {
    if (!window.__teamsVideoCapture) return null;
    const tokens = window.__teamsVideoCapture.getTokens?.();
    if (!tokens?.videoTemplate) return null;
    return {
      videoTemplate: tokens.videoTemplate,
      audioTemplate: tokens.audioTemplate
    };
  };

  /**
   * Parse the manifest to get segment information.
   */
  const parseManifest = async (manifestUrl) => {
    try {
      const resp = await fetch(manifestUrl);
      if (!resp.ok) return null;

      const contentType = resp.headers.get('content-type') || '';
      const text = await resp.text();

      // Try parsing as JSON (DASH-like manifest)
      if (contentType.includes('json') || text.startsWith('{')) {
        try {
          const data = JSON.parse(text);
          return parseJsonManifest(data);
        } catch (e) {}
      }

      // Try parsing as XML (MPD format)
      if (text.includes('<MPD') || text.includes('<SmoothStreamingMedia')) {
        return parseXmlManifest(text);
      }

      return null;
    } catch (e) {
      return null;
    }
  };

  const parseJsonManifest = (data) => {
    const result = { duration: 0, videoSegments: [], audioSegments: [] };

    // SharePoint/Stream manifest format
    if (data.Duration) result.duration = data.Duration / 10000000; // 100ns ticks to seconds

    const streams = data.Streams || data.streams || [];
    for (const stream of streams) {
      const type = (stream.Type || stream.type || '').toLowerCase();
      const chunks = stream.Chunks || stream.chunks || [];
      const segments = chunks.map((chunk, i) => ({
        index: i,
        time: chunk.Offset || chunk.offset || chunk.StartTime || 0,
        duration: chunk.Duration || chunk.duration || 2000
      }));

      if (type === 'video' || type === '0') {
        result.videoSegments = segments;
      } else if (type === 'audio' || type === '1') {
        result.audioSegments = segments;
      }
    }

    return result;
  };

  const parseXmlManifest = (xmlText) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, 'text/xml');
    const result = { duration: 0, videoSegments: [], audioSegments: [] };

    // Smooth Streaming format
    const root = doc.querySelector('SmoothStreamingMedia');
    if (root) {
      result.duration = parseInt(root.getAttribute('Duration') || '0') / 10000000;
    }

    const streamIndexes = doc.querySelectorAll('StreamIndex');
    streamIndexes.forEach(si => {
      const type = (si.getAttribute('Type') || '').toLowerCase();
      const chunks = si.querySelectorAll('c');
      let currentTime = 0;
      const segments = [];

      chunks.forEach((c, i) => {
        const t = parseInt(c.getAttribute('t') || currentTime);
        const d = parseInt(c.getAttribute('d') || '20000000');
        segments.push({ index: i, time: t, duration: d });
        currentTime = t + d;
      });

      if (type === 'video') result.videoSegments = segments;
      else if (type === 'audio') result.audioSegments = segments;
    });

    return result;
  };

  /**
   * Download a single segment with retry.
   */
  const fetchSegment = async (url, retries = 2) => {
    for (let attempt = 0; attempt <= retries; attempt++) {
      try {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), SEGMENT_TIMEOUT);
        const resp = await fetch(url, { signal: controller.signal });
        clearTimeout(timeout);
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        return await resp.arrayBuffer();
      } catch (e) {
        if (attempt === retries) throw e;
        await new Promise(r => setTimeout(r, 1000 * (attempt + 1)));
      }
    }
  };

  /**
   * Download all segments in parallel with concurrency limit.
   */
  const downloadSegments = async (template, segments, onProgress) => {
    const results = new Array(segments.length);
    let completed = 0;

    const downloadOne = async (seg) => {
      const url = template.replace('{TIME}', seg.time).replace('%7BTIME%7D', seg.time);
      results[seg.index] = await fetchSegment(url);
      completed++;
      if (onProgress) {
        onProgress({ completed, total: segments.length, percent: Math.round(completed / segments.length * 100) });
      }
    };

    // Process with concurrency limit
    const queue = [...segments];
    const workers = [];
    for (let i = 0; i < MAX_CONCURRENT; i++) {
      workers.push((async () => {
        while (queue.length > 0) {
          const seg = queue.shift();
          if (seg) await downloadOne(seg);
        }
      })());
    }
    await Promise.all(workers);

    return results.filter(Boolean);
  };

  /**
   * Check if this method is available.
   */
  const isAvailable = () => {
    const templates = findSegmentTemplates();
    if (!templates) return false;
    // Quick check: if URL contains enableEncryption=1, segments are likely encrypted
    // (full check happens at download time via areSegmentsEncrypted)
    if (templates.videoTemplate?.includes('enableEncryption=1')) return false;
    return true;
  };

  /**
   * Test if segments are encrypted by fetching the first one and checking for MP4 box headers.
   */
  const areSegmentsEncrypted = async (template) => {
    try {
      const url = template.replace('{TIME}', '0').replace('%7BTIME%7D', '0');
      const resp = await fetch(url);
      if (!resp.ok) return true; // assume encrypted if fetch fails
      const buf = await resp.arrayBuffer();
      const bytes = new Uint8Array(buf.slice(0, 8));
      const boxType = String.fromCharCode(...bytes.slice(4, 8));
      // Valid fMP4 segments start with styp, moof, ftyp, or similar
      return !['ftyp', 'moof', 'styp', 'mdat', 'free', 'sidx'].includes(boxType);
    } catch (e) {
      return true;
    }
  };

  /**
   * Attempt manifest-based download.
   * @param {Function} onProgress - progress callback({stage, message, percent})
   * @returns {Promise<{success: boolean, error?: string}>}
   */
  const download = async (onProgress) => {
    const templates = findSegmentTemplates();
    if (!templates) {
      return { success: false, error: 'No segment templates found. Play the video briefly first.' };
    }

    // Check if segments are encrypted (DRM)
    if (onProgress) onProgress({ stage: 'checking', message: 'Checking for encryption...' });
    const encrypted = await areSegmentsEncrypted(templates.videoTemplate);
    if (encrypted) {
      return { success: false, error: 'Segments are DRM-encrypted. Use "Record Stream" method instead.' };
    }

    // Try to get segment times from manifest
    const manifestUrl = findManifestUrl();
    let manifestData = null;
    if (manifestUrl) {
      if (onProgress) onProgress({ stage: 'manifest', message: 'Fetching video manifest...' });
      manifestData = await parseManifest(manifestUrl);
    }

    // If we have manifest data, use its segment times
    // Otherwise, generate segment times based on video duration
    const video = document.querySelector('video');
    const duration = video?.duration || 5000;
    const segmentDuration = 2; // 2 seconds per segment (typical)

    let videoSegments, audioSegments;
    if (manifestData && manifestData.videoSegments.length > 0) {
      videoSegments = manifestData.videoSegments;
      audioSegments = manifestData.audioSegments;
    } else {
      // Generate segment list from duration
      const segCount = Math.ceil(duration / segmentDuration);
      videoSegments = Array.from({ length: segCount }, (_, i) => ({
        index: i,
        time: i * segmentDuration * 1000, // milliseconds
        duration: segmentDuration * 1000
      }));
      audioSegments = videoSegments.map(s => ({ ...s }));
    }

    if (onProgress) onProgress({
      stage: 'downloading',
      message: `Downloading ${videoSegments.length} video segments...`,
      percent: 0
    });

    try {
      // Download video segments
      const videoBuffers = await downloadSegments(
        templates.videoTemplate,
        videoSegments,
        (p) => {
          if (onProgress) onProgress({
            stage: 'video',
            message: `Video: ${p.completed}/${p.total} segments`,
            percent: Math.round(p.percent * 0.6) // 0-60%
          });
        }
      );

      // Download audio segments
      if (onProgress) onProgress({ stage: 'audio', message: 'Downloading audio segments...', percent: 60 });
      const audioBuffers = await downloadSegments(
        templates.audioTemplate,
        audioSegments,
        (p) => {
          if (onProgress) onProgress({
            stage: 'audio',
            message: `Audio: ${p.completed}/${p.total} segments`,
            percent: 60 + Math.round(p.percent * 0.3) // 60-90%
          });
        }
      );

      if (onProgress) onProgress({ stage: 'combining', message: 'Combining video and audio...', percent: 90 });

      // Combine all buffers into a single blob
      const videoBlob = new Blob(videoBuffers, { type: 'video/mp4' });
      const audioBlob = new Blob(audioBuffers, { type: 'audio/mp4' });

      // Try to mux with MP4Muxer if available, otherwise download separately
      if (window.MP4Muxer) {
        try {
          const combined = window.MP4Muxer.mux(
            null, videoBuffers, null, audioBuffers
          );
          const blob = new Blob([combined], { type: 'video/mp4' });
          triggerDownload(blob, getFileName('mp4'));
          if (onProgress) onProgress({ stage: 'done', message: 'Download complete!', percent: 100 });
          return { success: true };
        } catch (e) {
          console.warn('[manifestDownload] Muxing failed, downloading separately:', e);
        }
      }

      // Fallback: download video+audio separately
      triggerDownload(videoBlob, getFileName('video.mp4'));
      triggerDownload(audioBlob, getFileName('audio.mp4'));
      if (onProgress) onProgress({ stage: 'done', message: 'Downloaded video and audio separately', percent: 100 });
      return { success: true };

    } catch (err) {
      return { success: false, error: `Download failed: ${err.message}` };
    }
  };

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
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  };

  // Expose module
  window.__videoDownloadModules = window.__videoDownloadModules || {};
  window.__videoDownloadModules[MODULE_NAME] = {
    name: MODULE_NAME,
    label: 'Manifest Download',
    description: 'Fast parallel segment download (play video briefly first to capture tokens)',
    isAvailable,
    download,
    findManifestUrl,
    findSegmentTemplates
  };

  console.log('[Teams Chat Exporter] Video download module loaded: manifestDownload');
})();
