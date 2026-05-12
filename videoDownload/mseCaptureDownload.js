/**
 * MSE Capture Download Module
 * Saves the *exact* decrypted bytes the browser fed to MediaSource — captured
 * passively via the SourceBuffer.appendBuffer hook in videoDownloadOverride.js.
 *
 * No re-fetch, no transcode, no playback required. Works for SharePoint Stream
 * MP3/audio (which is wrapped in MSE+CDN-encryption) because we read AFTER the
 * page has already decrypted the segments.
 *
 * Falls back to MP4Muxer when both video and audio tracks are present.
 */
(() => {
  const MODULE_NAME = 'mseCaptureDownload';

  const getCapture = () => window.__teamsVideoCapture;

  const getStatus = () => {
    const cap = getCapture();
    if (!cap?.getStats) return null;
    const s = cap.getStats();
    return {
      hasVideo: !!s.hasVideoInit && s.mseVideoCount > 0,
      hasAudio: !!s.hasAudioInit && s.mseAudioCount > 0,
      videoSegments: s.mseVideoCount,
      audioSegments: s.mseAudioCount,
      usage: cap.getMseUsage ? cap.getMseUsage() : null,
    };
  };

  const isAvailable = () => {
    const st = getStatus();
    return !!(st && (st.hasVideo || st.hasAudio));
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

  const download = async (onProgress) => {
    const cap = getCapture();
    if (!cap?.getBlobs) {
      return { success: false, error: 'Capture API not available' };
    }

    if (onProgress) onProgress({ stage: 'assembling', message: 'Assembling captured segments...', percent: 30 });

    const { videoBlob, audioBlob, videoCount, audioCount, source, hasMseInit } = cap.getBlobs();

    if (!videoBlob && !audioBlob) {
      return { success: false, error: 'No captured segments — play the media first to populate the MSE buffer' };
    }

    if (!hasMseInit) {
      // Came from the fetch path (likely encrypted). Bail out so the user falls
      // through to manifestDownload/captureStream rather than getting a broken file.
      return { success: false, error: 'No MSE init segment captured (bytes may be encrypted); try playing the media again' };
    }

    let outBlob;
    let ext;

    if (videoBlob && audioBlob && window.MP4Muxer?.mux) {
      if (onProgress) onProgress({ stage: 'muxing', message: 'Muxing video+audio...', percent: 70 });
      // Re-extract the buffers MP4Muxer needs. getBlobs() returned Blobs of
      // [init, ...segments] concatenated already, which MP4Muxer doesn't accept.
      // Mux directly from the override's internal arrays via a small helper.
      if (typeof cap.getMseRaw === 'function') {
        const raw = cap.getMseRaw();
        const muxed = window.MP4Muxer.mux(raw.videoInit, raw.videoSegments, raw.audioInit, raw.audioSegments);
        outBlob = new Blob([muxed], { type: 'video/mp4' });
      } else {
        // Fallback: ship video-only, since muxing concatenated blobs would be wrong.
        outBlob = videoBlob;
      }
      ext = 'mp4';
    } else if (videoBlob) {
      outBlob = videoBlob;
      ext = 'mp4';
    } else {
      outBlob = audioBlob;
      // Audio-only fragmented MP4 from Stream is AAC-in-MP4 — .m4a is the right extension.
      ext = 'm4a';
    }

    const filename = getFileName(ext);
    if (onProgress) onProgress({
      stage: 'saving',
      message: `Saving ${filename} (${Math.round(outBlob.size / 1024 / 1024)} MB, source: ${source})...`,
      percent: 100,
    });
    triggerDownload(outBlob, filename);

    return {
      success: true,
      fileName: filename,
      fileSize: outBlob.size,
      videoSegments: videoCount,
      audioSegments: audioCount,
      source,
    };
  };

  window.__videoDownloadModules = window.__videoDownloadModules || {};
  window.__videoDownloadModules[MODULE_NAME] = {
    name: MODULE_NAME,
    label: 'Save MSE Capture',
    description: 'Save the exact decrypted bytes the browser played (no re-fetch, no transcode)',
    isAvailable,
    download,
    getStatus,
  };

  console.log('[Teams Chat Exporter] Video download module loaded: mseCaptureDownload');
})();
