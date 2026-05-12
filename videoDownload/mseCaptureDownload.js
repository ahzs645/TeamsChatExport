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
    // Available whenever there's a media element with a known duration —
    // we can prebuffer it ourselves at 16x to populate the MSE buffers.
    const el = document.querySelector('video, audio');
    return !!(el && isFinite(el.duration) && el.duration > 0);
  };

  const formatTime = (s) => {
    const m = Math.floor(s / 60);
    const sec = Math.floor(s % 60);
    return `${m}:${String(sec).padStart(2, '0')}`;
  };

  /**
   * Drive the media element from 0 → duration at 16x so the page's DASH
   * player demands every segment. Our SourceBuffer.appendBuffer hook then
   * captures each one as it's fed to MSE.
   *
   * Clears existing MSE buffers first so the resulting set is exactly one
   * clean pass — otherwise re-played segments would duplicate.
   */
  const prebufferAt16x = async (onProgress) => {
    const video = document.querySelector('video, audio');
    if (!video) return { ok: false, error: 'No media element found' };
    const duration = video.duration;
    if (!isFinite(duration) || duration <= 0) {
      return { ok: false, error: 'Duration unknown — start playback once, then retry' };
    }

    const cap = getCapture();
    cap?.clear?.();

    const orig = {
      currentTime: video.currentTime,
      playbackRate: video.playbackRate,
      muted: video.muted,
      paused: video.paused,
    };

    video.muted = true;
    try { video.currentTime = 0; } catch (_) {}
    video.playbackRate = 16;

    try {
      await video.play();
    } catch (e) {
      video.playbackRate = orig.playbackRate;
      video.muted = orig.muted;
      return { ok: false, error: 'Autoplay blocked — click the play button on the page once, then retry' };
    }

    await new Promise((resolve) => {
      let lastTime = -1;
      let stuckTicks = 0;
      const tick = setInterval(() => {
        const t = video.currentTime;
        if (onProgress) onProgress({
          stage: 'prebuffering',
          message: `Fast-play 16x: ${formatTime(t)} / ${formatTime(duration)}`,
          percent: Math.min(60, Math.round((t / duration) * 60)),
        });

        if (video.ended || t >= duration - 0.25) {
          clearInterval(tick);
          resolve();
          return;
        }
        if (Math.abs(t - lastTime) < 0.05) {
          stuckTicks++;
          if (stuckTicks > 24) { // ~6s of no progress → bail
            clearInterval(tick);
            resolve();
            return;
          }
        } else {
          stuckTicks = 0;
        }
        lastTime = t;
      }, 250);
    });

    // Brief wait so the trailing appendBuffer for the tail segment lands.
    await new Promise((r) => setTimeout(r, 600));

    video.pause();
    video.playbackRate = orig.playbackRate;
    video.muted = orig.muted;
    try { video.currentTime = orig.currentTime; } catch (_) {}
    if (!orig.paused) video.play().catch(() => {});

    return { ok: true };
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

    // Step 1: drive playback at 16x so MSE receives every segment.
    if (onProgress) onProgress({ stage: 'prebuffering', message: 'Starting 16x prebuffer...', percent: 0 });
    const pre = await prebufferAt16x(onProgress);
    if (!pre.ok) return { success: false, error: pre.error };

    if (onProgress) onProgress({ stage: 'assembling', message: 'Assembling captured segments...', percent: 70 });

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
