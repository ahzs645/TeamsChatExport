/**
 * CaptureStream Download Module
 * Records the video using captureStream() + MediaRecorder.
 * Captures the decoded, decrypted video+audio directly from the video element.
 * Works on view-only/DRM content where direct download is blocked.
 * Produces a playable webm or mp4 file.
 */
(() => {
  const MODULE_NAME = 'captureStreamDownload';
  const DEFAULT_VIDEO_BITRATE = 5000000;       // 5 Mbps
  const DEFAULT_AUDIO_BITRATE = 128000;        // 128 kbps (with video)
  const DEFAULT_AUDIO_ONLY_BITRATE = 320000;   // 320 kbps (audio-only, matches high-quality MP3)
  const MAX_PLAYBACK_RATE = 16;

  let activeRecorder = null;
  let activeChunks = [];
  let isRecording = false;
  let recordingStartTime = 0;

  /**
   * Get the best supported MIME type for recording.
   * Picks an audio-only container when the stream has no video tracks,
   * otherwise picks a video container.
   */
  const getBestMimeType = (audioOnly = false) => {
    const types = audioOnly
      ? [
          'audio/webm;codecs=opus',
          'audio/mp4;codecs=mp4a.40.2', // Chrome 130+
          'audio/webm',
          'audio/mp4'
        ]
      : [
          'video/webm;codecs=vp9,opus',
          'video/webm;codecs=vp8,opus',
          'video/webm;codecs=h264,opus',
          'video/webm',
          'video/mp4'
        ];
    return types.find(t => MediaRecorder.isTypeSupported(t))
      || (audioOnly ? 'audio/webm' : 'video/webm');
  };

  /**
   * Check if this method is available.
   */
  const isAvailable = () => {
    const video = document.querySelector('video');
    if (!video) return false;
    if (typeof MediaRecorder === 'undefined') return false;

    // Test if captureStream is available
    try {
      const stream = video.captureStream ? video.captureStream(0) : video.mozCaptureStream?.(0);
      if (stream) {
        stream.getTracks().forEach(t => t.stop());
        return true;
      }
    } catch (e) {}
    return false;
  };

  /**
   * Record the full video at high speed.
   * @param {Object} options - {playbackRate, videoBitrate, audioBitrate}
   * @param {Function} onProgress - progress callback({stage, message, percent, currentTime, duration})
   * @returns {Promise<{success: boolean, error?: string}>}
   */
  const download = async (onProgress, options = {}) => {
    if (isRecording) {
      return { success: false, error: 'Recording already in progress' };
    }

    const video = document.querySelector('video');
    if (!video) {
      return { success: false, error: 'No video element found' };
    }

    const videoBitrate = options.videoBitrate || DEFAULT_VIDEO_BITRATE;
    const duration = video.duration || 0;

    if (!duration || duration === Infinity) {
      return { success: false, error: 'Video duration unknown' };
    }

    return new Promise((resolve) => {
      try {
        // Get the stream from the video element
        const stream = video.captureStream ? video.captureStream() : video.mozCaptureStream();
        if (!stream || stream.getTracks().length === 0) {
          resolve({ success: false, error: 'Could not capture video stream' });
          return;
        }

        const hasVideoTrack = stream.getVideoTracks().length > 0;
        const audioOnly = !hasVideoTrack;
        const mimeType = getBestMimeType(audioOnly);
        const audioBitrate = options.audioBitrate
          || (audioOnly ? DEFAULT_AUDIO_ONLY_BITRATE : DEFAULT_AUDIO_BITRATE);
        // Audio is captured in real time, so high playback rates produce sped-up output.
        // Force 1x for audio-only recordings.
        const playbackRate = audioOnly ? 1 : (options.playbackRate || MAX_PLAYBACK_RATE);
        activeChunks = [];
        isRecording = true;
        recordingStartTime = Date.now();

        const recorderOptions = { mimeType, audioBitsPerSecond: audioBitrate };
        if (!audioOnly) recorderOptions.videoBitsPerSecond = videoBitrate;
        activeRecorder = new MediaRecorder(stream, recorderOptions);

        activeRecorder.ondataavailable = (e) => {
          if (e.data && e.data.size > 0) {
            activeChunks.push(e.data);
          }
        };

        activeRecorder.onstop = () => {
          isRecording = false;
          video.pause();
          video.playbackRate = 1;
          video.muted = false;

          if (activeChunks.length === 0) {
            resolve({ success: false, error: 'No data recorded' });
            return;
          }

          const blob = new Blob(activeChunks, { type: mimeType });
          const ext = mimeType.startsWith('audio/')
            ? (mimeType.includes('mp4') ? 'm4a' : 'webm')
            : (mimeType.includes('mp4') ? 'mp4' : 'webm');
          const filename = getFileName(ext);

          if (onProgress) onProgress({
            stage: 'saving',
            message: `Saving ${filename} (${Math.round(blob.size / 1024 / 1024)}MB)...`,
            percent: 100
          });

          triggerDownload(blob, filename);
          activeChunks = [];
          activeRecorder = null;

          resolve({ success: true, fileName: filename, fileSize: blob.size });
        };

        activeRecorder.onerror = (e) => {
          isRecording = false;
          video.pause();
          video.playbackRate = 1;
          activeRecorder = null;
          resolve({ success: false, error: `Recording error: ${e.error?.message || 'unknown'}` });
        };

        // Start recording with data chunks every second
        activeRecorder.start(1000);

        // Configure video for high-speed playback
        video.muted = true;
        video.currentTime = 0;
        video.playbackRate = playbackRate;

        if (onProgress) onProgress({
          stage: 'recording',
          message: `Recording at ${playbackRate}x speed...`,
          percent: 0,
          currentTime: 0,
          duration
        });

        video.play().then(() => {
          // Monitor progress
          const progressInterval = setInterval(() => {
            if (!isRecording) {
              clearInterval(progressInterval);
              return;
            }

            const currentTime = video.currentTime;
            const percent = Math.min(99, Math.round((currentTime / duration) * 100));
            const elapsed = (Date.now() - recordingStartTime) / 1000;
            const eta = elapsed > 0 ? Math.round((duration - currentTime) / playbackRate) : 0;

            if (onProgress) onProgress({
              stage: 'recording',
              message: `Recording: ${formatTime(currentTime)} / ${formatTime(duration)} (~${eta}s remaining)`,
              percent,
              currentTime,
              duration
            });

            // Check if video ended
            if (video.ended || currentTime >= duration - 1) {
              clearInterval(progressInterval);
              setTimeout(() => {
                if (activeRecorder && activeRecorder.state === 'recording') {
                  activeRecorder.stop();
                }
              }, 500); // Brief delay to capture final frames
            }
          }, 500);
        }).catch((err) => {
          activeRecorder.stop();
          isRecording = false;
          resolve({ success: false, error: `Playback failed: ${err.message}. Keep this tab in the foreground.` });
        });

      } catch (err) {
        isRecording = false;
        resolve({ success: false, error: err.message });
      }
    });
  };

  /**
   * Stop an in-progress recording.
   */
  const stop = () => {
    if (activeRecorder && activeRecorder.state === 'recording') {
      activeRecorder.stop();
    }
    const video = document.querySelector('video');
    if (video) {
      video.pause();
      video.playbackRate = 1;
      video.muted = false;
    }
    isRecording = false;
  };

  /**
   * Check if currently recording.
   */
  const getStatus = () => ({
    isRecording,
    elapsedMs: isRecording ? Date.now() - recordingStartTime : 0,
    chunkCount: activeChunks.length,
    dataSizeMB: Math.round(activeChunks.reduce((sum, c) => sum + c.size, 0) / 1024 / 1024)
  });

  const formatTime = (seconds) => {
    const h = Math.floor(seconds / 3600);
    const m = Math.floor((seconds % 3600) / 60);
    const s = Math.floor(seconds % 60);
    return h > 0
      ? `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`
      : `${m}:${String(s).padStart(2, '0')}`;
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
    label: 'Record Stream',
    description: 'Record decoded video+audio at high speed (works on view-only content, keep tab in foreground)',
    isAvailable,
    download,
    stop,
    getStatus
  };

  console.log('[Teams Chat Exporter] Video download module loaded: captureStreamDownload');
})();
