/**
 * Video Download Override
 * Captures and downloads video/audio segments in real-time as they're requested.
 * Uses high-speed playback to trigger segment requests, then downloads immediately
 * with valid auth tokens (similar to how HLS/DASH downloaders work).
 */
(() => {
  const { fetch: originalFetch } = window;

  // Storage for captured and downloaded segments
  const videoSegments = new Map(); // time -> ArrayBuffer
  const audioSegments = new Map(); // time -> ArrayBuffer
  let isCapturing = false;
  let downloadAsCapture = false; // Download segments as they're captured

  // Stats for UI updates
  let stats = {
    videoCount: 0,
    audioCount: 0,
    videoPending: 0,
    audioPending: 0,
    maxTime: 0,
    errors: 0
  };

  const dispatchUpdate = () => {
    window.dispatchEvent(new CustomEvent('teamsVideoSegmentUpdate', {
      detail: { ...stats }
    }));
  };

  // Download a segment immediately when captured
  const downloadSegment = async (url, time, isAudio) => {
    try {
      const response = await originalFetch(url);
      if (!response.ok) {
        stats.errors++;
        console.warn(`[Teams Video] Segment ${time}ms failed: HTTP ${response.status}`);
        return null;
      }
      const buffer = await response.arrayBuffer();

      if (isAudio) {
        audioSegments.set(time, buffer);
        stats.audioCount = audioSegments.size;
      } else {
        videoSegments.set(time, buffer);
        stats.videoCount = videoSegments.size;
        stats.maxTime = Math.max(stats.maxTime, time);
      }

      dispatchUpdate();
      return buffer;
    } catch (err) {
      stats.errors++;
      console.warn(`[Teams Video] Segment ${time}ms error:`, err.message);
      return null;
    }
  };

  // Intercept fetch to capture AND download segment URLs
  window.fetch = async (...args) => {
    const [resource, config] = args;
    const url = typeof resource === 'string' ? resource : resource?.url || '';

    // Check if this is a video/audio segment request
    if (isCapturing && url.includes('videomanifest') && url.includes('segmentTime=')) {
      const timeMatch = url.match(/segmentTime=(\d+)/);
      const isAudio = url.includes('aformat') || url.includes('/audio');
      const isVideo = !isAudio && (url.includes('vformat') || url.includes('vcopy') || url.includes('/video') || timeMatch);

      if (timeMatch) {
        const time = parseInt(timeMatch[1]);

        // Download immediately if enabled and not already captured
        if (downloadAsCapture) {
          if (isVideo && !videoSegments.has(time)) {
            stats.videoPending++;
            dispatchUpdate();
            // Don't await - download in background
            downloadSegment(url, time, false).then(() => {
              stats.videoPending--;
              dispatchUpdate();
            });
          } else if (isAudio && !audioSegments.has(time)) {
            stats.audioPending++;
            dispatchUpdate();
            downloadSegment(url, time, true).then(() => {
              stats.audioPending--;
              dispatchUpdate();
            });
          }
        }
      }
    }

    return originalFetch(...args);
  };

  // API exposed to window for content script access
  window.__teamsVideoCapture = {
    // Start capturing and downloading segments
    startCapture: (downloadImmediately = true) => {
      isCapturing = true;
      downloadAsCapture = downloadImmediately;
      videoSegments.clear();
      audioSegments.clear();
      stats = { videoCount: 0, audioCount: 0, videoPending: 0, audioPending: 0, maxTime: 0, errors: 0 };
      console.log('[Teams Video] Capture started (download immediately:', downloadImmediately, ')');
      dispatchUpdate();
      return true;
    },

    // Stop capturing
    stopCapture: () => {
      isCapturing = false;
      console.log('[Teams Video] Capture stopped. Video:', videoSegments.size, 'Audio:', audioSegments.size);
      return {
        videoCount: videoSegments.size,
        audioCount: audioSegments.size
      };
    },

    // Check if capturing
    isCapturing: () => isCapturing,

    // Get current stats
    getStats: () => ({
      ...stats,
      isCapturing,
      videoCount: videoSegments.size,
      audioCount: audioSegments.size
    }),

    // Wait for all pending downloads to complete
    waitForPending: async (timeoutMs = 30000) => {
      const start = Date.now();
      while ((stats.videoPending > 0 || stats.audioPending > 0) && (Date.now() - start < timeoutMs)) {
        await new Promise(r => setTimeout(r, 100));
      }
      return stats.videoPending === 0 && stats.audioPending === 0;
    },

    // High-speed capture - seeks through video rapidly to trigger all segment downloads
    highSpeedCapture: async (video, onProgress) => {
      if (!video) {
        video = document.querySelector('video');
      }
      if (!video) {
        throw new Error('No video element found');
      }

      const duration = video.duration;
      if (!duration || !isFinite(duration)) {
        throw new Error('Video duration not available');
      }

      // Start capturing
      window.__teamsVideoCapture.startCapture(true);

      // Save video state
      const wasPlaying = !video.paused;
      const wasTime = video.currentTime;
      const wasMuted = video.muted;
      const wasPlaybackRate = video.playbackRate;

      // Configure for fast capture
      video.muted = true;
      video.currentTime = 0;

      // Use maximum playback rate (usually 16x max, but try higher)
      const maxRate = 16;
      video.playbackRate = maxRate;

      if (onProgress) onProgress(0, duration, 'Starting high-speed capture...');

      // Start playing
      try {
        await video.play();
      } catch (e) {
        console.warn('[Teams Video] Autoplay blocked, trying with user gesture simulation');
        // If autoplay blocked, we need user interaction
        throw new Error('Please click the video play button first, then try again');
      }

      // Monitor progress
      return new Promise((resolve, reject) => {
        let lastTime = 0;
        let stuckCount = 0;

        const checkProgress = setInterval(() => {
          const currentTime = video.currentTime;
          const progress = currentTime / duration;

          if (onProgress) {
            const captured = videoSegments.size;
            const expected = Math.ceil(currentTime);
            onProgress(currentTime, duration, `${Math.round(progress * 100)}% - ${captured} segments captured`);
          }

          // Check if stuck
          if (Math.abs(currentTime - lastTime) < 0.1) {
            stuckCount++;
            if (stuckCount > 30) { // 3 seconds stuck
              // Try to unstick by seeking forward
              video.currentTime = Math.min(currentTime + 2, duration - 1);
              stuckCount = 0;
            }
          } else {
            stuckCount = 0;
          }
          lastTime = currentTime;

          // Check if done
          if (currentTime >= duration - 1 || video.ended) {
            clearInterval(checkProgress);
            video.pause();

            // Wait for pending downloads
            window.__teamsVideoCapture.waitForPending(10000).then(() => {
              // Restore video state
              video.muted = wasMuted;
              video.playbackRate = wasPlaybackRate;
              video.currentTime = wasTime;

              window.__teamsVideoCapture.stopCapture();

              if (onProgress) {
                onProgress(duration, duration, `Complete! ${videoSegments.size} video + ${audioSegments.size} audio segments`);
              }

              resolve({
                videoCount: videoSegments.size,
                audioCount: audioSegments.size,
                duration
              });
            });
          }
        }, 100);

        // Timeout after 10 minutes
        setTimeout(() => {
          clearInterval(checkProgress);
          video.pause();
          video.muted = wasMuted;
          video.playbackRate = wasPlaybackRate;
          reject(new Error('Capture timed out'));
        }, 600000);
      });
    },

    // Get captured segments as blobs
    getBlobs: () => {
      // Sort segments by time
      const sortedVideo = Array.from(videoSegments.entries())
        .sort((a, b) => a[0] - b[0])
        .map(([_, buffer]) => buffer);

      const sortedAudio = Array.from(audioSegments.entries())
        .sort((a, b) => a[0] - b[0])
        .map(([_, buffer]) => buffer);

      return {
        videoBlob: sortedVideo.length > 0 ? new Blob(sortedVideo, { type: 'video/mp4' }) : null,
        audioBlob: sortedAudio.length > 0 ? new Blob(sortedAudio, { type: 'audio/mp4' }) : null,
        videoCount: sortedVideo.length,
        audioCount: sortedAudio.length
      };
    },

    // Clear captured segments
    clear: () => {
      videoSegments.clear();
      audioSegments.clear();
      stats = { videoCount: 0, audioCount: 0, videoPending: 0, audioPending: 0, maxTime: 0, errors: 0 };
      dispatchUpdate();
    }
  };

  console.log('[Teams Chat Exporter] Video download override installed (high-speed capture mode)');
})();
