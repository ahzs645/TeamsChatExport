/**
 * Video Download Override
 * Captures and downloads video/audio segments in real-time as they're requested.
 * Uses high-speed playback to trigger segment requests, then downloads immediately
 * with valid auth tokens (similar to how HLS/DASH downloaders work).
 *
 * Supports multiple URL patterns:
 * - SharePoint/Stream: mediap.svc.ms with part=mediasegment
 * - Teams recordings: videomanifest with segmentTime=
 */
(() => {
  const { fetch: originalFetch } = window;

  // Storage for captured and downloaded segments
  const videoSegments = new Map(); // index -> ArrayBuffer
  const audioSegments = new Map(); // index -> ArrayBuffer
  let isCapturing = false;
  let downloadAsCapture = false;

  // Captured auth tokens and URL templates
  let capturedVideoToken = null;
  let capturedAudioToken = null;
  let capturedVideoUrlTemplate = null;
  let capturedAudioUrlTemplate = null;
  let segmentIndex = 0;

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

  // Check if URL is a media segment request
  const isMediaSegmentUrl = (url) => {
    const lower = url.toLowerCase();
    return (
      // SharePoint/Stream pattern
      (lower.includes('mediap.svc.ms') && lower.includes('mediasegment')) ||
      (lower.includes('mediaservice') && lower.includes('segment')) ||
      // Teams recordings pattern
      (lower.includes('videomanifest') && lower.includes('segmenttime=')) ||
      // Generic patterns
      (lower.includes('/segment') && (lower.includes('.mp4') || lower.includes('video') || lower.includes('audio'))) ||
      (lower.includes('part=media'))
    );
  };

  // Parse segment info from URL
  const parseSegmentInfo = (url) => {
    const lower = url.toLowerCase();

    // Determine if audio or video
    const isAudio = lower.includes('audio') || lower.includes('aformat') ||
                    (lower.includes('type=audio') || lower.includes('/audio/'));

    // Use sequential index for consistent ordering (SharePoint URLs don't have reliable time info)
    // Segments are typically ~2-4 seconds each
    const currentIndex = segmentIndex;
    const estimatedTimeMs = currentIndex * 2000; // Assume ~2 second segments

    return { isAudio, time: currentIndex, estimatedTimeMs };
  };

  // Clone response and save segment data
  const captureSegment = async (response, url, isAudio, index, estimatedTimeMs) => {
    try {
      const clone = response.clone();
      const buffer = await clone.arrayBuffer();

      if (buffer.byteLength < 100) {
        // Too small, probably not a real segment
        return;
      }

      if (isAudio) {
        if (!audioSegments.has(index)) {
          audioSegments.set(index, buffer);
          stats.audioCount = audioSegments.size;
        }
      } else {
        if (!videoSegments.has(index)) {
          videoSegments.set(index, buffer);
          stats.videoCount = videoSegments.size;
          stats.maxTime = Math.max(stats.maxTime, estimatedTimeMs || 0);
        }
      }

      dispatchUpdate();
    } catch (err) {
      stats.errors++;
      console.warn(`[Teams Video] Segment capture error:`, err.message);
    }
  };

  // Intercept fetch to capture segment responses
  window.fetch = async (...args) => {
    const [resource, config] = args;
    const url = typeof resource === 'string' ? resource : resource?.url || '';

    // Capture auth tokens from outgoing requests (like chat API does)
    if (config?.headers) {
      const headers = config.headers;
      let authToken = null;

      if (typeof headers.get === 'function') {
        authToken = headers.get('Authorization') || headers.get('authorization');
      } else if (typeof headers === 'object') {
        authToken = headers.Authorization || headers.authorization;
      }

      if (authToken && url.includes('mediap.svc.ms')) {
        if (url.toLowerCase().includes('audio')) {
          capturedAudioToken = authToken;
        } else {
          capturedVideoToken = authToken;
        }
        console.log('[Teams Video] Captured auth token from media request');
      }
    }

    // Make the actual request
    const response = await originalFetch(...args);

    // Check if this is a media segment we should capture
    if (isCapturing && isMediaSegmentUrl(url)) {
      const { isAudio, time, estimatedTimeMs } = parseSegmentInfo(url);
      segmentIndex++;

      // Capture the response data
      if (downloadAsCapture) {
        // Capture in background (don't block)
        if (isAudio) {
          stats.audioPending++;
        } else {
          stats.videoPending++;
        }
        dispatchUpdate();

        captureSegment(response, url, isAudio, time, estimatedTimeMs).then(() => {
          if (isAudio) {
            stats.audioPending--;
          } else {
            stats.videoPending--;
          }
          dispatchUpdate();
        });
      }

      // Save URL template for potential direct downloads
      const template = url.replace(/segment[=/]\d+/i, 'segment={INDEX}')
                          .replace(/segmentTime=\d+/i, 'segmentTime={TIME}');
      if (isAudio && !capturedAudioUrlTemplate) {
        capturedAudioUrlTemplate = template;
      } else if (!isAudio && !capturedVideoUrlTemplate) {
        capturedVideoUrlTemplate = template;
      }
    }

    return response;
  };

  // Also intercept XHR (some media players use XHR instead of fetch)
  const originalXHROpen = XMLHttpRequest.prototype.open;
  const originalXHRSend = XMLHttpRequest.prototype.send;

  XMLHttpRequest.prototype.open = function(method, url, ...rest) {
    this.__teamsVideoUrl = url;
    return originalXHROpen.call(this, method, url, ...rest);
  };

  XMLHttpRequest.prototype.send = function(...args) {
    const url = this.__teamsVideoUrl || '';

    if (isCapturing && isMediaSegmentUrl(url)) {
      this.addEventListener('load', () => {
        if (this.response && this.response.byteLength > 100) {
          const { isAudio, time } = parseSegmentInfo(url);

          if (isAudio) {
            if (!audioSegments.has(time)) {
              audioSegments.set(time, this.response);
              stats.audioCount = audioSegments.size;
              console.log(`[Teams Video] Audio segment (XHR): ${time}ms`);
            }
          } else {
            if (!videoSegments.has(time)) {
              videoSegments.set(time, this.response);
              stats.videoCount = videoSegments.size;
              stats.maxTime = Math.max(stats.maxTime, time);
              console.log(`[Teams Video] Video segment (XHR): ${time}ms`);
            }
          }
          dispatchUpdate();
        }
      });
    }

    return originalXHRSend.apply(this, args);
  };

  // API exposed to window for content script access
  window.__teamsVideoCapture = {
    // Start capturing and downloading segments
    startCapture: (downloadImmediately = true) => {
      isCapturing = true;
      downloadAsCapture = downloadImmediately;
      segmentIndex = 0;
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
      audioCount: audioSegments.size,
      hasVideoToken: !!capturedVideoToken,
      hasAudioToken: !!capturedAudioToken,
      hasVideoTemplate: !!capturedVideoUrlTemplate,
      hasAudioTemplate: !!capturedAudioUrlTemplate
    }),

    // Get captured tokens and templates (for API-based download attempts)
    getTokens: () => ({
      videoToken: capturedVideoToken,
      audioToken: capturedAudioToken,
      videoTemplate: capturedVideoUrlTemplate,
      audioTemplate: capturedAudioUrlTemplate
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
      capturedVideoToken = null;
      capturedAudioToken = null;
      capturedVideoUrlTemplate = null;
      capturedAudioUrlTemplate = null;
      segmentIndex = 0;
      stats = { videoCount: 0, audioCount: 0, videoPending: 0, audioPending: 0, maxTime: 0, errors: 0 };
      dispatchUpdate();
    },

    // Debug: Log current state
    debug: () => {
      console.log('[Teams Video Debug]');
      console.log('  isCapturing:', isCapturing);
      console.log('  segmentIndex:', segmentIndex);
      console.log('  videoSegments:', videoSegments.size);
      console.log('  audioSegments:', audioSegments.size);
      console.log('  videoToken:', capturedVideoToken ? 'captured' : 'none');
      console.log('  audioToken:', capturedAudioToken ? 'captured' : 'none');
      console.log('  videoTemplate:', capturedVideoUrlTemplate?.substring(0, 100) || 'none');
      console.log('  audioTemplate:', capturedAudioUrlTemplate?.substring(0, 100) || 'none');
      console.log('  stats:', stats);
    }
  };

  // Listen for commands from content script via custom events
  document.addEventListener('teamsVideoCommand', async (e) => {
    const { command, data } = e.detail || {};
    let result = null;

    const notify = (msg) => {
      console.log('[Teams Video]', msg);
      document.dispatchEvent(new CustomEvent('teamsVideoDownloadProgress', { detail: { message: msg } }));
    };

    switch (command) {
      case 'startCapture':
        result = window.__teamsVideoCapture.startCapture(data?.downloadImmediately ?? true);
        break;
      case 'stopCapture':
        result = window.__teamsVideoCapture.stopCapture();
        break;
      case 'getStats':
        result = window.__teamsVideoCapture.getStats();
        break;
      case 'getBlobs':
        result = window.__teamsVideoCapture.getBlobs();
        break;
      case 'getTokens':
        result = window.__teamsVideoCapture.getTokens();
        break;
      case 'clear':
        window.__teamsVideoCapture.clear();
        result = { success: true };
        break;
      case 'debug':
        window.__teamsVideoCapture.debug();
        result = { success: true };
        break;
      case 'ping':
        result = { available: true, version: 'v5.8' };
        break;

      case 'downloadFiles':
        // Handle file download directly in page context (avoids CSP issues)
        try {
          notify('Step 1: Waiting for pending downloads...');
          await window.__teamsVideoCapture.waitForPending(10000);

          const stats = window.__teamsVideoCapture.getStats();
          notify(`Step 2: Found ${stats.videoCount} video, ${stats.audioCount} audio segments`);

          if (stats.videoCount === 0) {
            throw new Error('No video segments found');
          }

          notify(`Step 3: Creating blobs from ${stats.videoCount} segments...`);
          const blobs = window.__teamsVideoCapture.getBlobs();

          if (!blobs.videoBlob) {
            throw new Error('Failed to create video blob');
          }

          const videoSizeMB = (blobs.videoBlob.size / 1024 / 1024).toFixed(1);
          notify(`Step 4: Video blob ready (${videoSizeMB} MB)`);

          // Get video title
          const titleEl = document.querySelector('h1[class*="videoTitleViewModeHeading"] label');
          const title = titleEl?.innerText?.trim() || document.title?.trim() || 'video';
          const safeTitle = title.replace(/[^a-zA-Z0-9\s-]/g, '').trim() || 'recording';

          notify('Step 5: Triggering video download...');
          const videoUrl = URL.createObjectURL(blobs.videoBlob);
          const a = document.createElement('a');
          a.href = videoUrl;
          a.download = `${safeTitle}-video.mp4`;
          a.style.display = 'none';
          document.body.appendChild(a);
          a.click();
          await new Promise(r => setTimeout(r, 500));
          document.body.removeChild(a);
          URL.revokeObjectURL(videoUrl);

          // Download audio if available
          if (blobs.audioBlob && blobs.audioCount > 0) {
            const audioSizeMB = (blobs.audioBlob.size / 1024 / 1024).toFixed(1);
            notify(`Step 6: Downloading audio (${audioSizeMB} MB)...`);
            await new Promise(r => setTimeout(r, 1000));

            const audioUrl = URL.createObjectURL(blobs.audioBlob);
            const b = document.createElement('a');
            b.href = audioUrl;
            b.download = `${safeTitle}-audio.mp4`;
            b.style.display = 'none';
            document.body.appendChild(b);
            b.click();
            await new Promise(r => setTimeout(r, 500));
            document.body.removeChild(b);
            URL.revokeObjectURL(audioUrl);
          }

          notify('Step 7: Downloads complete!');
          result = {
            success: true,
            videoCount: blobs.videoCount,
            audioCount: blobs.audioCount,
            title: safeTitle
          };
        } catch (err) {
          console.error('[Teams Video] Download error:', err);
          result = { success: false, error: err.message };
        }
        break;
    }

    // Send result back via another custom event
    document.dispatchEvent(new CustomEvent('teamsVideoResponse', {
      detail: { command, result }
    }));
  });

  // Notify that the API is ready
  document.dispatchEvent(new CustomEvent('teamsVideoReady', { detail: { version: 'v5.9' } }));

  console.log('[Teams Chat Exporter] Video download override installed (high-speed capture mode v2)');
})();
