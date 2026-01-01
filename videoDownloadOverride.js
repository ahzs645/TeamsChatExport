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

  // Store actual captured URLs for analysis
  let capturedUrls = [];
  let manifestUrl = null;

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

    // Capture manifest URL (contains segment list)
    if (url.includes('videomanifest') || url.includes('manifest') || url.includes('.mpd') || url.includes('.m3u8')) {
      manifestUrl = url;
      console.log('[Teams Video] Captured manifest URL:', url.substring(0, 100));
    }

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
    if (isMediaSegmentUrl(url)) {
      // Always store URLs for analysis (even if not capturing)
      if (capturedUrls.length < 20) {
        capturedUrls.push({
          url: url,
          time: Date.now(),
          index: segmentIndex
        });
      }
    }

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

  // ============================================
  // MSE/SourceBuffer Interception (for decrypted data)
  // This captures data AFTER DRM decryption
  // ============================================

  const mseVideoBuffers = []; // Array of ArrayBuffers for video
  const mseAudioBuffers = []; // Array of ArrayBuffers for audio
  let mseVideoInit = null;    // Initialization segment for video
  let mseAudioInit = null;    // Initialization segment for audio
  let mseSegmentIndex = 0;

  // Detect if buffer is an initialization segment (contains moov/ftyp atoms)
  const isInitSegment = (buffer) => {
    if (buffer.byteLength < 8) return false;
    const view = new DataView(buffer);
    // Check for ftyp or moov box at start
    const type = String.fromCharCode(
      view.getUint8(4), view.getUint8(5), view.getUint8(6), view.getUint8(7)
    );
    return type === 'ftyp' || type === 'moov' || type === 'styp';
  };

  // Detect if buffer contains video or audio based on codec info
  const detectTrackType = (mimeType) => {
    if (!mimeType) return 'unknown';
    const lower = mimeType.toLowerCase();
    if (lower.includes('video')) return 'video';
    if (lower.includes('audio')) return 'audio';
    return 'unknown';
  };

  // Hook into SourceBuffer.appendBuffer
  const originalAppendBuffer = SourceBuffer.prototype.appendBuffer;
  SourceBuffer.prototype.appendBuffer = function(data) {
    if (isCapturing && data) {
      try {
        // Convert to ArrayBuffer if needed
        let buffer;
        if (data instanceof ArrayBuffer) {
          buffer = data;
        } else if (data.buffer instanceof ArrayBuffer) {
          buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        } else {
          buffer = new Uint8Array(data).buffer;
        }

        // Determine track type from parent MediaSource
        const mimeType = this._mseTrackType || this.mimeType || '';
        const trackType = detectTrackType(mimeType);

        if (buffer.byteLength > 100) { // Ignore tiny buffers
          if (isInitSegment(buffer)) {
            // Store initialization segment
            if (trackType === 'audio') {
              if (!mseAudioInit) {
                mseAudioInit = buffer.slice(0);
                console.log(`[Teams Video MSE] Audio init segment: ${buffer.byteLength} bytes`);
              }
            } else {
              if (!mseVideoInit) {
                mseVideoInit = buffer.slice(0);
                console.log(`[Teams Video MSE] Video init segment: ${buffer.byteLength} bytes`);
              }
            }
          } else {
            // Store media segment
            if (trackType === 'audio') {
              mseAudioBuffers.push(buffer.slice(0));
              stats.audioCount = mseAudioBuffers.length;
            } else {
              mseVideoBuffers.push(buffer.slice(0));
              stats.videoCount = mseVideoBuffers.length;
              mseSegmentIndex++;
            }
            dispatchUpdate();
          }
        }
      } catch (err) {
        console.warn('[Teams Video MSE] Capture error:', err);
      }
    }

    return originalAppendBuffer.call(this, data);
  };

  // Hook into MediaSource.addSourceBuffer to track mime types
  const originalAddSourceBuffer = MediaSource.prototype.addSourceBuffer;
  MediaSource.prototype.addSourceBuffer = function(mimeType) {
    const sourceBuffer = originalAddSourceBuffer.call(this, mimeType);
    sourceBuffer._mseTrackType = mimeType;
    console.log(`[Teams Video MSE] SourceBuffer created: ${mimeType}`);
    return sourceBuffer;
  };

  // Build complete MP4 from init + segments
  const buildMseBlob = (initSegment, mediaSegments, mimeType) => {
    if (!initSegment || mediaSegments.length === 0) {
      return null;
    }

    // Combine init segment + all media segments
    const totalSize = initSegment.byteLength + mediaSegments.reduce((sum, b) => sum + b.byteLength, 0);
    const combined = new Uint8Array(totalSize);

    let offset = 0;
    combined.set(new Uint8Array(initSegment), offset);
    offset += initSegment.byteLength;

    for (const segment of mediaSegments) {
      combined.set(new Uint8Array(segment), offset);
      offset += segment.byteLength;
    }

    return new Blob([combined], { type: mimeType });
  };

  // ============================================
  // MP4 Muxer - Combines video and audio streams
  // ============================================

  const MP4Muxer = (() => {
    const readUint32 = (data, offset) => {
      return (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
    };

    const writeUint32 = (data, offset, value) => {
      data[offset] = (value >> 24) & 0xff;
      data[offset + 1] = (value >> 16) & 0xff;
      data[offset + 2] = (value >> 8) & 0xff;
      data[offset + 3] = value & 0xff;
    };

    const getBoxType = (data, offset) => {
      return String.fromCharCode(data[offset], data[offset + 1], data[offset + 2], data[offset + 3]);
    };

    const findBox = (data, type, start = 0, end = null) => {
      end = end || data.length;
      let offset = start;
      while (offset < end - 8) {
        const size = readUint32(data, offset);
        const boxType = getBoxType(data, offset + 4);
        if (size < 8) break;
        if (boxType === type) return { offset, size };
        offset += size;
      }
      return null;
    };

    // Find all boxes of a given type within a range
    const findAllBoxes = (data, type, start = 0, end = null) => {
      end = end || data.length;
      const boxes = [];
      let offset = start;
      while (offset < end - 8) {
        const size = readUint32(data, offset);
        const boxType = getBoxType(data, offset + 4);
        if (size < 8) break;
        if (boxType === type) boxes.push({ offset, size });
        offset += size;
      }
      return boxes;
    };

    const updateTfhdTrackId = (data, moofOffset, moofSize, newTrackId) => {
      const trafBox = findBox(data, 'traf', moofOffset + 8, moofOffset + moofSize);
      if (!trafBox) return false;
      const tfhdBox = findBox(data, 'tfhd', trafBox.offset + 8, trafBox.offset + trafBox.size);
      if (!tfhdBox) return false;
      writeUint32(data, tfhdBox.offset + 12, newTrackId);
      return true;
    };

    // Update track ID in tkhd box (inside trak)
    const updateTkhdTrackId = (data, trakOffset, trakSize, newTrackId) => {
      const tkhdBox = findBox(data, 'tkhd', trakOffset + 8, trakOffset + trakSize);
      if (!tkhdBox) return false;
      const version = data[tkhdBox.offset + 8];
      const trackIdOffset = tkhdBox.offset + (version === 1 ? 20 : 12);
      writeUint32(data, trackIdOffset, newTrackId);
      return true;
    };

    // Create a merged moov box containing both video and audio tracks
    const createMergedMoov = (videoData, videoMoov, audioData, audioMoov) => {
      // Find trak boxes in video moov
      const videoTraks = findAllBoxes(videoData, 'trak', videoMoov.offset + 8, videoMoov.offset + videoMoov.size);
      // Find trak boxes in audio moov
      const audioTraks = findAllBoxes(audioData, 'trak', audioMoov.offset + 8, audioMoov.offset + audioMoov.size);

      if (videoTraks.length === 0) {
        console.error('[MP4Muxer] No video trak found');
        return null;
      }

      // Find mvhd box in video moov (movie header - we'll use this)
      const mvhdBox = findBox(videoData, 'mvhd', videoMoov.offset + 8, videoMoov.offset + videoMoov.size);
      if (!mvhdBox) {
        console.error('[MP4Muxer] No mvhd found');
        return null;
      }

      // Find mvex box if present (for fragmented mp4)
      const videoMvex = findBox(videoData, 'mvex', videoMoov.offset + 8, videoMoov.offset + videoMoov.size);
      const audioMvex = audioMoov ? findBox(audioData, 'mvex', audioMoov.offset + 8, audioMoov.offset + audioMoov.size) : null;

      // Calculate new moov size
      let newMoovContentSize = mvhdBox.size; // mvhd
      for (const trak of videoTraks) newMoovContentSize += trak.size;

      // Add audio traks
      let audioTrakData = [];
      if (audioMoov && audioTraks.length > 0) {
        for (const trak of audioTraks) {
          // Copy audio trak and update track ID to 2
          const trakCopy = audioData.slice(trak.offset, trak.offset + trak.size);
          const trakArray = new Uint8Array(trakCopy);
          updateTkhdTrackId(trakArray, 0, trak.size, 2);
          audioTrakData.push(trakArray);
          newMoovContentSize += trak.size;
        }
      }

      // Add mvex boxes if present
      let mvexSize = 0;
      if (videoMvex) mvexSize += videoMvex.size;
      // For audio mvex, we need to copy trex boxes with updated track IDs
      let audioTrexData = null;
      if (audioMvex) {
        const audioTrex = findBox(audioData, 'trex', audioMvex.offset + 8, audioMvex.offset + audioMvex.size);
        if (audioTrex) {
          audioTrexData = new Uint8Array(audioData.slice(audioTrex.offset, audioTrex.offset + audioTrex.size));
          // Update track ID in trex (offset 12 from box start)
          writeUint32(audioTrexData, 12, 2);
          mvexSize += audioTrex.size;
        }
      }

      // Create new mvex if we have both video and audio
      let newMvexData = null;
      if (videoMvex && audioTrexData) {
        const videoTrex = findBox(videoData, 'trex', videoMvex.offset + 8, videoMvex.offset + videoMvex.size);
        if (videoTrex) {
          const newMvexSize = 8 + videoTrex.size + audioTrexData.length;
          newMvexData = new Uint8Array(newMvexSize);
          writeUint32(newMvexData, 0, newMvexSize);
          newMvexData[4] = 'm'.charCodeAt(0);
          newMvexData[5] = 'v'.charCodeAt(0);
          newMvexData[6] = 'e'.charCodeAt(0);
          newMvexData[7] = 'x'.charCodeAt(0);
          newMvexData.set(videoData.subarray(videoTrex.offset, videoTrex.offset + videoTrex.size), 8);
          newMvexData.set(audioTrexData, 8 + videoTrex.size);
          newMoovContentSize += newMvexSize;
        }
      } else if (videoMvex) {
        newMoovContentSize += videoMvex.size;
      }

      // Build new moov
      const newMoovSize = 8 + newMoovContentSize;
      const newMoov = new Uint8Array(newMoovSize);
      let offset = 0;

      // Write moov header
      writeUint32(newMoov, 0, newMoovSize);
      newMoov[4] = 'm'.charCodeAt(0);
      newMoov[5] = 'o'.charCodeAt(0);
      newMoov[6] = 'o'.charCodeAt(0);
      newMoov[7] = 'v'.charCodeAt(0);
      offset = 8;

      // Write mvhd
      newMoov.set(videoData.subarray(mvhdBox.offset, mvhdBox.offset + mvhdBox.size), offset);
      offset += mvhdBox.size;

      // Write video traks
      for (const trak of videoTraks) {
        newMoov.set(videoData.subarray(trak.offset, trak.offset + trak.size), offset);
        offset += trak.size;
      }

      // Write audio traks
      for (const trakArray of audioTrakData) {
        newMoov.set(trakArray, offset);
        offset += trakArray.length;
      }

      // Write mvex
      if (newMvexData) {
        newMoov.set(newMvexData, offset);
        offset += newMvexData.length;
      } else if (videoMvex) {
        newMoov.set(videoData.subarray(videoMvex.offset, videoMvex.offset + videoMvex.size), offset);
        offset += videoMvex.size;
      }

      console.log(`[MP4Muxer] Created merged moov: ${newMoovSize} bytes (${videoTraks.length} video + ${audioTrakData.length} audio tracks)`);
      return newMoov;
    };

    const mux = (videoInit, videoSegments, audioInit, audioSegments) => {
      if (!videoInit || videoSegments.length === 0) {
        console.error('[MP4Muxer] No video data');
        return null;
      }

      const videoData = new Uint8Array(videoInit);
      const hasAudio = audioInit && audioSegments.length > 0;
      const audioData = hasAudio ? new Uint8Array(audioInit) : null;

      const ftypBox = findBox(videoData, 'ftyp');
      if (!ftypBox) {
        console.error('[MP4Muxer] No ftyp box found');
        return null;
      }

      const videoMoov = findBox(videoData, 'moov');
      if (!videoMoov) {
        console.error('[MP4Muxer] No video moov box found');
        return null;
      }

      const audioMoov = hasAudio ? findBox(audioData, 'moov') : null;

      // Create merged moov with both tracks
      const mergedMoov = hasAudio && audioMoov
        ? createMergedMoov(videoData, videoMoov, audioData, audioMoov)
        : videoData.subarray(videoMoov.offset, videoMoov.offset + videoMoov.size);

      if (!mergedMoov) {
        console.error('[MP4Muxer] Failed to create merged moov');
        return null;
      }

      // Calculate output size
      let totalSize = ftypBox.size + mergedMoov.length;
      for (const seg of videoSegments) totalSize += seg.byteLength;
      if (hasAudio) {
        for (const seg of audioSegments) totalSize += seg.byteLength;
      }

      const output = new Uint8Array(totalSize);
      let writeOffset = 0;

      // Write ftyp
      output.set(videoData.subarray(ftypBox.offset, ftypBox.offset + ftypBox.size), writeOffset);
      writeOffset += ftypBox.size;

      // Write merged moov
      output.set(mergedMoov, writeOffset);
      writeOffset += mergedMoov.length;

      // Interleave segments
      const maxSegments = Math.max(videoSegments.length, hasAudio ? audioSegments.length : 0);
      for (let i = 0; i < maxSegments; i++) {
        if (i < videoSegments.length) {
          const seg = new Uint8Array(videoSegments[i]);
          output.set(seg, writeOffset);
          writeOffset += seg.byteLength;
        }
        if (hasAudio && i < audioSegments.length) {
          const segCopy = new Uint8Array(audioSegments[i]);
          const moofBox = findBox(segCopy, 'moof');
          if (moofBox) updateTfhdTrackId(segCopy, moofBox.offset, moofBox.size, 2);
          output.set(segCopy, writeOffset);
          writeOffset += segCopy.byteLength;
        }
      }

      console.log(`[MP4Muxer] Created ${(output.byteLength / 1024 / 1024).toFixed(1)} MB combined file`);
      return output.buffer;
    };

    return { mux, findBox };
  })();

  // API exposed to window for content script access
  window.__teamsVideoCapture = {
    // Start capturing and downloading segments
    startCapture: (downloadImmediately = true) => {
      isCapturing = true;
      downloadAsCapture = downloadImmediately;
      segmentIndex = 0;
      mseSegmentIndex = 0;
      videoSegments.clear();
      audioSegments.clear();
      // Clear MSE buffers
      mseVideoBuffers.length = 0;
      mseAudioBuffers.length = 0;
      mseVideoInit = null;
      mseAudioInit = null;
      stats = { videoCount: 0, audioCount: 0, videoPending: 0, audioPending: 0, maxTime: 0, errors: 0 };
      console.log('[Teams Video] Capture started (MSE + fetch interception)');
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
      // Report MSE counts if available (preferred), else fetch counts
      videoCount: mseVideoBuffers.length || videoSegments.size,
      audioCount: mseAudioBuffers.length || audioSegments.size,
      // MSE-specific stats
      mseVideoCount: mseVideoBuffers.length,
      mseAudioCount: mseAudioBuffers.length,
      hasVideoInit: !!mseVideoInit,
      hasAudioInit: !!mseAudioInit,
      // Fetch-based stats
      fetchVideoCount: videoSegments.size,
      fetchAudioCount: audioSegments.size,
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
      // Prefer MSE captured data (decrypted with init segments)
      // Fall back to fetch-captured data if MSE not available

      let videoBlob = null;
      let audioBlob = null;
      let videoCount = 0;
      let audioCount = 0;
      let source = 'none';

      // Try MSE captured data first (has init segments for proper playback)
      if (mseVideoInit && mseVideoBuffers.length > 0) {
        videoBlob = buildMseBlob(mseVideoInit, mseVideoBuffers, 'video/mp4');
        videoCount = mseVideoBuffers.length;
        source = 'MSE';
        console.log(`[Teams Video] Using MSE video: init(${mseVideoInit.byteLength}) + ${mseVideoBuffers.length} segments`);
      } else if (videoSegments.size > 0) {
        // Fallback to fetch-captured (may be encrypted for DRM content)
        const sortedVideo = Array.from(videoSegments.entries())
          .sort((a, b) => a[0] - b[0])
          .map(([_, buffer]) => buffer);
        videoBlob = new Blob(sortedVideo, { type: 'video/mp4' });
        videoCount = sortedVideo.length;
        source = 'fetch';
        console.log(`[Teams Video] Using fetch video: ${sortedVideo.length} segments (may be encrypted)`);
      }

      if (mseAudioInit && mseAudioBuffers.length > 0) {
        audioBlob = buildMseBlob(mseAudioInit, mseAudioBuffers, 'audio/mp4');
        audioCount = mseAudioBuffers.length;
        console.log(`[Teams Video] Using MSE audio: init(${mseAudioInit.byteLength}) + ${mseAudioBuffers.length} segments`);
      } else if (audioSegments.size > 0) {
        const sortedAudio = Array.from(audioSegments.entries())
          .sort((a, b) => a[0] - b[0])
          .map(([_, buffer]) => buffer);
        audioBlob = new Blob(sortedAudio, { type: 'audio/mp4' });
        audioCount = sortedAudio.length;
        console.log(`[Teams Video] Using fetch audio: ${sortedAudio.length} segments (may be encrypted)`);
      }

      return {
        videoBlob,
        audioBlob,
        videoCount,
        audioCount,
        source,
        hasMseInit: !!(mseVideoInit || mseAudioInit)
      };
    },

    // Clear captured segments
    clear: () => {
      videoSegments.clear();
      audioSegments.clear();
      // Clear MSE buffers
      mseVideoBuffers.length = 0;
      mseAudioBuffers.length = 0;
      mseVideoInit = null;
      mseAudioInit = null;
      mseSegmentIndex = 0;
      // Clear other state
      capturedVideoToken = null;
      capturedAudioToken = null;
      capturedVideoUrlTemplate = null;
      capturedAudioUrlTemplate = null;
      capturedUrls = [];
      manifestUrl = null;
      segmentIndex = 0;
      stats = { videoCount: 0, audioCount: 0, videoPending: 0, audioPending: 0, maxTime: 0, errors: 0 };
      dispatchUpdate();
    },

    // Debug: Log current state
    debug: () => {
      console.log('[Teams Video Debug]');
      console.log('  isCapturing:', isCapturing);
      console.log('  === MSE Capture (decrypted) ===');
      console.log('  mseVideoInit:', mseVideoInit ? `${mseVideoInit.byteLength} bytes` : 'none');
      console.log('  mseAudioInit:', mseAudioInit ? `${mseAudioInit.byteLength} bytes` : 'none');
      console.log('  mseVideoBuffers:', mseVideoBuffers.length);
      console.log('  mseAudioBuffers:', mseAudioBuffers.length);
      console.log('  === Fetch Capture (may be encrypted) ===');
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
        result = { available: true, version: 'v6.4-mux' };
        break;

      case 'analyzeUrls':
        // Analyze captured URLs to understand the pattern
        result = {
          manifestUrl: manifestUrl,
          capturedUrls: capturedUrls,
          videoToken: capturedVideoToken ? capturedVideoToken.substring(0, 50) + '...' : null,
          audioToken: capturedAudioToken ? capturedAudioToken.substring(0, 50) + '...' : null,
          videoTemplate: capturedVideoUrlTemplate,
          audioTemplate: capturedAudioUrlTemplate
        };
        console.log('[Teams Video] URL Analysis:', result);
        break;

      case 'directDownload':
        // Attempt to download all segments directly using API
        try {
          notify('Starting direct API download...');

          if (capturedUrls.length === 0) {
            throw new Error('No URLs captured yet. Play the video for a few seconds first.');
          }

          // Find a video segment URL (not audio)
          const videoUrl = capturedUrls.find(u => u.url.includes('track=video'));
          if (!videoUrl) {
            throw new Error('No video segment URL captured');
          }

          const sampleUrl = videoUrl.url;
          notify('Analyzing URL pattern...');
          console.log('[Teams Video] Sample URL:', sampleUrl.substring(0, 200));

          // Parse URL parameters
          const urlObj = new URL(sampleUrl);
          const segmentTime = parseInt(urlObj.searchParams.get('segmentTime') || '0');
          const wsd = parseInt(urlObj.searchParams.get('wsd') || '96000'); // segment duration in ms
          const tempauth = urlObj.searchParams.get('tempauth');

          // Check for DRM encryption indicators
          const enableEncryption = urlObj.searchParams.get('enableEncryption');
          const kid = urlObj.searchParams.get('kid');
          const isDrmProtected = enableEncryption === '1' || !!kid;

          console.log('[Teams Video] segmentTime:', segmentTime, 'wsd:', wsd, 'tempauth:', tempauth ? 'present' : 'missing');
          console.log('[Teams Video] DRM check: enableEncryption=', enableEncryption, 'kid=', kid ? 'present' : 'none');

          if (isDrmProtected) {
            notify('⚠️ DRM detected - segments will be encrypted');
            console.log('[Teams Video] WARNING: Content is DRM-protected. Downloaded segments will be encrypted and unplayable.');
          }

          if (!tempauth) {
            throw new Error('No tempauth token found in URL');
          }

          // Get video duration
          const video = document.querySelector('video');
          if (!video || !video.duration) {
            throw new Error('Could not get video duration');
          }

          const totalDurationMs = video.duration * 1000;
          const totalSegments = Math.ceil(totalDurationMs / wsd);
          notify(`Video: ${Math.round(video.duration)}s, ${totalSegments} segments @ ${wsd/1000}s each`);

          // Test: Try to download segment 0 with the captured URL pattern
          notify('Testing API access with segment 0...');

          // Build URL for segment 0
          const testUrl = sampleUrl.replace(/segmentTime=\d+/, 'segmentTime=0');
          console.log('[Teams Video] Test URL:', testUrl.substring(0, 200));

          const testResponse = await originalFetch(testUrl);

          if (testResponse.ok) {
            const testBuffer = await testResponse.arrayBuffer();
            const testSizeKB = (testBuffer.byteLength / 1024).toFixed(1);
            notify(`✓ Segment 0 downloaded (${testSizeKB} KB)`);

            // SUCCESS! Now download all segments
            notify(`Downloading all ${totalSegments} video segments...`);

            const downloadedSegments = [];
            let failCount = 0;

            for (let i = 0; i < totalSegments && failCount < 3; i++) {
              const segTime = i * wsd;
              const segUrl = sampleUrl.replace(/segmentTime=\d+/, `segmentTime=${segTime}`);

              try {
                const response = await originalFetch(segUrl);
                if (response.ok) {
                  const buffer = await response.arrayBuffer();
                  downloadedSegments.push({ index: i, buffer });
                  if (i % 5 === 0 || i === totalSegments - 1) {
                    notify(`Downloaded ${i + 1}/${totalSegments} segments...`);
                  }
                } else {
                  console.warn(`[Teams Video] Segment ${i} failed: ${response.status}`);
                  failCount++;
                }
              } catch (err) {
                console.warn(`[Teams Video] Segment ${i} error:`, err);
                failCount++;
              }
            }

            if (downloadedSegments.length > 0) {
              // Store in videoSegments map for later saving
              downloadedSegments.forEach(seg => {
                videoSegments.set(seg.index, seg.buffer);
              });
              stats.videoCount = videoSegments.size;
              dispatchUpdate();

              notify(`✓ Downloaded ${downloadedSegments.length}/${totalSegments} video segments!`);

              // Now try to download audio segments
              const audioUrl = capturedUrls.find(u => u.url.includes('track=audio'));
              let audioDownloaded = 0;

              if (audioUrl) {
                notify('Downloading audio segments...');
                const audioSampleUrl = audioUrl.url;
                const audioFailMax = 3;
                let audioFailCount = 0;

                for (let i = 0; i < totalSegments && audioFailCount < audioFailMax; i++) {
                  const segTime = i * wsd;
                  const audioSegUrl = audioSampleUrl.replace(/segmentTime=\d+/, `segmentTime=${segTime}`);

                  try {
                    const response = await originalFetch(audioSegUrl);
                    if (response.ok) {
                      const buffer = await response.arrayBuffer();
                      audioSegments.set(i, buffer);
                      audioDownloaded++;
                      if (i % 10 === 0 || i === totalSegments - 1) {
                        notify(`Audio: ${audioDownloaded}/${totalSegments}...`);
                      }
                    } else {
                      audioFailCount++;
                    }
                  } catch (err) {
                    audioFailCount++;
                  }
                }

                stats.audioCount = audioSegments.size;
                dispatchUpdate();
              }

              if (isDrmProtected) {
                notify(`⚠️ Downloaded ${downloadedSegments.length}+${audioDownloaded} ENCRYPTED segments. Use 🚀 Start Capture for playable video!`);
                result = {
                  success: true,
                  videoDownloaded: downloadedSegments.length,
                  audioDownloaded: audioDownloaded,
                  total: totalSegments,
                  isDrmProtected: true,
                  message: `⚠️ DRM: ${downloadedSegments.length}+${audioDownloaded} encrypted. Use 🚀 Start Capture instead!`
                };
              } else {
                notify(`✓ Complete! Video: ${downloadedSegments.length}, Audio: ${audioDownloaded}. Click Save Files!`);
                result = {
                  success: true,
                  videoDownloaded: downloadedSegments.length,
                  audioDownloaded: audioDownloaded,
                  total: totalSegments,
                  isDrmProtected: false,
                  message: `Video: ${downloadedSegments.length}, Audio: ${audioDownloaded} segments. Click Save Files!`
                };
              }
            } else {
              result = {
                success: false,
                message: 'All segment downloads failed'
              };
            }
          } else {
            notify(`✗ Test failed: ${testResponse.status}`);
            result = {
              success: false,
              status: testResponse.status,
              message: `API access failed (${testResponse.status}). Token may have expired.`
            };
          }
        } catch (err) {
          console.error('[Teams Video] Direct download error:', err);
          result = { success: false, error: err.message };
        }
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

      case 'downloadCombined':
        // Download combined video+audio file using inlined MP4 muxer
        try {
          notify('Preparing combined download...');

          // Check if we have MSE captured data
          if (!mseVideoInit || mseVideoBuffers.length === 0) {
            throw new Error('No MSE video data captured. Use Start Capture first.');
          }

          const hasAudioData = mseAudioInit && mseAudioBuffers.length > 0;
          notify(`Found: ${mseVideoBuffers.length} video + ${mseAudioBuffers.length} audio segments`);

          notify('Muxing video and audio...');

          // Use the inlined muxer to combine
          const combinedBuffer = MP4Muxer.mux(
            mseVideoInit,
            mseVideoBuffers,
            hasAudioData ? mseAudioInit : null,
            hasAudioData ? mseAudioBuffers : []
          );

          if (!combinedBuffer) {
            throw new Error('Muxing failed - no ftyp/moov boxes found');
          }

          const sizeMB = (combinedBuffer.byteLength / 1024 / 1024).toFixed(1);
          notify(`Combined file ready: ${sizeMB} MB`);

          // Get video title
          const titleEl = document.querySelector('h1[class*="videoTitleViewModeHeading"] label');
          const title = titleEl?.innerText?.trim() || document.title?.trim() || 'video';
          const safeTitle = title.replace(/[^a-zA-Z0-9\s-]/g, '').trim() || 'recording';

          // Download combined file
          notify('Downloading combined file...');
          const blob = new Blob([combinedBuffer], { type: 'video/mp4' });
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `${safeTitle}.mp4`;
          a.style.display = 'none';
          document.body.appendChild(a);
          a.click();
          await new Promise(r => setTimeout(r, 500));
          document.body.removeChild(a);
          URL.revokeObjectURL(url);

          notify(`✅ Downloaded: ${safeTitle}.mp4 (${sizeMB} MB)`);
          result = {
            success: true,
            size: combinedBuffer.byteLength,
            sizeMB: sizeMB,
            title: safeTitle,
            hasAudio: hasAudioData
          };
        } catch (err) {
          console.error('[Teams Video] Combined download error:', err);
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
  document.dispatchEvent(new CustomEvent('teamsVideoReady', { detail: { version: 'v6.4-mux' } }));

  console.log('[Teams Chat Exporter] Video download override v6.4 installed (MSE capture + muxing)');
})();
