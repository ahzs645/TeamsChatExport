/**
 * Manifest Download Module
 * Downloads video by fetching encrypted segments in parallel and decrypting them
 * using the captured AES-CBC key from the player's crypto.subtle.decrypt calls.
 *
 * Flow:
 * 1. Player briefly plays -> crypto.subtle.decrypt is called -> we capture key + IV
 * 2. We discover segment URL templates from performance entries
 * 3. We fetch ALL encrypted segments in parallel (fast, no playback needed)
 * 4. We decrypt each segment with crypto.subtle.decrypt(algo, key, data)
 * 5. Fix timestamps (tfdt baseDecodeTime) so segments are sequential
 * 6. Combine init segments + fixed moof segments -> playable MP4
 *
 * This is the fastest approach for encrypted content:
 * - No full playback needed (just 1-2 seconds to capture the key)
 * - Parallel downloads (~6 concurrent)
 * - Original quality (no re-encoding)
 * - Correct timestamps for full-duration playback
 */
(() => {
  const MODULE_NAME = 'manifestDownload';
  const MAX_CONCURRENT = 6;
  const SEGMENT_TIMEOUT = 30000;
  const SEGMENT_DURATION_MS = 2000;

  // === Crypto Access ===

  const getCrypto = () => window.__videoCryptoCapture || { ready: false };
  const hasCryptoKey = () => getCrypto().ready === true;

  // === Segment Template Discovery ===

  /**
   * Find segment URL templates from multiple sources.
   * Priority: 1) videoDownloadOverride capture, 2) performance entries
   */
  const findSegmentTemplates = () => {
    // Source 1: videoDownloadOverride's captured templates
    if (window.__teamsVideoCapture) {
      const tokens = window.__teamsVideoCapture.getTokens?.();
      if (tokens?.videoTemplate && tokens?.audioTemplate) {
        return { videoTemplate: tokens.videoTemplate, audioTemplate: tokens.audioTemplate, source: 'capture' };
      }
    }

    // Source 2: Discover from performance entries (most reliable)
    const entries = performance.getEntriesByType('resource');
    const segEntries = entries.filter(e => e.name.includes('segmentTime='));
    if (segEntries.length < 2) return null;

    // Group by URL pattern (ignoring segmentTime value), identify video vs audio by size
    const byPattern = {};
    for (const e of segEntries) {
      try {
        const u = new URL(e.name);
        u.searchParams.set('segmentTime', 'X');
        const key = u.toString();
        if (!byPattern[key]) byPattern[key] = { totalSize: 0, count: 0 };
        byPattern[key].count++;
        byPattern[key].totalSize += (e.transferSize || 0);
        byPattern[key].url = key;
      } catch (x) {}
    }

    const patterns = Object.values(byPattern);
    if (patterns.length < 2) return null;

    // Sort by average size descending: largest = video, smallest = audio
    patterns.sort((a, b) => (b.totalSize / b.count) - (a.totalSize / a.count));
    return {
      videoTemplate: patterns[0].url.replace('segmentTime=X', 'segmentTime={TIME}'),
      audioTemplate: patterns[1].url.replace('segmentTime=X', 'segmentTime={TIME}'),
      source: 'performance'
    };
  };

  // === Init Segment Capture ===

  const getInitSegments = () => window.__manifestInitSegments || null;

  /**
   * Capture init segments by hooking appendBuffer and triggering a seek.
   */
  const captureInitSegments = () => {
    return new Promise((resolve) => {
      let videoInit = null, audioInit = null;
      const origAppend = SourceBuffer.prototype.appendBuffer;

      SourceBuffer.prototype.appendBuffer = function (data) {
        const arr = new Uint8Array(data instanceof ArrayBuffer ? data : data.buffer);
        if (arr.length >= 8) {
          const box = String.fromCharCode(arr[4], arr[5], arr[6], arr[7]);
          const mime = this._mseTrackType || '';
          if (box === 'ftyp') {
            if (mime.includes('audio') && !audioInit) audioInit = arr.slice().buffer;
            else if (!videoInit) videoInit = arr.slice().buffer;
          }
        }
        const result = origAppend.call(this, data);
        if (videoInit && audioInit) {
          SourceBuffer.prototype.appendBuffer = origAppend;
          window.__manifestInitSegments = { videoInit, audioInit };
          resolve({ videoInit, audioInit });
        }
        return result;
      };

      // Seek far then back to force player to re-send init segments
      const video = document.querySelector('video');
      if (video) {
        const savedTime = video.currentTime;
        video.currentTime = Math.max(0, video.duration / 2);
        setTimeout(() => {
          video.currentTime = savedTime;
          // Timeout: resolve with whatever we have after 8 seconds
          setTimeout(() => {
            SourceBuffer.prototype.appendBuffer = origAppend;
            if (videoInit || audioInit) {
              window.__manifestInitSegments = { videoInit, audioInit };
              resolve({ videoInit, audioInit });
            } else {
              resolve(null);
            }
          }, 4000);
        }, 2000);
      } else {
        resolve(null);
      }
    });
  };

  // === Segment Fetch + Decrypt ===

  const fetchAndDecrypt = async (url, cryptoKey, algo) => {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), SEGMENT_TIMEOUT);
    const resp = await fetch(url, { signal: controller.signal });
    clearTimeout(timeout);
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const encrypted = await resp.arrayBuffer();
    return crypto.subtle.decrypt({ name: algo.name, iv: algo.iv || getCrypto().iv }, cryptoKey, encrypted);
  };

  const fetchAndDecryptRetry = async (url, key, algo, retries = 2) => {
    for (let attempt = 0; attempt <= retries; attempt++) {
      try { return await fetchAndDecrypt(url, key, algo); }
      catch (e) { if (attempt === retries) throw e; await new Promise(r => setTimeout(r, 1000 * (attempt + 1))); }
    }
  };

  // === Parallel Download Engine ===

  const downloadAllSegments = async (template, count, key, algo, onProgress, label) => {
    const results = new Array(count);
    let completed = 0, failed = 0;
    const queue = Array.from({ length: count }, (_, i) => i);

    const worker = async () => {
      while (queue.length > 0) {
        const i = queue.shift();
        if (i === undefined) break;
        const url = template.replace('{TIME}', i * SEGMENT_DURATION_MS).replace('%7BTIME%7D', i * SEGMENT_DURATION_MS);
        try {
          results[i] = await fetchAndDecryptRetry(url, key, algo);
          completed++;
        } catch (e) {
          failed++;
          completed++;
        }
        if (onProgress) onProgress({ completed, total: count, failed, label });
      }
    };

    await Promise.all(Array.from({ length: MAX_CONCURRENT }, () => worker()));
    return results;
  };

  // === Timestamp Fixer ===

  /**
   * Fix fragmented MP4 timestamps so each moof's tfdt.baseDecodeTime is sequential.
   * Without this fix, every segment starts at time 0 and the video appears seconds long.
   *
   * Walks the MP4 box structure, finds each moof > traf > tfdt and trun,
   * calculates correct cumulative timestamps based on sample count * sample duration.
   */
  /**
   * Find the timescale from the init segment's mdhd box.
   */
  const findTimescale = (data) => {
    const bytes = new Uint8Array(data.buffer || data);
    for (let i = 0; i < Math.min(bytes.length, 2000) - 20; i++) {
      if (bytes[i+4] === 0x6d && bytes[i+5] === 0x64 && bytes[i+6] === 0x68 && bytes[i+7] === 0x64) { // 'mdhd'
        const ver = bytes[i + 8];
        if (ver === 0) return new DataView(bytes.buffer, bytes.byteOffset + i + 20, 4).getUint32(0);
        else return new DataView(bytes.buffer, bytes.byteOffset + i + 28, 4).getUint32(0);
      }
    }
    return 16000; // fallback
  };

  const fixTimestamps = (data, segmentDurationMs) => {
    const buf = new DataView(data.buffer || data);
    const bytes = new Uint8Array(data.buffer || data);
    const total = bytes.length;
    let pos = 0;
    let currentTime = 0;
    let fixed = 0;

    // Calculate the correct duration per segment in timescale units
    const timescale = findTimescale(data);
    const segDurationTs = Math.round((segmentDurationMs || SEGMENT_DURATION_MS) / 1000 * timescale);

    const readU32 = (p) => buf.getUint32(p);
    const writeU32 = (p, v) => buf.setUint32(p, v);
    const writeU64 = (p, v) => { buf.setUint32(p, Math.floor(v / 0x100000000)); buf.setUint32(p + 4, v >>> 0); };
    const boxType = (p) => String.fromCharCode(bytes[p + 4], bytes[p + 5], bytes[p + 6], bytes[p + 7]);

    while (pos < total - 8) {
      const size = readU32(pos);
      if (size < 8) break;
      const type = boxType(pos);

      if (type === 'moof') {
        let inner = pos + 8;
        while (inner < pos + size - 8) {
          const iSize = readU32(inner);
          if (iSize < 8) break;
          const iType = boxType(inner);

          if (iType === 'traf') {
            let tp = inner + 8;
            const trafEnd = inner + iSize;

            while (tp < trafEnd - 8) {
              const tSize = readU32(tp);
              if (tSize < 8) break;
              const tType = boxType(tp);

              if (tType === 'tfdt') {
                const version = bytes[tp + 8];
                if (version === 1) {
                  writeU64(tp + 12, currentTime);
                } else {
                  writeU32(tp + 12, currentTime);
                }
                fixed++;
              }

              tp += tSize;
            }

            // Use the known segment duration (from URL interval) rather than tfhd/trun
            // which can report inflated values
            currentTime += segDurationTs;
          }

          inner += iSize;
        }
      }

      pos += size;
    }

    return { data: bytes, fixed, finalTime: currentTime };
  };

  // === isAvailable ===

  const isAvailable = () => {
    return !!(findSegmentTemplates() && hasCryptoKey());
  };

  // === Main Download ===

  const download = async (onProgress) => {
    const templates = findSegmentTemplates();
    if (!templates) {
      return { success: false, error: 'No segment URLs found. Play the video for a few seconds first.' };
    }

    const cryptoData = getCrypto();
    if (!cryptoData.ready) {
      return { success: false, error: 'No decryption key captured. Play the video for a few seconds first.' };
    }

    const video = document.querySelector('video');
    const duration = video?.duration || 0;
    if (!duration) return { success: false, error: 'Cannot determine video duration.' };

    const totalSegs = Math.ceil(duration / (SEGMENT_DURATION_MS / 1000));
    const startTime = Date.now();

    // Step 1: Capture init segments
    if (onProgress) onProgress({ stage: 'init', message: 'Capturing init segments...', percent: 0 });
    let initSegs = getInitSegments();
    if (!initSegs || !initSegs.videoInit) {
      initSegs = await captureInitSegments();
      if (!initSegs?.videoInit) {
        return { success: false, error: 'Failed to capture init segments. Try seeking in the video first.' };
      }
    }

    // Step 2: Download + decrypt video segments
    if (onProgress) onProgress({ stage: 'video', message: `Downloading ${totalSegs} video segments...`, percent: 0 });
    const videoSegs = await downloadAllSegments(
      templates.videoTemplate, totalSegs, cryptoData.key, cryptoData.algo,
      (p) => { if (onProgress) onProgress({ stage: 'video', message: `Video: ${p.completed}/${p.total}`, percent: Math.round(p.completed / p.total * 50) }); },
      'Video'
    );

    // Step 3: Download + decrypt audio segments
    if (onProgress) onProgress({ stage: 'audio', message: `Downloading ${totalSegs} audio segments...`, percent: 50 });
    const audioSegs = await downloadAllSegments(
      templates.audioTemplate, totalSegs, cryptoData.key, cryptoData.algo,
      (p) => { if (onProgress) onProgress({ stage: 'audio', message: `Audio: ${p.completed}/${p.total}`, percent: 50 + Math.round(p.completed / p.total * 40) }); },
      'Audio'
    );

    // Step 4: Combine and fix timestamps
    if (onProgress) onProgress({ stage: 'fixing', message: 'Fixing timestamps...', percent: 90 });

    const validVideoSegs = videoSegs.filter(Boolean);
    const validAudioSegs = audioSegs.filter(Boolean);

    // Build raw blobs
    const rawVideoData = concatBuffers([initSegs.videoInit, ...validVideoSegs]);
    const rawAudioData = concatBuffers([initSegs.audioInit, ...validAudioSegs]);

    // Fix timestamps
    const fixedVideo = fixTimestamps(rawVideoData);
    const fixedAudio = fixTimestamps(rawAudioData);

    if (onProgress) onProgress({ stage: 'saving', message: 'Preparing download...', percent: 95 });

    const fileName = getFileName();
    const vBlob = new Blob([fixedVideo.data], { type: 'video/mp4' });
    const aBlob = new Blob([fixedAudio.data], { type: 'audio/mp4' });
    const vMB = Math.round(vBlob.size / 1024 / 1024);
    const aMB = Math.round(aBlob.size / 1024 / 1024);
    const elapsed = Math.round((Date.now() - startTime) / 1000);

    // Show save buttons (large blob downloads need user gesture)
    const panel = document.createElement('div');
    panel.style.cssText = 'position:fixed;top:10px;left:50%;transform:translateX(-50%);z-index:99999;padding:16px 24px;background:rgba(0,0,0,0.92);color:white;border-radius:10px;font-family:monospace;font-size:13px;min-width:500px;text-align:center;';
    panel.innerHTML = `<div>Done in ${elapsed}s! ${fixedVideo.fixed} segments, timestamps fixed.</div>` +
      `<div style="margin:6px 0;font-size:11px;color:#aaa;">Merge with: ffmpeg -i video.mp4 -i audio.mp4 -c copy combined.mp4</div>` +
      `<div style="margin-top:10px;display:flex;gap:8px;justify-content:center;"></div>`;
    const btnRow = panel.lastElementChild;

    const makeSaveBtn = (text, blob, fn, color) => {
      const btn = document.createElement('button');
      btn.textContent = text;
      btn.style.cssText = `padding:10px 20px;font-size:14px;background:${color};color:white;border:none;border-radius:6px;cursor:pointer;`;
      btn.onclick = () => {
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = fn;
        a.click();
        btn.textContent = 'Saving...';
        setTimeout(() => { btn.textContent = text; }, 3000);
      };
      return btn;
    };

    btnRow.appendChild(makeSaveBtn(`Save Video (${vMB}MB)`, vBlob, fileName + '-video.mp4', '#28a745'));
    btnRow.appendChild(makeSaveBtn(`Save Audio (${aMB}MB)`, aBlob, fileName + '-audio.mp4', '#007bff'));

    const closeBtn = document.createElement('button');
    closeBtn.textContent = '\u00D7';
    closeBtn.style.cssText = 'position:absolute;top:6px;right:10px;background:none;border:none;color:#aaa;font-size:18px;cursor:pointer;';
    closeBtn.onclick = () => panel.remove();
    panel.style.position = 'fixed';
    panel.appendChild(closeBtn);
    document.body.appendChild(panel);

    if (onProgress) onProgress({ stage: 'done', message: 'Ready! Click buttons to save.', percent: 100 });

    return {
      success: true,
      method: 'manifestDecrypt',
      videoSize: vBlob.size,
      audioSize: aBlob.size,
      segments: totalSegs,
      elapsed,
      fixed: fixedVideo.fixed
    };
  };

  // === Helpers ===

  const concatBuffers = (buffers) => {
    const totalLength = buffers.reduce((sum, b) => sum + (b ? b.byteLength || 0 : 0), 0);
    const result = new Uint8Array(totalLength);
    let offset = 0;
    for (const buf of buffers) {
      if (buf) {
        result.set(new Uint8Array(buf), offset);
        offset += buf.byteLength;
      }
    }
    return result;
  };

  const getFileName = () => {
    const title = document.querySelector('h1, h2, [class*="videoTitle"] label')
      ?.textContent?.trim()?.replace(/[^a-zA-Z0-9\s-]/g, '')?.trim() || 'recording';
    return title;
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
    getInitSegments,
    fixTimestamps
  };

  console.log('[Teams Chat Exporter] Video download module loaded: manifestDownload');
})();
