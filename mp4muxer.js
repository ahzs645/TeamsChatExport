/**
 * Simple fMP4 Muxer
 * Combines separate video and audio fMP4 streams into a single file
 */

const MP4Muxer = (() => {
  // Read a 32-bit big-endian integer
  const readUint32 = (data, offset) => {
    return (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
  };

  // Write a 32-bit big-endian integer
  const writeUint32 = (data, offset, value) => {
    data[offset] = (value >> 24) & 0xff;
    data[offset + 1] = (value >> 16) & 0xff;
    data[offset + 2] = (value >> 8) & 0xff;
    data[offset + 3] = value & 0xff;
  };

  // Get box type as string
  const getBoxType = (data, offset) => {
    return String.fromCharCode(data[offset], data[offset + 1], data[offset + 2], data[offset + 3]);
  };

  // Find a box within data
  const findBox = (data, type, start = 0, end = null) => {
    end = end || data.length;
    let offset = start;

    while (offset < end - 8) {
      const size = readUint32(data, offset);
      const boxType = getBoxType(data, offset + 4);

      if (size < 8) break; // Invalid box

      if (boxType === type) {
        return { offset, size };
      }

      offset += size;
    }

    return null;
  };

  // Find all boxes of a type
  const findAllBoxes = (data, type, start = 0, end = null) => {
    end = end || data.length;
    const boxes = [];
    let offset = start;

    while (offset < end - 8) {
      const size = readUint32(data, offset);
      const boxType = getBoxType(data, offset + 4);

      if (size < 8) break;

      if (boxType === type) {
        boxes.push({ offset, size });
      }

      offset += size;
    }

    return boxes;
  };

  // Find a nested box (e.g., moov/trak/tkhd)
  const findNestedBox = (data, path, start = 0, end = null) => {
    let currentStart = start;
    let currentEnd = end || data.length;

    for (const boxType of path) {
      const box = findBox(data, boxType, currentStart, currentEnd);
      if (!box) return null;

      currentStart = box.offset + 8; // Skip size and type
      currentEnd = box.offset + box.size;
    }

    return { offset: currentStart - 8, size: currentEnd - currentStart + 8 };
  };

  // Extract track ID from tkhd box
  const getTrackId = (data, tkhdOffset) => {
    // tkhd box: 8 bytes header + 1 byte version + 3 bytes flags + ...
    // Track ID is at offset 12 (v0) or 20 (v1) from box start
    const version = data[tkhdOffset + 8];
    const trackIdOffset = tkhdOffset + (version === 1 ? 20 : 12);
    return readUint32(data, trackIdOffset);
  };

  // Set track ID in tkhd box
  const setTrackId = (data, tkhdOffset, newId) => {
    const version = data[tkhdOffset + 8];
    const trackIdOffset = tkhdOffset + (version === 1 ? 20 : 12);
    writeUint32(data, trackIdOffset, newId);
  };

  // Update track ID in tfhd box (inside moof/traf)
  const updateTfhdTrackId = (data, moofOffset, moofSize, newTrackId) => {
    // Find traf box inside moof
    const trafBox = findBox(data, 'traf', moofOffset + 8, moofOffset + moofSize);
    if (!trafBox) return false;

    // Find tfhd box inside traf
    const tfhdBox = findBox(data, 'tfhd', trafBox.offset + 8, trafBox.offset + trafBox.size);
    if (!tfhdBox) return false;

    // tfhd: 8 bytes header + 1 byte version + 3 bytes flags + 4 bytes track_ID
    writeUint32(data, tfhdBox.offset + 12, newTrackId);
    return true;
  };

  // Combine video and audio fMP4 streams
  const mux = (videoInit, videoSegments, audioInit, audioSegments) => {
    if (!videoInit || videoSegments.length === 0) {
      console.error('[MP4Muxer] No video data');
      return null;
    }

    const videoData = new Uint8Array(videoInit);
    const hasAudio = audioInit && audioSegments.length > 0;
    const audioData = hasAudio ? new Uint8Array(audioInit) : null;

    // Find ftyp box from video
    const ftypBox = findBox(videoData, 'ftyp');
    if (!ftypBox) {
      console.error('[MP4Muxer] No ftyp box found');
      return null;
    }

    // Find moov boxes
    const videoMoov = findBox(videoData, 'moov');
    if (!videoMoov) {
      console.error('[MP4Muxer] No video moov box found');
      return null;
    }

    let audioMoov = null;
    if (hasAudio) {
      audioMoov = findBox(audioData, 'moov');
    }

    // For simplicity, we'll use the video moov and append audio trak if present
    // A full implementation would properly merge moov boxes

    // Calculate output size
    let totalSize = ftypBox.size + videoMoov.size;

    // Add video segments
    for (const seg of videoSegments) {
      totalSize += seg.byteLength;
    }

    // Add audio segments if present
    if (hasAudio) {
      for (const seg of audioSegments) {
        totalSize += seg.byteLength;
      }
    }

    // Create output buffer
    const output = new Uint8Array(totalSize);
    let writeOffset = 0;

    // Write ftyp
    output.set(videoData.subarray(ftypBox.offset, ftypBox.offset + ftypBox.size), writeOffset);
    writeOffset += ftypBox.size;

    // Write moov (just video for now - full merge is complex)
    output.set(videoData.subarray(videoMoov.offset, videoMoov.offset + videoMoov.size), writeOffset);
    writeOffset += videoMoov.size;

    // Interleave video and audio segments
    const maxSegments = Math.max(videoSegments.length, hasAudio ? audioSegments.length : 0);

    for (let i = 0; i < maxSegments; i++) {
      // Write video segment
      if (i < videoSegments.length) {
        const seg = new Uint8Array(videoSegments[i]);
        output.set(seg, writeOffset);
        writeOffset += seg.byteLength;
      }

      // Write audio segment
      if (hasAudio && i < audioSegments.length) {
        const seg = new Uint8Array(audioSegments[i]);
        // Update track ID in tfhd to 2 (audio track)
        const segCopy = new Uint8Array(seg);
        const moofBox = findBox(segCopy, 'moof');
        if (moofBox) {
          updateTfhdTrackId(segCopy, moofBox.offset, moofBox.size, 2);
        }
        output.set(segCopy, writeOffset);
        writeOffset += segCopy.byteLength;
      }
    }

    console.log(`[MP4Muxer] Created ${(output.byteLength / 1024 / 1024).toFixed(1)} MB combined file`);
    return output.buffer;
  };

  // Alternative: Simple concatenation (video only with embedded audio if present in same stream)
  const muxSimple = (videoInit, videoSegments, audioInit, audioSegments) => {
    // If we have both separate streams, try the complex mux
    if (audioInit && audioSegments.length > 0) {
      return mux(videoInit, videoSegments, audioInit, audioSegments);
    }

    // Otherwise just concatenate video
    const totalSize = videoInit.byteLength + videoSegments.reduce((sum, s) => sum + s.byteLength, 0);
    const output = new Uint8Array(totalSize);
    let offset = 0;

    output.set(new Uint8Array(videoInit), offset);
    offset += videoInit.byteLength;

    for (const seg of videoSegments) {
      output.set(new Uint8Array(seg), offset);
      offset += seg.byteLength;
    }

    return output.buffer;
  };

  return { mux, muxSimple, findBox, findAllBoxes };
})();

// Export for use in other scripts
if (typeof window !== 'undefined') {
  window.MP4Muxer = MP4Muxer;
}
