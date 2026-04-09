/**
 * Fragmented MP4 to Standard MP4 Converter
 * Converts fMP4 (moof/mdat segments) to a standard MP4 with a single moov+mdat.
 * This makes the output playable in QuickTime and other players that don't support fMP4.
 *
 * Input: Uint8Array of fMP4 data (ftyp + moov + [moof+mdat]...)
 * Output: Uint8Array of standard MP4 (ftyp + moov + mdat)
 */
(() => {
  // === Box parsing helpers ===

  const readU32 = (d, p) => (d[p] << 24) | (d[p+1] << 16) | (d[p+2] << 8) | d[p+3];
  const readU16 = (d, p) => (d[p] << 8) | d[p+1];
  const boxType = (d, p) => String.fromCharCode(d[p+4], d[p+5], d[p+6], d[p+7]);

  const writeU32 = (d, p, v) => { d[p] = (v >>> 24) & 0xff; d[p+1] = (v >>> 16) & 0xff; d[p+2] = (v >>> 8) & 0xff; d[p+3] = v & 0xff; };
  const writeU16 = (d, p, v) => { d[p] = (v >>> 8) & 0xff; d[p+1] = v & 0xff; };

  /**
   * Find a box at the top level of data starting from offset.
   */
  const findBox = (data, type, start, end) => {
    let pos = start || 0;
    const limit = end || data.length;
    while (pos < limit - 8) {
      const size = readU32(data, pos);
      if (size < 8) return null;
      if (boxType(data, pos) === type) return { offset: pos, size };
      pos += size;
    }
    return null;
  };

  /**
   * Find a nested box inside a container box.
   */
  const findBoxIn = (data, containerOffset, containerSize, type) => {
    return findBox(data, type, containerOffset + 8, containerOffset + containerSize);
  };

  /**
   * Build a box with the given type and content.
   */
  const makeBox = (type, ...contents) => {
    let totalSize = 8;
    for (const c of contents) totalSize += c.length;
    const box = new Uint8Array(totalSize);
    writeU32(box, 0, totalSize);
    box[4] = type.charCodeAt(0);
    box[5] = type.charCodeAt(1);
    box[6] = type.charCodeAt(2);
    box[7] = type.charCodeAt(3);
    let offset = 8;
    for (const c of contents) {
      box.set(c, offset);
      offset += c.length;
    }
    return box;
  };

  /**
   * Build a full-box (with version + flags).
   */
  const makeFullBox = (type, version, flags, ...contents) => {
    let totalSize = 12;
    for (const c of contents) totalSize += c.length;
    const box = new Uint8Array(totalSize);
    writeU32(box, 0, totalSize);
    box[4] = type.charCodeAt(0);
    box[5] = type.charCodeAt(1);
    box[6] = type.charCodeAt(2);
    box[7] = type.charCodeAt(3);
    box[8] = version;
    box[9] = (flags >> 16) & 0xff;
    box[10] = (flags >> 8) & 0xff;
    box[11] = flags & 0xff;
    let offset = 12;
    for (const c of contents) {
      box.set(c, offset);
      offset += c.length;
    }
    return box;
  };

  // === Main converter ===

  /**
   * Convert fragmented MP4 to standard MP4.
   * @param {Uint8Array} fmp4Data - The fragmented MP4 data
   * @param {number} timescale - The media timescale (e.g. 16000)
   * @param {number} segmentDurationMs - Duration of each segment in milliseconds
   * @returns {Uint8Array} Standard MP4 data
   */
  const convert = (fmp4Data, timescale, segmentDurationMs) => {
    const data = fmp4Data;
    const total = data.length;

    // Step 1: Find ftyp and init moov
    const ftyp = findBox(data, 'ftyp', 0);
    const initMoov = findBox(data, 'moov', 0);
    if (!ftyp || !initMoov) {
      throw new Error('Missing ftyp or moov in input');
    }

    // Step 2: Extract codec info from init moov > trak > mdia > minf > stbl > stsd
    const initMoovData = data.slice(initMoov.offset, initMoov.offset + initMoov.size);
    const trak = findBox(initMoovData, 'trak', 8);
    if (!trak) throw new Error('No trak in moov');
    const mdia = findBoxIn(initMoovData, trak.offset, trak.size, 'mdia');
    if (!mdia) throw new Error('No mdia in trak');
    const minf = findBoxIn(initMoovData, mdia.offset, mdia.size, 'minf');
    if (!minf) throw new Error('No minf in mdia');
    const stbl = findBoxIn(initMoovData, minf.offset, minf.size, 'stbl');
    if (!stbl) throw new Error('No stbl in minf');
    const stsd = findBoxIn(initMoovData, stbl.offset, stbl.size, 'stsd');
    if (!stsd) throw new Error('No stsd in stbl');

    // Extract stsd data (codec description)
    const stsdData = initMoovData.slice(stsd.offset, stsd.offset + stsd.size);

    // Extract handler type from hdlr
    const hdlr = findBoxIn(initMoovData, mdia.offset, mdia.size, 'hdlr');
    let handlerType = 'vide'; // default
    if (hdlr) {
      handlerType = String.fromCharCode(
        initMoovData[hdlr.offset + 16],
        initMoovData[hdlr.offset + 17],
        initMoovData[hdlr.offset + 18],
        initMoovData[hdlr.offset + 19]
      );
    }
    const hdlrData = hdlr ? initMoovData.slice(hdlr.offset, hdlr.offset + hdlr.size) : null;

    // Extract tkhd
    const tkhd = findBoxIn(initMoovData, trak.offset, trak.size, 'tkhd');
    const tkhdData = tkhd ? initMoovData.slice(tkhd.offset, tkhd.offset + tkhd.size) : null;

    // Extract mdhd
    const mdhd = findBoxIn(initMoovData, mdia.offset, mdia.size, 'mdhd');
    const mdhdData = mdhd ? initMoovData.slice(mdhd.offset, mdhd.offset + mdhd.size) : null;

    // Known segment duration in timescale units
    const segDurationTs = Math.round(segmentDurationMs / 1000 * timescale);

    // Step 3: Walk all moof+mdat pairs, collect sample info
    const samples = []; // {size, duration}
    const mdatChunks = []; // ArrayBuffer chunks of media data
    let totalMediaSize = 0;
    let pos = initMoov.offset + initMoov.size;

    while (pos < total - 8) {
      const size = readU32(data, pos);
      const type = boxType(data, pos);
      if (size < 8) break;

      if (type === 'moof') {
        let inner = pos + 8;
        while (inner < pos + size - 8) {
          const iSize = readU32(data, inner);
          if (iSize < 8) break;
          const iType = boxType(data, inner);

          if (iType === 'traf') {
            let tp = inner + 8;
            const trafEnd = inner + iSize;
            const sampleSizes = [];

            while (tp < trafEnd - 8) {
              const tSize = readU32(data, tp);
              if (tSize < 8) break;
              const tType = boxType(data, tp);

              if (tType === 'trun') {
                const fl = (data[tp+9] << 16) | (data[tp+10] << 8) | data[tp+11];
                const sampleCount = readU32(data, tp + 12);
                let o = 16;
                if (fl & 0x1) o += 4;
                if (fl & 0x4) o += 4;

                const hasDuration = !!(fl & 0x100);
                const hasSize = !!(fl & 0x200);
                const hasFlags = !!(fl & 0x400);
                const hasCTO = !!(fl & 0x800);

                // Calculate per-sample duration from known segment duration
                const perSampleDur = sampleCount > 0 ? Math.round(segDurationTs / sampleCount) : 1;

                for (let s = 0; s < sampleCount; s++) {
                  if (hasDuration) o += 4; // skip tfhd/trun duration (unreliable)
                  const sz = hasSize ? readU32(data, tp + o) : 0;
                  if (hasSize) o += 4;
                  if (hasFlags) o += 4;
                  if (hasCTO) o += 4;
                  sampleSizes.push({ size: sz, duration: perSampleDur });
                }
              }

              tp += tSize;
            }

            for (const s of sampleSizes) {
              samples.push(s);
            }
          }

          inner += iSize;
        }
      }

      if (type === 'mdat') {
        mdatChunks.push(data.slice(pos + 8, pos + size));
        totalMediaSize += size - 8;
      }

      pos += size;
    }

    if (samples.length === 0) throw new Error('No samples found');

    // Step 4: Build standard MP4 sample tables

    // stts (time-to-sample) - run-length encode durations
    const sttsRuns = [];
    let currentDur = samples[0].duration;
    let currentCount = 1;
    for (let i = 1; i < samples.length; i++) {
      if (samples[i].duration === currentDur) {
        currentCount++;
      } else {
        sttsRuns.push({ count: currentCount, duration: currentDur });
        currentDur = samples[i].duration;
        currentCount = 1;
      }
    }
    sttsRuns.push({ count: currentCount, duration: currentDur });

    const sttsPayload = new Uint8Array(4 + sttsRuns.length * 8);
    writeU32(sttsPayload, 0, sttsRuns.length);
    for (let i = 0; i < sttsRuns.length; i++) {
      writeU32(sttsPayload, 4 + i * 8, sttsRuns[i].count);
      writeU32(sttsPayload, 4 + i * 8 + 4, sttsRuns[i].duration);
    }
    const stts = makeFullBox('stts', 0, 0, sttsPayload);

    // stsz (sample sizes)
    const stszPayload = new Uint8Array(8 + samples.length * 4);
    writeU32(stszPayload, 0, 0); // sample_size = 0 (variable)
    writeU32(stszPayload, 4, samples.length);
    for (let i = 0; i < samples.length; i++) {
      writeU32(stszPayload, 8 + i * 4, samples[i].size);
    }
    const stsz = makeFullBox('stsz', 0, 0, stszPayload);

    // stsc (sample-to-chunk) - all samples in one chunk
    const stscPayload = new Uint8Array(4 + 12);
    writeU32(stscPayload, 0, 1); // entry count
    writeU32(stscPayload, 4, 1); // first chunk
    writeU32(stscPayload, 8, samples.length); // samples per chunk
    writeU32(stscPayload, 12, 1); // sample description index
    const stsc = makeFullBox('stsc', 0, 0, stscPayload);

    // co64 (chunk offset - 64-bit since file can be > 4GB)
    // We'll fill this in after calculating the moov size
    const co64Payload = new Uint8Array(4 + 8);
    writeU32(co64Payload, 0, 1); // entry count
    // offset will be set later
    const co64 = makeFullBox('co64', 0, 0, co64Payload);

    // Build stbl
    const newStbl = makeBox('stbl', stsdData, stts, stsc, stsz, co64);

    // Build minf - need vmhd/smhd depending on handler type
    let mediaHeader;
    if (handlerType === 'vide') {
      mediaHeader = makeFullBox('vmhd', 0, 1, new Uint8Array(8)); // graphicsmode + opcolor
    } else {
      mediaHeader = makeFullBox('smhd', 0, 0, new Uint8Array(4)); // balance + reserved
    }
    const dinf = makeBox('dinf', makeFullBox('dref', 0, 0, (() => {
      const d = new Uint8Array(4 + 12);
      writeU32(d, 0, 1); // entry count
      writeU32(d, 4, 12); // entry size
      d[8] = 0x75; d[9] = 0x72; d[10] = 0x6c; d[11] = 0x20; // 'url '
      d[12] = 0; d[13] = 0; d[14] = 0; d[15] = 1; // flags = self-contained
      return d;
    })()));
    const newMinf = makeBox('minf', mediaHeader, dinf, newStbl);

    // Build mdia
    const totalDuration = samples.reduce((sum, s) => sum + s.duration, 0);
    // Update mdhd with correct duration
    let newMdhd;
    if (mdhdData) {
      newMdhd = new Uint8Array(mdhdData);
      const ver = newMdhd[8];
      if (ver === 0) {
        writeU32(newMdhd, 24, totalDuration); // duration
      } else {
        // 64-bit duration at offset 32
        writeU32(newMdhd, 32, 0);
        writeU32(newMdhd, 36, totalDuration);
      }
    } else {
      const mdhdPayload = new Uint8Array(20);
      writeU32(mdhdPayload, 0, 0); // creation time
      writeU32(mdhdPayload, 4, 0); // modification time
      writeU32(mdhdPayload, 8, timescale);
      writeU32(mdhdPayload, 12, totalDuration);
      writeU16(mdhdPayload, 16, 0x55C4); // language (und)
      newMdhd = makeFullBox('mdhd', 0, 0, mdhdPayload);
    }
    const newHdlr = hdlrData || makeFullBox('hdlr', 0, 0, (() => {
      const d = new Uint8Array(21);
      d[4] = 0x76; d[5] = 0x69; d[6] = 0x64; d[7] = 0x65; // 'vide'
      return d;
    })());
    const newMdia = makeBox('mdia', newMdhd, newHdlr, newMinf);

    // Build trak
    let newTkhd;
    if (tkhdData) {
      newTkhd = new Uint8Array(tkhdData);
      // Update duration in tkhd (in movie timescale, not media timescale)
      const ver = newTkhd[8];
      const movieDuration = Math.round(totalDuration / timescale * 1000); // assuming movie timescale = 1000
      if (ver === 0) {
        writeU32(newTkhd, 28, movieDuration);
      } else {
        writeU32(newTkhd, 36, 0);
        writeU32(newTkhd, 40, movieDuration);
      }
    } else {
      const tkhdPayload = new Uint8Array(80);
      writeU32(tkhdPayload, 12, 1); // track ID
      writeU32(tkhdPayload, 20, Math.round(totalDuration / timescale * 1000)); // duration
      newTkhd = makeFullBox('tkhd', 0, 3, tkhdPayload); // flags = enabled + in_movie
    }
    const newTrak = makeBox('trak', newTkhd, newMdia);

    // Build mvhd
    const movieDuration = Math.round(totalDuration / timescale * 1000);
    const mvhdPayload = new Uint8Array(96);
    writeU32(mvhdPayload, 8, 1000); // timescale
    writeU32(mvhdPayload, 12, movieDuration); // duration
    writeU32(mvhdPayload, 16, 0x00010000); // rate = 1.0
    writeU16(mvhdPayload, 20, 0x0100); // volume = 1.0
    // matrix (identity)
    writeU32(mvhdPayload, 32, 0x00010000);
    writeU32(mvhdPayload, 48, 0x00010000);
    writeU32(mvhdPayload, 64, 0x40000000);
    writeU32(mvhdPayload, 92, 2); // next_track_ID
    const mvhd = makeFullBox('mvhd', 0, 0, mvhdPayload);

    // Build moov
    const newMoov = makeBox('moov', mvhd, newTrak);

    // Build mdat
    const mdatHeader = new Uint8Array(8);
    const mdatTotalSize = 8 + totalMediaSize;
    writeU32(mdatHeader, 0, mdatTotalSize);
    mdatHeader[4] = 0x6d; mdatHeader[5] = 0x64; mdatHeader[6] = 0x61; mdatHeader[7] = 0x74; // 'mdat'

    // Calculate total output size and the mdat offset for co64
    const ftypData = data.slice(ftyp.offset, ftyp.offset + ftyp.size);
    const mdatOffset = ftypData.length + newMoov.length + 8; // ftyp + moov + mdat header

    // Fix co64 offset in the built moov
    // Find co64 in our built moov
    for (let i = 0; i < newMoov.length - 16; i++) {
      if (newMoov[i+4] === 0x63 && newMoov[i+5] === 0x6f && newMoov[i+6] === 0x36 && newMoov[i+7] === 0x34) { // 'co64'
        // Write the mdat data offset (after the 8-byte mdat header)
        const dataOffset = mdatOffset;
        writeU32(newMoov, i + 16, 0); // high 32 bits
        writeU32(newMoov, i + 20, dataOffset); // low 32 bits
        break;
      }
    }

    // Step 5: Assemble final output
    const outputSize = ftypData.length + newMoov.length + mdatTotalSize;
    const output = new Uint8Array(outputSize);
    let outPos = 0;

    output.set(ftypData, outPos); outPos += ftypData.length;
    output.set(newMoov, outPos); outPos += newMoov.length;
    output.set(mdatHeader, outPos); outPos += 8;
    for (const chunk of mdatChunks) {
      output.set(chunk, outPos);
      outPos += chunk.length;
    }

    return output;
  };

  /**
   * Extract track info from fMP4 data without building the full output.
   * Returns: { ftyp, initMoov, samples, mdatChunks, totalMediaSize, stsdData, hdlrData, tkhdData, mdhdData, handlerType, timescale }
   */
  const extractTrackInfo = (fmp4Data, timescale, segmentDurationMs) => {
    const data = fmp4Data;
    const total = data.length;
    const ftyp = findBox(data, 'ftyp', 0);
    const initMoov = findBox(data, 'moov', 0);
    if (!ftyp || !initMoov) throw new Error('Missing ftyp or moov');

    const initMoovData = data.slice(initMoov.offset, initMoov.offset + initMoov.size);
    const trak = findBox(initMoovData, 'trak', 8);
    const mdia = trak ? findBoxIn(initMoovData, trak.offset, trak.size, 'mdia') : null;
    const minf = mdia ? findBoxIn(initMoovData, mdia.offset, mdia.size, 'minf') : null;
    const stbl = minf ? findBoxIn(initMoovData, minf.offset, minf.size, 'stbl') : null;
    const stsd = stbl ? findBoxIn(initMoovData, stbl.offset, stbl.size, 'stsd') : null;
    if (!stsd) throw new Error('No stsd found');

    const stsdData = initMoovData.slice(stsd.offset, stsd.offset + stsd.size);
    const hdlr = mdia ? findBoxIn(initMoovData, mdia.offset, mdia.size, 'hdlr') : null;
    let handlerType = 'vide';
    if (hdlr) handlerType = String.fromCharCode(initMoovData[hdlr.offset+16], initMoovData[hdlr.offset+17], initMoovData[hdlr.offset+18], initMoovData[hdlr.offset+19]);
    const hdlrData = hdlr ? initMoovData.slice(hdlr.offset, hdlr.offset + hdlr.size) : null;
    const tkhd = trak ? findBoxIn(initMoovData, trak.offset, trak.size, 'tkhd') : null;
    const tkhdData = tkhd ? initMoovData.slice(tkhd.offset, tkhd.offset + tkhd.size) : null;
    const mdhdBox = mdia ? findBoxIn(initMoovData, mdia.offset, mdia.size, 'mdhd') : null;
    const mdhdData = mdhdBox ? initMoovData.slice(mdhdBox.offset, mdhdBox.offset + mdhdBox.size) : null;

    const segDurationTs = Math.round(segmentDurationMs / 1000 * timescale);
    const samples = [];
    const mdatChunks = [];
    let totalMediaSize = 0;
    let pos = initMoov.offset + initMoov.size;

    while (pos < total - 8) {
      const size = readU32(data, pos);
      const type = boxType(data, pos);
      if (size < 8) break;
      if (type === 'moof') {
        let inner = pos + 8;
        while (inner < pos + size - 8) {
          const iSize = readU32(data, inner);
          if (iSize < 8) break;
          if (boxType(data, inner) === 'traf') {
            let tp = inner + 8;
            while (tp < inner + iSize - 8) {
              const tSize = readU32(data, tp);
              if (tSize < 8) break;
              if (boxType(data, tp) === 'trun') {
                const fl = (data[tp+9]<<16)|(data[tp+10]<<8)|data[tp+11];
                const sc = readU32(data, tp+12);
                const perSampleDur = sc > 0 ? Math.round(segDurationTs / sc) : 1;
                let o = 16;
                if (fl & 0x1) o += 4;
                if (fl & 0x4) o += 4;
                for (let s = 0; s < sc; s++) {
                  if (fl & 0x100) o += 4;
                  const sz = (fl & 0x200) ? readU32(data, tp + o) : 0;
                  if (fl & 0x200) o += 4;
                  if (fl & 0x400) o += 4;
                  if (fl & 0x800) o += 4;
                  samples.push({ size: sz, duration: perSampleDur });
                }
              }
              tp += tSize;
            }
          }
          inner += iSize;
        }
      }
      if (type === 'mdat') {
        mdatChunks.push(data.slice(pos + 8, pos + size));
        totalMediaSize += size - 8;
      }
      pos += size;
    }

    return {
      ftypData: data.slice(ftyp.offset, ftyp.offset + ftyp.size),
      samples, mdatChunks, totalMediaSize,
      stsdData, hdlrData, tkhdData, mdhdData, handlerType, timescale
    };
  };

  /**
   * Build a trak box from track info.
   */
  const buildTrak = (trackId, info, mdatDataOffset) => {
    const { samples, stsdData, hdlrData, tkhdData, mdhdData, handlerType, timescale } = info;
    const totalDuration = samples.reduce((sum, s) => sum + s.duration, 0);

    // stts
    const sttsRuns = [];
    let curDur = samples[0]?.duration || 1, curCount = 1;
    for (let i = 1; i < samples.length; i++) {
      if (samples[i].duration === curDur) curCount++;
      else { sttsRuns.push({ count: curCount, duration: curDur }); curDur = samples[i].duration; curCount = 1; }
    }
    if (samples.length > 0) sttsRuns.push({ count: curCount, duration: curDur });
    const sttsP = new Uint8Array(4 + sttsRuns.length * 8);
    writeU32(sttsP, 0, sttsRuns.length);
    for (let i = 0; i < sttsRuns.length; i++) { writeU32(sttsP, 4+i*8, sttsRuns[i].count); writeU32(sttsP, 8+i*8, sttsRuns[i].duration); }
    const stts = makeFullBox('stts', 0, 0, sttsP);

    // stsz
    const stszP = new Uint8Array(8 + samples.length * 4);
    writeU32(stszP, 4, samples.length);
    for (let i = 0; i < samples.length; i++) writeU32(stszP, 8+i*4, samples[i].size);
    const stsz = makeFullBox('stsz', 0, 0, stszP);

    // stsc
    const stscP = new Uint8Array(16);
    writeU32(stscP, 0, 1);
    writeU32(stscP, 4, 1);
    writeU32(stscP, 8, samples.length);
    writeU32(stscP, 12, 1);
    const stsc = makeFullBox('stsc', 0, 0, stscP);

    // co64 (placeholder - will be fixed later)
    const co64P = new Uint8Array(12);
    writeU32(co64P, 0, 1);
    writeU32(co64P, 8, mdatDataOffset); // low 32 bits
    const co64 = makeFullBox('co64', 0, 0, co64P);

    const stbl = makeBox('stbl', stsdData, stts, stsc, stsz, co64);
    const mh = handlerType === 'vide'
      ? makeFullBox('vmhd', 0, 1, new Uint8Array(8))
      : makeFullBox('smhd', 0, 0, new Uint8Array(4));
    const drefEntry = new Uint8Array(16);
    writeU32(drefEntry, 0, 1); writeU32(drefEntry, 4, 12);
    drefEntry[8]=0x75;drefEntry[9]=0x72;drefEntry[10]=0x6c;drefEntry[11]=0x20;drefEntry[15]=1;
    const dinf = makeBox('dinf', makeFullBox('dref', 0, 0, drefEntry));
    const minf = makeBox('minf', mh, dinf, stbl);

    let mdhd;
    if (mdhdData) {
      mdhd = new Uint8Array(mdhdData);
      const ver = mdhd[8];
      if (ver === 0) writeU32(mdhd, 24, totalDuration);
      else { writeU32(mdhd, 32, 0); writeU32(mdhd, 36, totalDuration); }
    } else {
      const p = new Uint8Array(20);
      writeU32(p, 8, timescale); writeU32(p, 12, totalDuration); writeU16(p, 16, 0x55C4);
      mdhd = makeFullBox('mdhd', 0, 0, p);
    }
    const hdlr = hdlrData || makeFullBox('hdlr', 0, 0, new Uint8Array(21));
    const mdia = makeBox('mdia', mdhd, hdlr, minf);

    let tkhd;
    if (tkhdData) {
      tkhd = new Uint8Array(tkhdData);
      // Fix track ID
      const ver = tkhd[8];
      if (ver === 0) { writeU32(tkhd, 20, trackId); writeU32(tkhd, 28, Math.round(totalDuration/timescale*1000)); }
      else { writeU32(tkhd, 28, trackId); writeU32(tkhd, 36, 0); writeU32(tkhd, 40, Math.round(totalDuration/timescale*1000)); }
    } else {
      const p = new Uint8Array(80);
      writeU32(p, 12, trackId);
      writeU32(p, 20, Math.round(totalDuration/timescale*1000));
      tkhd = makeFullBox('tkhd', 0, 3, p);
    }

    return { trak: makeBox('trak', tkhd, mdia), totalDuration, timescale };
  };

  /**
   * Convert and merge video + audio fMP4 into a single standard MP4.
   * @param {Uint8Array} videoData - Video fMP4 data (with timestamps already fixed)
   * @param {number} videoTimescale
   * @param {Uint8Array} audioData - Audio fMP4 data (with timestamps already fixed)
   * @param {number} audioTimescale
   * @param {number} segmentDurationMs
   * @returns {Uint8Array} Combined standard MP4
   */
  const convertAndMerge = (videoData, videoTimescale, audioData, audioTimescale, segmentDurationMs) => {
    const vInfo = extractTrackInfo(videoData, videoTimescale, segmentDurationMs);
    const aInfo = extractTrackInfo(audioData, audioTimescale, segmentDurationMs);

    // Combine mdat from both tracks
    const totalMdatSize = 8 + vInfo.totalMediaSize + aInfo.totalMediaSize;
    const videoMdatOffset = 0; // placeholder, fixed after moov is built
    const audioMdatOffset = 0;

    // Build tracks (offsets are placeholders)
    const vTrack = buildTrak(1, vInfo, 0);
    const aTrack = buildTrak(2, aInfo, 0);

    // Build mvhd
    const movieDuration = Math.max(
      Math.round(vTrack.totalDuration / vTrack.timescale * 1000),
      Math.round(aTrack.totalDuration / aTrack.timescale * 1000)
    );
    const mvhdP = new Uint8Array(96);
    writeU32(mvhdP, 8, 1000);
    writeU32(mvhdP, 12, movieDuration);
    writeU32(mvhdP, 16, 0x00010000);
    writeU16(mvhdP, 20, 0x0100);
    writeU32(mvhdP, 32, 0x00010000);
    writeU32(mvhdP, 48, 0x00010000);
    writeU32(mvhdP, 64, 0x40000000);
    writeU32(mvhdP, 92, 3);
    const mvhd = makeFullBox('mvhd', 0, 0, mvhdP);

    const moov = makeBox('moov', mvhd, vTrack.trak, aTrack.trak);

    // Calculate mdat offsets
    const ftypData = vInfo.ftypData;
    const mdatHeaderSize = 8;
    const mdatStart = ftypData.length + moov.length + mdatHeaderSize;
    const audioDataStart = mdatStart + vInfo.totalMediaSize;

    // Fix co64 offsets in moov — find both co64 boxes
    let co64Count = 0;
    for (let i = 0; i < moov.length - 16; i++) {
      if (moov[i+4]===0x63 && moov[i+5]===0x6f && moov[i+6]===0x36 && moov[i+7]===0x34) {
        co64Count++;
        const offset = co64Count === 1 ? mdatStart : audioDataStart;
        writeU32(moov, i + 16, 0);
        writeU32(moov, i + 20, offset);
      }
    }

    // Assemble output
    const mdatHeader = new Uint8Array(8);
    writeU32(mdatHeader, 0, totalMdatSize);
    mdatHeader[4]=0x6d;mdatHeader[5]=0x64;mdatHeader[6]=0x61;mdatHeader[7]=0x74;

    const outputSize = ftypData.length + moov.length + totalMdatSize;
    const output = new Uint8Array(outputSize);
    let p = 0;
    output.set(ftypData, p); p += ftypData.length;
    output.set(moov, p); p += moov.length;
    output.set(mdatHeader, p); p += 8;
    for (const c of vInfo.mdatChunks) { output.set(c, p); p += c.length; }
    for (const c of aInfo.mdatChunks) { output.set(c, p); p += c.length; }

    return output;
  };

  // Expose
  window.__fmp4ToMp4 = { convert, convertAndMerge };

  console.log('[Teams Chat Exporter] fMP4 to MP4 converter loaded');
})();
