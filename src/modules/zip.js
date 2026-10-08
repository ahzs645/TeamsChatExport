/**
 * Minimal dependency-free ZIP writer (STORE method, no compression).
 *
 * Used to bundle batch transcript downloads into a single file instead of
 * firing N separate downloads. Text compresses well, but transcripts are small
 * and STORE keeps this tiny and obviously correct: local file header + data
 * per entry, then a central directory and end-of-central-directory record.
 * File names are written as UTF-8 (general purpose flag bit 11) so non-ASCII
 * titles survive. No ZIP64: fine for transcript-sized archives (< 4 GiB).
 *
 * Loaded by content.js via dynamic import (src/modules/* is web-accessible),
 * and importable from Node for testing.
 */

let crcTable = null;
const getCrcTable = () => {
  if (crcTable) return crcTable;
  crcTable = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[n] = c >>> 0;
  }
  return crcTable;
};

export const crc32 = (bytes) => {
  const table = getCrcTable();
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) crc = table[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
};

const toDosDateTime = (date) => {
  const d = date instanceof Date && !isNaN(date) ? date : new Date();
  const year = Math.max(1980, d.getFullYear());
  return {
    time: ((d.getHours() & 0x1f) << 11) | ((d.getMinutes() & 0x3f) << 5) | ((d.getSeconds() >> 1) & 0x1f),
    date: (((year - 1980) & 0x7f) << 9) | (((d.getMonth() + 1) & 0x0f) << 5) | (d.getDate() & 0x1f)
  };
};

/**
 * Make names unique within the archive: "a.vtt", "a (2).vtt", ...
 */
const uniqueName = (name, used) => {
  let candidate = name;
  let n = 2;
  while (used.has(candidate.toLowerCase())) {
    const dot = name.lastIndexOf('.');
    candidate = dot > 0 ? `${name.slice(0, dot)} (${n})${name.slice(dot)}` : `${name} (${n})`;
    n++;
  }
  used.add(candidate.toLowerCase());
  return candidate;
};

/**
 * @param {Array<{name: string, data: string|Uint8Array|ArrayBuffer, date?: Date}>} files
 * @returns {Uint8Array} the complete .zip file
 */
export const createZip = (files) => {
  const enc = new TextEncoder();
  const used = new Set();
  const entries = [];
  let offset = 0;

  for (const file of files) {
    const name = uniqueName(String(file.name || 'file').replace(/\\/g, '/').replace(/^\/+/, ''), used);
    const nameBytes = enc.encode(name);
    const data = typeof file.data === 'string'
      ? enc.encode(file.data)
      : file.data instanceof Uint8Array ? file.data : new Uint8Array(file.data || []);
    const crc = crc32(data);
    const { time, date } = toDosDateTime(file.date);

    const header = new Uint8Array(30 + nameBytes.length);
    const hv = new DataView(header.buffer);
    hv.setUint32(0, 0x04034b50, true); // local file header signature
    hv.setUint16(4, 20, true);         // version needed to extract (2.0)
    hv.setUint16(6, 0x0800, true);     // flags: UTF-8 names
    hv.setUint16(8, 0, true);          // method: STORE
    hv.setUint16(10, time, true);
    hv.setUint16(12, date, true);
    hv.setUint32(14, crc, true);
    hv.setUint32(18, data.length, true); // compressed size
    hv.setUint32(22, data.length, true); // uncompressed size
    hv.setUint16(26, nameBytes.length, true);
    hv.setUint16(28, 0, true);           // extra length
    header.set(nameBytes, 30);

    entries.push({ nameBytes, data, crc, time, date, offset, header });
    offset += header.length + data.length;
  }

  const centralStart = offset;
  const centrals = entries.map((e) => {
    const c = new Uint8Array(46 + e.nameBytes.length);
    const cv = new DataView(c.buffer);
    cv.setUint32(0, 0x02014b50, true); // central directory header signature
    cv.setUint16(4, 20, true);         // version made by
    cv.setUint16(6, 20, true);         // version needed
    cv.setUint16(8, 0x0800, true);     // flags: UTF-8
    cv.setUint16(10, 0, true);         // STORE
    cv.setUint16(12, e.time, true);
    cv.setUint16(14, e.date, true);
    cv.setUint32(16, e.crc, true);
    cv.setUint32(20, e.data.length, true);
    cv.setUint32(24, e.data.length, true);
    cv.setUint16(28, e.nameBytes.length, true);
    cv.setUint16(30, 0, true);         // extra
    cv.setUint16(32, 0, true);         // comment
    cv.setUint16(34, 0, true);         // disk number start
    cv.setUint16(36, 0, true);         // internal attrs
    cv.setUint32(38, 0, true);         // external attrs
    cv.setUint32(42, e.offset, true);  // local header offset
    c.set(e.nameBytes, 46);
    return c;
  });
  const centralSize = centrals.reduce((s, c) => s + c.length, 0);

  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, centralStart, true);
  ev.setUint16(20, 0, true);

  const out = new Uint8Array(centralStart + centralSize + eocd.length);
  let p = 0;
  for (const e of entries) {
    out.set(e.header, p); p += e.header.length;
    out.set(e.data, p); p += e.data.length;
  }
  for (const c of centrals) { out.set(c, p); p += c.length; }
  out.set(eocd, p);
  return out;
};

/** Same as createZip but returns a Blob (application/zip). */
export const createZipBlob = (files) => new Blob([createZip(files)], { type: 'application/zip' });
