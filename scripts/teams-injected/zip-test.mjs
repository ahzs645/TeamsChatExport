// Unit check for src/modules/zip.js: writes a zip, then validate with
//   unzip -t <out> && python3 -I -m zipfile -t <out>
// Usage: node scripts/teams-injected/zip-test.mjs <out.zip>
import { createZip, crc32 } from '../../src/modules/zip.js';
import fs from 'fs';

const out = process.argv[2] || 'test.zip';
if (crc32(new TextEncoder().encode('123456789')) !== 0xcbf43926) throw new Error('crc32 check value mismatch');
const big = 'WEBVTT\n\n' + Array.from({ length: 2000 }, (_, i) => `${i + 1}\n00:00:${String(i % 60).padStart(2, '0')}.000 --> 00:00:01.000\n<v Speaker>Line ${i}</v>`).join('\n\n');
const zip = createZip([
  { name: 'Weekly sync - 2026-10-01.vtt', data: big },
  { name: '週次ミーティング - 10月1日.vtt', data: 'WEBVTT\n\n1\n00:00:00.000 --> 00:00:01.000\n<v 田中>こんにちは</v>\n' },
  { name: 'Weekly sync - 2026-10-01.vtt', data: 'duplicate name' },
  { name: 'empty.txt', data: '' },
  { name: 'bytes.bin', data: new Uint8Array([0, 1, 2, 255]) }
]);
fs.writeFileSync(out, zip);
console.log(`wrote ${out} (${zip.length} bytes)`);
