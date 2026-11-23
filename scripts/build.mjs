import { cp, mkdir, readFile, rename, rm } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const output = path.join(root, 'build');
const staging = path.join(root, '.build-staging');
const next = path.join(staging, 'next');
const previous = path.join(staging, 'previous');
const { version } = JSON.parse(await readFile(path.join(root, 'manifest.json'), 'utf8'));

// The checkout can also contain local HTML files and other unrelated work.
const extensionFiles = [
  'README.md',
  'TROUBLESHOOTING.md',
  'background.js',
  'batchTranscriptDownload.js',
  'chatFetchOverride.js',
  'content.js',
  'contextBridge.js',
  'manifest.json',
  'mp4muxer.js',
  'popup.css',
  'popup.html',
  'popup.js',
  'results.html',
  'results.js',
  'style.css',
  'tokens.css',
  'transcriptAPIFetcher.js',
  'transcriptFetchOverride.js',
  'transcriptStyles.css',
  'transcriptUtils.js',
  'videoDownloadOverride.js',
];
const extensionDirectories = ['assets', 'src', 'ui', 'videoDownload'];

// Reserve the staging directory so two builds cannot replace the output at once.
await mkdir(staging);
let keepPrevious = false;
try {
  await mkdir(next);
  for (const name of extensionFiles) {
    await cp(path.join(root, name), path.join(next, name));
  }
  for (const name of extensionDirectories) {
    await cp(path.join(root, name), path.join(next, name), {
      recursive: true,
      filter: source => !path.basename(source).startsWith('.'),
    });
  }

  let hadPrevious = false;
  try {
    await rename(output, previous);
    hadPrevious = true;
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }

  try {
    await rename(next, output);
  } catch (error) {
    if (hadPrevious) {
      keepPrevious = true;
      await rename(previous, output);
      keepPrevious = false;
    }
    throw error;
  }
} finally {
  if (!keepPrevious) await rm(staging, { recursive: true, force: true });
}

console.log(`Built Teams Chat Exporter v${version} in ${output}`);
