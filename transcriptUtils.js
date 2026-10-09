/** Shared by the isolated content script and page-world transcript scripts. */
(() => {
  const decode = (s) => s.replace(/&(amp|lt|gt|nbsp);/g, (_, key) =>
    ({ amp: '&', lt: '<', gt: '>', nbsp: ' ' })[key]);
  const escape = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const blocks = (vtt) => String(vtt || '').replace(/^\uFEFF/, '').split(/\r?\n\s*\r?\n/);
  const cue = (block) => {
    const lines = block.split(/\r?\n/);
    if (/^(WEBVTT|NOTE|STYLE|REGION)(?:\s|$)/.test(lines[0])) return null;
    const timingIndex = lines[0]?.includes('-->') ? 0 : 1;
    if (!/^\d{2,}:\d{2}(?::\d{2})?\.\d{3}\s+-->/.test(lines[timingIndex] || '')) return null;
    return { lines, timingIndex, id: timingIndex ? lines[0] : '', payload: lines.slice(timingIndex + 1).join('\n') };
  };
  const voice = (payload) => payload.match(/<v(?:\.[^\s>]+)*\s+([^>]+)>/);
  const speakerMap = (vtt) => {
    const result = Object.create(null);
    for (const block of blocks(vtt)) {
      const c = cue(block);
      const name = c && voice(c.payload);
      if (c?.id && name && !/^Unknown(?: speaker)?$/i.test(name[1])) result[c.id] = decode(name[1]);
    }
    return result;
  };
  // Stream splits entry GUID/9 into GUID/9-0, GUID/9-1, etc. Match only
  // these exact IDs; timestamps and overlapping speaker ranges are ambiguous.
  const enrichVtt = (vtt, speakers = {}) => String(vtt || '').split(/(\r?\n\s*\r?\n)/).map((block, i) => {
    if (i % 2) return block;
    const c = cue(block);
    if (!c || voice(c.payload)) return block;
    const parentId = c.id.replace(/(\/\d+)-\d+$/, '$1');
    const name = Object.hasOwn(speakers, c.id) ? speakers[c.id]
      : Object.hasOwn(speakers, parentId) ? speakers[parentId] : null;
    if (typeof name !== 'string' || !name.trim()) return block;
    const eol = block.includes('\r\n') ? '\r\n' : '\n';
    const lines = block.split(eol);
    lines[c.timingIndex + 1] = `<v ${escape(name).replace(/[\r\n]+/g, ' ')}>${lines[c.timingIndex + 1]}`;
    // Close before trailing blank lines, preserving the original newline style.
    let end = lines.length - 1;
    while (end > c.timingIndex && !lines[end]) end--;
    lines[end] += '</v>';
    return lines.join(eol);
  }).join('');
  const convertVttToTxt = (vtt) => blocks(vtt).flatMap((block) => {
    const c = cue(block);
    if (!c) return [];
    const name = voice(c.payload);
    const text = decode(c.payload.replace(/<[^>]*>/g, '')).trim();
    const timestamp = c.lines[c.timingIndex].split('-->')[0].trim().split('.')[0];
    return text ? [`[${timestamp}] ${name ? decode(name[1]) + ': ' : ''}${text}`] : [];
  }).join('\n');
  const speakerCoverage = (vtt) => {
    const cues = blocks(vtt).map(cue).filter(Boolean);
    return { total: cues.length, named: cues.filter(c => voice(c.payload)).length };
  };
  globalThis.__tceTranscript = { speakerMap, enrichVtt, convertVttToTxt, speakerCoverage };
})();
