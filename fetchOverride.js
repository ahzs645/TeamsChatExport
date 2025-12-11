(() => {
  const { fetch: originalFetch } = window;
  const hiddenDivId = 'transcript-extractor-for-microsoft-stream-hidden-div-with-transcript';

  const ensureHiddenDiv = () => {
    const existing = document.getElementById(hiddenDivId);
    if (existing) {
      return existing;
    }

    const hiddenDiv = document.createElement('div');
    hiddenDiv.style.display = 'none';
    hiddenDiv.id = hiddenDivId;
    document.body.appendChild(hiddenDiv);
    return hiddenDiv;
  };

  const toVttTimestamp = (offset) => {
    if (!offset || typeof offset !== 'string') {
      return '00:00:00.000';
    }

    const [hhmmss = '00:00:00', fraction = '000'] = offset.split('.');
    const millis = (fraction + '000').slice(0, 3);
    return `${hhmmss}.${millis}`;
  };

  const buildVttTranscript = (entries) => {
    if (!Array.isArray(entries) || entries.length === 0) {
      return 'WEBVTT\n\nNOTE Transcript not available.';
    }

    const sorted = [...entries].sort((a, b) => {
      const aStart = a.startOffset || '';
      const bStart = b.startOffset || '';
      return aStart.localeCompare(bStart);
    });

    const cues = sorted.map((entry, index) => {
      const start = toVttTimestamp(entry.startOffset);
      const end = toVttTimestamp(entry.endOffset);
      const speaker = entry.speakerDisplayName || entry.speakerId || 'Unknown speaker';
      const text = entry.text || '';
      const cueId = entry.id || `${index + 1}`;

      return `${cueId}\n${start} --> ${end}\n<v ${speaker}>${text}</v>`;
    });

    return ['WEBVTT', ...cues].join('\n\n');
  };

  window.fetch = async (...args) => {
    const response = await originalFetch(...args);
    const [resource] = args;
    const url = typeof resource === 'string' ? resource : resource?.url || '';

    if (url.includes('streamContent')) {
      const clone = response.clone();

      clone
        .json()
        .then((data) => {
          const transcriptText = buildVttTranscript(data.entries);
          const hiddenDiv = ensureHiddenDiv();
          hiddenDiv.textContent = transcriptText;
        })
        .catch((err) => console.error('[Teams Transcript] Failed to parse transcript JSON', err));
    }

    return response;
  };
})();
