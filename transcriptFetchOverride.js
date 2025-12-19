/**
 * Transcript Fetch Override
 * Intercepts fetch calls to capture transcript data from Microsoft Stream/Teams recordings
 */
(() => {
	const { fetch: originalFetch } = window;
	const hiddenDivId = 'teams-chat-exporter-transcript-data';

	const ensureHiddenDiv = () => {
		const existing = document.getElementById(hiddenDivId);
		if (existing) {
			return existing;
		}

		const hiddenDiv = document.createElement('div');
		hiddenDiv.style.display = 'none';
		hiddenDiv.id = hiddenDivId;
		window.document.body.appendChild(hiddenDiv);
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

	const toSimpleTimestamp = (offset) => {
		if (!offset || typeof offset !== 'string') {
			return '00:00:00';
		}
		const [hhmmss = '00:00:00'] = offset.split('.');
		return hhmmss;
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

	const buildTxtTranscript = (entries) => {
		if (!Array.isArray(entries) || entries.length === 0) {
			return 'Transcript not available.';
		}

		const sorted = [...entries].sort((a, b) => {
			const aStart = a.startOffset || '';
			const bStart = b.startOffset || '';
			return aStart.localeCompare(bStart);
		});

		const lines = sorted.map((entry) => {
			const timestamp = toSimpleTimestamp(entry.startOffset);
			const speaker = entry.speakerDisplayName || entry.speakerId || 'Unknown';
			const text = entry.text || '';

			return `[${timestamp}] ${speaker}: ${text}`;
		});

		return lines.join('\n');
	};

	window.fetch = async (...args) => {
		const response = await originalFetch(...args);
		const [resource] = args;
		const url = typeof resource === 'string' ? resource : resource?.url || '';

		if (url.includes('streamContent')) {
			const clone = response.clone();

			clone.json()
				.then((data) => {
					const vttTranscript = buildVttTranscript(data.entries);
					const txtTranscript = buildTxtTranscript(data.entries);
					const hiddenDiv = ensureHiddenDiv();

					// Store both formats as JSON
					hiddenDiv.setAttribute('data-vtt', vttTranscript);
					hiddenDiv.setAttribute('data-txt', txtTranscript);
					hiddenDiv.textContent = vttTranscript; // Default to VTT for backwards compatibility

					console.log('[Teams Chat Exporter] Transcript captured successfully');
				})
				.catch((err) => console.error('[Teams Chat Exporter] Error parsing transcript:', err));
		}

		return response;
	};

	console.log('[Teams Chat Exporter] Transcript fetch override installed');
})();
