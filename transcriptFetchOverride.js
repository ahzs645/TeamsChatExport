/**
 * Transcript Fetch Override
 * Captures transcript data from Microsoft Stream/Teams recordings via:
 * 1. DOM extraction (primary method - for newer Stream/SharePoint)
 * 2. Fetch interception (fallback - for older implementations)
 */
(() => {
	const { fetch: originalFetch } = window;
	const hiddenDivId = 'teams-chat-exporter-transcript-data';
	let domExtractionInterval = null;
	let lastExtractedCount = 0;

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

	// Convert "X minutes Y seconds" or "Y seconds" to HH:MM:SS format
	const parseTimestampFromLabel = (label) => {
		if (!label) return '00:00:00';

		let hours = 0, minutes = 0, seconds = 0;

		const hourMatch = label.match(/(\d+)\s*hours?/i);
		const minMatch = label.match(/(\d+)\s*minutes?/i);
		const secMatch = label.match(/(\d+)\s*seconds?/i);

		if (hourMatch) hours = parseInt(hourMatch[1], 10);
		if (minMatch) minutes = parseInt(minMatch[1], 10);
		if (secMatch) seconds = parseInt(secMatch[1], 10);

		return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
	};

	// Convert HH:MM:SS to total seconds for sorting
	const timestampToSeconds = (ts) => {
		const parts = ts.split(':').map(Number);
		return (parts[0] || 0) * 3600 + (parts[1] || 0) * 60 + (parts[2] || 0);
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

	// Convert ISO 8601 duration (PT1H23M17S) to HH:MM:SS
	const parsePTDuration = (pt) => {
		if (!pt || typeof pt !== 'string') return '00:00:00';
		const h = pt.match(/(\d+)H/i);
		const m = pt.match(/(\d+)M/i);
		const s = pt.match(/(\d+)S/i);
		const hours = h ? parseInt(h[1], 10) : 0;
		const mins = m ? parseInt(m[1], 10) : 0;
		const secs = s ? parseInt(s[1], 10) : 0;
		return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
	};

	// Extract complete transcript from React fiber state (all entries, not just visible)
	const extractFromReactState = () => {
		// Try multiple starting elements - the data lives deep in the tree
		// and the scroll container class name changes between page loads
		const startElements = [
			document.querySelector('.ms-List'),
			document.querySelector('[role="list"]'),
			document.querySelector('[role="group"]'),
			document.querySelector('.focusZoneWithAutoScroll-309'),
			document.querySelector('[aria-label="Transcript"]'),
			...Array.from(document.querySelectorAll('[class*="focusZone" i]')).slice(0, 3)
		].filter(Boolean);

		const visited = new WeakSet();

		for (const startEl of startElements) {
			const fiberKey = Object.keys(startEl).find(k => k.startsWith('__reactFiber'));
			if (!fiberKey) continue;

			let fiber = startEl[fiberKey];
			let depth = 0;
			while (fiber && depth < 50) {
				if (visited.has(fiber)) { fiber = fiber.return; depth++; continue; }
				visited.add(fiber);

				let hookState = fiber.memoizedState;
				let hookIdx = 0;
				while (hookState && hookIdx < 30) {
					const mem = hookState.memoizedState;
					if (Array.isArray(mem) && mem.length > 0) {
						const sample = mem[0];
						if (sample && sample.speakerDisplayName !== undefined && sample.text !== undefined && sample.timestamp !== undefined) {
							return mem.map((entry, index) => ({
								startOffset: parsePTDuration(entry.timestamp),
								endOffset: parsePTDuration(entry.timestamp),
								speakerDisplayName: entry.speakerDisplayName || 'Unknown',
								text: entry.text || '',
								id: entry.entryId || `react-${index + 1}`
							})).filter(e => e.text.trim());
						}
					}
					hookState = hookState.next;
					hookIdx++;
				}
				fiber = fiber.return;
				depth++;
			}
		}
		return null;
	};

	// Extract transcript entries from DOM (fallback for visible entries only)
	const extractFromDOM = () => {
		// Try React state first for complete transcript
		const reactEntries = extractFromReactState();
		if (reactEntries && reactEntries.length > 0) {
			console.log(`[Teams Chat Exporter] Extracted ${reactEntries.length} entries from React state (complete)`);
			return reactEntries;
		}

		// Fallback: DOM extraction (visible entries only)
		const entries = [];
		const groups = document.querySelectorAll('[role="group"]');

		groups.forEach((group, index) => {
			const listitem = group.querySelector('[role="listitem"]');

			if (listitem && listitem.textContent.trim()) {
				const groupLabel = group.getAttribute('aria-label') || '';

				// Format: "Speaker Name HH minutes SS seconds" or "Speaker Name SS seconds"
				const match = groupLabel.match(/^(.*?)\s*(\d+\s+hours?\s+\d+\s+minutes?\s+\d+\s+seconds?|\d+\s+minutes?\s+\d+\s+seconds?|\d+\s+seconds?)$/i);

				let speaker = 'Unknown';
				let timestamp = '00:00:00';

				if (match) {
					speaker = match[1].trim() || 'Unknown';
					timestamp = parseTimestampFromLabel(match[2]);
				}

				entries.push({
					startOffset: timestamp,
					endOffset: timestamp, // DOM doesn't provide end time
					speakerDisplayName: speaker,
					text: listitem.textContent.trim(),
					id: `dom-${index + 1}`
				});
			}
		});

		return entries;
	};

	const buildVttTranscript = (entries) => {
		if (!Array.isArray(entries) || entries.length === 0) {
			return 'WEBVTT\n\nNOTE Transcript not available.';
		}

		const sorted = [...entries].sort((a, b) => {
			const aSeconds = timestampToSeconds(a.startOffset || '00:00:00');
			const bSeconds = timestampToSeconds(b.startOffset || '00:00:00');
			return aSeconds - bSeconds;
		});

		const cues = sorted.map((entry, index) => {
			const startSeconds = timestampToSeconds(entry.startOffset || '00:00:00');
			// Estimate end time as start + 5 seconds if not provided or same as start
			let endSeconds = timestampToSeconds(entry.endOffset || entry.startOffset || '00:00:00');
			if (endSeconds <= startSeconds) {
				endSeconds = startSeconds + 5;
			}

			// Format start timestamp
			const startHours = Math.floor(startSeconds / 3600);
			const startMins = Math.floor((startSeconds % 3600) / 60);
			const startSecs = startSeconds % 60;
			const start = `${String(startHours).padStart(2, '0')}:${String(startMins).padStart(2, '0')}:${String(startSecs).padStart(2, '0')}.000`;

			// Format end timestamp
			const endHours = Math.floor(endSeconds / 3600);
			const endMins = Math.floor((endSeconds % 3600) / 60);
			const endSecs = endSeconds % 60;
			const end = `${String(endHours).padStart(2, '0')}:${String(endMins).padStart(2, '0')}:${String(endSecs).padStart(2, '0')}.000`;

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
			const aSeconds = timestampToSeconds(a.startOffset || '00:00:00');
			const bSeconds = timestampToSeconds(b.startOffset || '00:00:00');
			return aSeconds - bSeconds;
		});

		const lines = sorted.map((entry) => {
			const timestamp = toSimpleTimestamp(entry.startOffset);
			const speaker = entry.speakerDisplayName || entry.speakerId || 'Unknown';
			const text = entry.text || '';

			return `[${timestamp}] ${speaker}: ${text}`;
		});

		return lines.join('\n');
	};

	const storeTranscript = (entries, source) => {
		if (!entries || entries.length === 0) return false;

		const vttTranscript = buildVttTranscript(entries);
		const txtTranscript = buildTxtTranscript(entries);
		const hiddenDiv = ensureHiddenDiv();

		hiddenDiv.setAttribute('data-vtt', vttTranscript);
		hiddenDiv.setAttribute('data-txt', txtTranscript);
		hiddenDiv.setAttribute('data-source', source);
		hiddenDiv.setAttribute('data-count', entries.length.toString());
		hiddenDiv.textContent = vttTranscript;

		console.log(`[Teams Chat Exporter] Transcript captured via ${source}: ${entries.length} entries`);
		return true;
	};

	// DOM extraction polling - checks periodically for transcript content
	const startDOMExtraction = () => {
		if (domExtractionInterval) return;

		const checkDOM = () => {
			const entries = extractFromDOM();

			// Only update if we found entries and count changed (transcript may load progressively)
			if (entries.length > 0 && entries.length !== lastExtractedCount) {
				lastExtractedCount = entries.length;
				storeTranscript(entries, 'DOM');
			}
		};

		// Check immediately
		checkDOM();

		// Then poll every 2 seconds for new content (transcript may load as video plays)
		domExtractionInterval = setInterval(checkDOM, 2000);

		// Also observe DOM changes for transcript panel
		const observer = new MutationObserver(() => {
			checkDOM();
		});

		// Start observing once body is ready
		if (document.body) {
			observer.observe(document.body, { childList: true, subtree: true });
		} else {
			document.addEventListener('DOMContentLoaded', () => {
				observer.observe(document.body, { childList: true, subtree: true });
			});
		}

		console.log('[Teams Chat Exporter] DOM extraction started');
	};

	// Fetch interception (fallback for older implementations)
	window.fetch = async (...args) => {
		const response = await originalFetch(...args);
		const [resource] = args;
		const url = typeof resource === 'string' ? resource : resource?.url || '';

		// Check for transcript-related URLs
		if (url.includes('streamContent') || url.includes('transcript') || url.includes('captions')) {
			const clone = response.clone();

			clone.json()
				.then((data) => {
					if (data.entries && Array.isArray(data.entries)) {
						storeTranscript(data.entries, 'fetch');
					}
				})
				.catch(() => {}); // Silently fail if not JSON
		}

		return response;
	};

	// Start DOM extraction immediately
	startDOMExtraction();

	// Expose utilities for batch transcript download
	window.__teamsTranscriptUtils = {
		extractFromDOM,
		extractFromReactState,
		buildVttTranscript,
		buildTxtTranscript,
		parseTimestampFromLabel,
		parsePTDuration,
		timestampToSeconds,
		toSimpleTimestamp,
		storeTranscript,
		ensureHiddenDiv
	};

	console.log('[Teams Chat Exporter] Transcript capture initialized (DOM + fetch)');
})();
