/**
 * Batch Transcript Download
 * Iterates through all meeting instances in the Teams Intelligent Recap dropdown,
 * extracts each transcript via API fetch (preferred) or DOM scraping (fallback).
 *
 * Strategy per meeting:
 *   1. Select the meeting in the dropdown (triggers readcollabobject API call)
 *   2. Wait for transcriptAPIFetcher.js to capture the response metadata
 *   3. If a TranscriptV2 URL + SharePoint token are available, fetch directly via API
 *   4. If API fetch fails or isn't available, fall back to DOM extraction
 *
 * Runs in page context (injected script). Communicates with content script
 * via CustomEvent pattern.
 */
(() => {
	const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

	const DROPDOWN_SELECTOR = '[data-testid="intelligent-recap-instance-select-dropdown"]';
	const OPTION_SELECTOR = '[role="option"]';
	const NO_TRANSCRIPT_SELECTOR = '[data-testid="one-transcript-panel-transcript-status-page"]';
	const LOAD_MORE_TEXT = 'Load more meetings';
	const STABILIZE_INTERVAL = 500;
	const STABILIZE_CHECKS = 2;
	const MEETING_SWITCH_TIMEOUT = 15000;
	const LOAD_MORE_WAIT = 1500;
	const POST_SELECT_WAIT = 1000;
	const API_METADATA_WAIT = 3000; // wait for readcollabobject response to be intercepted
	const API_FETCH_TIMEOUT = 10000;

	let batchRunning = false;
	let batchCancelled = false;

	const dispatchProgress = (data) => {
		document.dispatchEvent(new CustomEvent('teamsBatchTranscriptProgress', { detail: data }));
	};

	// ========== DROPDOWN & NAVIGATION ==========

	const getSeriesName = () => {
		const h2 = document.querySelector('h2');
		return h2?.textContent?.trim() || 'Meeting';
	};

	const openDropdown = async () => {
		const dropdown = document.querySelector(DROPDOWN_SELECTOR);
		if (!dropdown) throw new Error('Meeting dropdown not found');
		dropdown.click();
		for (let i = 0; i < 20; i++) {
			await sleep(200);
			if (document.querySelectorAll(OPTION_SELECTOR).length > 0) return;
		}
		throw new Error('Dropdown options did not appear');
	};

	const closeDropdown = () => {
		const dropdown = document.querySelector(DROPDOWN_SELECTOR);
		if (dropdown) dropdown.click();
	};

	const getVisibleOptions = () => {
		return Array.from(document.querySelectorAll(OPTION_SELECTOR))
			.filter((o) => o.textContent.trim() !== LOAD_MORE_TEXT);
	};

	const getAllMeetingOptions = async () => {
		await openDropdown();
		let loadMoreClicks = 0;
		const maxLoadMore = 50;
		while (loadMoreClicks < maxLoadMore) {
			const options = document.querySelectorAll(OPTION_SELECTOR);
			const loadMore = Array.from(options).find(
				(o) => o.textContent.trim() === LOAD_MORE_TEXT
			);
			if (!loadMore) break;
			loadMore.click();
			loadMoreClicks++;
			await sleep(LOAD_MORE_WAIT);
		}
		const meetings = getVisibleOptions().map((o, i) => ({
			index: i,
			text: o.textContent.trim()
		}));
		closeDropdown();
		await sleep(300);
		return meetings;
	};

	const clickTranscriptTab = async () => {
		const tabs = document.querySelectorAll('[role="tab"]');
		const transcriptTab = Array.from(tabs).find(
			(t) => t.textContent.includes('Transcript')
		);
		if (transcriptTab) {
			transcriptTab.click();
			await sleep(500);
		}
	};

	const getGroupCount = () => {
		return document.querySelectorAll('[role="group"]').length;
	};

	const isNoTranscriptVisible = () => {
		const el = document.querySelector(NO_TRANSCRIPT_SELECTOR);
		if (!el) return false;
		return getComputedStyle(el).display !== 'none';
	};

	const selectMeetingAndWait = async (index) => {
		const baselineCount = getGroupCount();
		await openDropdown();
		const options = getVisibleOptions();
		if (index >= options.length) {
			closeDropdown();
			throw new Error(`Meeting index ${index} out of range (${options.length} options)`);
		}
		options[index].click();
		await sleep(POST_SELECT_WAIT);
		await clickTranscriptTab();

		const startTime = Date.now();
		let droppedToZero = baselineCount === 0;
		let lastCount = -1;
		let consecutiveStable = 0;

		while (Date.now() - startTime < MEETING_SWITCH_TIMEOUT) {
			if (batchCancelled) return false;
			const count = getGroupCount();

			if (!droppedToZero) {
				if (count === 0) droppedToZero = true;
				await sleep(200);
				continue;
			}
			if (isNoTranscriptVisible()) return false;
			if (count > 0) {
				if (count === lastCount) {
					consecutiveStable++;
					if (consecutiveStable >= STABILIZE_CHECKS) return true;
				} else {
					consecutiveStable = 0;
				}
				lastCount = count;
			}
			await sleep(STABILIZE_INTERVAL);
		}
		if (getGroupCount() > 0) return true;
		return false;
	};

	// ========== API METADATA & FETCH ==========

	const getAPIMetadata = () => {
		if (window.__transcriptAPIData && Object.keys(window.__transcriptAPIData).length > 0) {
			return { ...window.__transcriptAPIData };
		}
		const div = document.getElementById('transcript-api-data');
		if (!div) return {};
		try {
			return JSON.parse(div.getAttribute('data-transcript-api') || '{}');
		} catch (e) {
			return {};
		}
	};

	const getSharePointTokens = () => {
		if (window.__sharePointTokens && Object.keys(window.__sharePointTokens).length > 0) {
			return { ...window.__sharePointTokens };
		}
		const div = document.getElementById('transcript-api-data');
		if (!div) return {};
		try {
			return JSON.parse(div.getAttribute('data-sp-tokens') || '{}');
		} catch (e) {
			return {};
		}
	};

	const findMetadataForMeeting = (allMetadata, capturedAfter) => {
		const entries = Object.entries(allMetadata);
		if (entries.length === 0) return null;

		const recent = entries
			.filter(([, v]) => new Date(v.capturedAt).getTime() >= capturedAfter)
			.sort(([, a], [, b]) => new Date(a.capturedAt).getTime() - new Date(b.capturedAt).getTime());

		if (recent.length > 0) {
			return { key: recent[0][0], ...recent[0][1] };
		}
		const sorted = entries.sort(
			([, a], [, b]) => new Date(b.capturedAt).getTime() - new Date(a.capturedAt).getTime()
		);
		return { key: sorted[0][0], ...sorted[0][1] };
	};

	/**
	 * Wait for API metadata to appear after selecting a meeting.
	 * The readcollabobject response is intercepted by transcriptAPIFetcher.js
	 * and stored in window.__transcriptAPIData.
	 */
	const waitForAPIMetadata = async (capturedAfter, timeoutMs = API_METADATA_WAIT) => {
		const start = Date.now();
		while (Date.now() - start < timeoutMs) {
			const allMetadata = getAPIMetadata();
			const meta = findMetadataForMeeting(allMetadata, capturedAfter);
			if (meta) return meta;
			await sleep(300);
		}
		return null;
	};

	/**
	 * Find the SharePoint token for a given transcript URL.
	 */
	const getTokenForUrl = (transcriptUrl) => {
		const tokens = getSharePointTokens();
		if (!transcriptUrl || Object.keys(tokens).length === 0) return null;
		try {
			const host = new URL(transcriptUrl).hostname;
			const tokenEntry = tokens[host];
			if (tokenEntry && tokenEntry.token) {
				// Check if token is not too old (30 minutes)
				if (Date.now() - tokenEntry.capturedAt < 30 * 60 * 1000) {
					return tokenEntry.token;
				}
			}
		} catch (e) {
			// invalid URL
		}
		return null;
	};

	/**
	 * Attempt to fetch transcript content directly via SharePoint API.
	 * Returns parsed transcript result or null if it fails.
	 */
	const fetchTranscriptViaAPI = async (apiMeta) => {
		if (!apiMeta || !apiMeta.resources) return null;

		// Find the TranscriptV2 resource (preferred) or any transcript resource
		const transcriptResource =
			apiMeta.resources['TranscriptV2'] ||
			apiMeta.resources['transcriptV2'] ||
			Object.values(apiMeta.resources).find((r) =>
				r.location && r.location.includes('transcript')
			);

		if (!transcriptResource || !transcriptResource.location) {
			console.log('[Batch Transcript] No transcript URL in API metadata');
			return null;
		}

		const transcriptUrl = transcriptResource.location;
		const token = getTokenForUrl(transcriptUrl);

		if (!token) {
			console.log('[Batch Transcript] No SharePoint token available for', transcriptUrl);
			return null;
		}

		console.log('[Batch Transcript] Attempting API fetch:', transcriptUrl.substring(0, 80) + '...');

		try {
			const controller = new AbortController();
			const timeout = setTimeout(() => controller.abort(), API_FETCH_TIMEOUT);

			const response = await fetch(transcriptUrl, {
				headers: { 'Authorization': token },
				signal: controller.signal
			});
			clearTimeout(timeout);

			if (!response.ok) {
				console.log(`[Batch Transcript] API fetch failed: ${response.status} ${response.statusText}`);
				return null;
			}

			const contentType = response.headers.get('content-type') || '';
			const text = await response.text();

			if (!text || text.length < 10) {
				console.log('[Batch Transcript] API response empty');
				return null;
			}

			// The transcript content could be VTT, JSON, or plain text
			let vtt = '';
			let txt = '';
			let entryCount = 0;

			if (text.startsWith('WEBVTT')) {
				// Already VTT format
				vtt = text;
				txt = convertVttToTxt(text);
				entryCount = (text.match(/-->/g) || []).length;
			} else if (contentType.includes('json') || text.startsWith('{') || text.startsWith('[')) {
				// JSON format — parse entries
				try {
					const data = JSON.parse(text);
					const entries = Array.isArray(data) ? data : (data.entries || data.captions || []);
					if (entries.length > 0) {
						const utils = window.__teamsTranscriptUtils;
						if (utils) {
							vtt = utils.buildVttTranscript(entries);
							txt = utils.buildTxtTranscript(entries);
						} else {
							// Inline build
							txt = entries.map((e) => {
								const ts = e.startOffset || e.offset || '00:00:00';
								const speaker = e.speakerDisplayName || e.speaker || 'Unknown';
								const content = e.text || e.content || '';
								return `[${ts}] ${speaker}: ${content}`;
							}).join('\n');
							vtt = 'WEBVTT\n\n' + entries.map((e, i) => {
								const ts = e.startOffset || e.offset || '00:00:00';
								const speaker = e.speakerDisplayName || e.speaker || 'Unknown';
								const content = e.text || e.content || '';
								return `${i + 1}\n${ts}.000 --> ${ts}.000\n<v ${speaker}>${content}</v>`;
							}).join('\n\n');
						}
						entryCount = entries.length;
					}
				} catch (e) {
					console.log('[Batch Transcript] Failed to parse JSON response');
					return null;
				}
			} else {
				// Plain text — treat as raw transcript
				txt = text;
				vtt = 'WEBVTT\n\nNOTE Raw transcript from API\n\n1\n00:00:00.000 --> 99:59:59.000\n' + text;
				entryCount = text.split('\n').filter((l) => l.trim()).length;
			}

			if (entryCount === 0 && txt.length < 10) return null;

			return {
				hasTranscript: true,
				vtt,
				txt,
				entryCount,
				source: 'api'
			};
		} catch (err) {
			if (err.name === 'AbortError') {
				console.log('[Batch Transcript] API fetch timed out');
			} else {
				console.log('[Batch Transcript] API fetch error:', err.message);
			}
			return null;
		}
	};

	/**
	 * Convert VTT content to plain text format.
	 */
	const convertVttToTxt = (vttContent) => {
		const lines = vttContent.split('\n');
		const txtLines = [];
		let currentTimestamp = '';

		for (const line of lines) {
			if (line.includes('-->')) {
				currentTimestamp = line.split('-->')[0].trim().split('.')[0] || '00:00:00';
			} else if (line.trim() && !line.startsWith('WEBVTT') && !line.startsWith('NOTE') && !line.match(/^\d+$/)) {
				// Extract speaker from <v Speaker>text</v> format
				const vMatch = line.match(/^<v\s+([^>]+)>(.+)<\/v>$/);
				if (vMatch) {
					txtLines.push(`[${currentTimestamp}] ${vMatch[1]}: ${vMatch[2]}`);
				} else if (line.trim()) {
					txtLines.push(`[${currentTimestamp}] ${line.trim()}`);
				}
			}
		}
		return txtLines.join('\n') || vttContent;
	};

	// ========== DOM EXTRACTION ==========

	const extractCurrentTranscript = () => {
		const utils = window.__teamsTranscriptUtils;
		if (utils) {
			const entries = utils.extractFromDOM();
			if (!entries || entries.length === 0) {
				return { hasTranscript: false, vtt: '', txt: '', entryCount: 0, source: 'dom' };
			}
			return {
				hasTranscript: true,
				vtt: utils.buildVttTranscript(entries),
				txt: utils.buildTxtTranscript(entries),
				entryCount: entries.length,
				source: 'dom'
			};
		}

		// Fallback: inline extraction
		const entries = [];
		const groups = document.querySelectorAll('[role="group"]');

		groups.forEach((group, index) => {
			const listitem = group.querySelector('[role="listitem"]');
			if (listitem && listitem.textContent.trim()) {
				const groupLabel = group.getAttribute('aria-label') || '';
				const match = groupLabel.match(
					/^(.*?)\s*(\d+\s+hours?\s+\d+\s+minutes?\s+\d+\s+seconds?|\d+\s+minutes?\s+\d+\s+seconds?|\d+\s+seconds?)$/i
				);

				let speaker = 'Unknown';
				let timestamp = '00:00:00';

				if (match) {
					speaker = match[1].trim() || 'Unknown';
					let h = 0, m = 0, s = 0;
					const hm = groupLabel.match(/(\d+)\s*hours?/i);
					const mm = groupLabel.match(/(\d+)\s*minutes?/i);
					const sm = groupLabel.match(/(\d+)\s*seconds?/i);
					if (hm) h = parseInt(hm[1], 10);
					if (mm) m = parseInt(mm[1], 10);
					if (sm) s = parseInt(sm[1], 10);
					timestamp = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
				}

				entries.push({
					startOffset: timestamp,
					endOffset: timestamp,
					speakerDisplayName: speaker,
					text: listitem.textContent.trim(),
					id: `dom-${index + 1}`
				});
			}
		});

		if (entries.length === 0) {
			return { hasTranscript: false, vtt: '', txt: '', entryCount: 0, source: 'dom' };
		}

		const txtLines = entries.map((e) => `[${e.startOffset}] ${e.speakerDisplayName}: ${e.text}`);
		const vttCues = entries.map((entry) => {
			const parts = entry.startOffset.split(':').map(Number);
			const startSec = (parts[0] || 0) * 3600 + (parts[1] || 0) * 60 + (parts[2] || 0);
			const endSec = startSec + 5;
			const fmt = (sec) => {
				const hh = String(Math.floor(sec / 3600)).padStart(2, '0');
				const mm = String(Math.floor((sec % 3600) / 60)).padStart(2, '0');
				const ss = String(sec % 60).padStart(2, '0');
				return `${hh}:${mm}:${ss}.000`;
			};
			return `${entry.id}\n${fmt(startSec)} --> ${fmt(endSec)}\n<v ${entry.speakerDisplayName}>${entry.text}</v>`;
		});

		return {
			hasTranscript: true,
			vtt: ['WEBVTT', ...vttCues].join('\n\n'),
			txt: txtLines.join('\n'),
			entryCount: entries.length,
			source: 'dom'
		};
	};

	// ========== MAIN ORCHESTRATOR ==========

	const downloadAllTranscripts = async () => {
		if (batchRunning) {
			return { error: 'Batch download already in progress' };
		}

		batchRunning = true;
		batchCancelled = false;
		const results = [];
		let apiSuccessCount = 0;
		let domFallbackCount = 0;

		try {
			const seriesName = getSeriesName();
			dispatchProgress({ phase: 'enumerating', message: 'Loading meeting list...' });

			const meetings = await getAllMeetingOptions();

			if (meetings.length === 0) {
				return { error: 'No meetings found in dropdown' };
			}

			dispatchProgress({
				phase: 'started',
				total: meetings.length,
				current: 0,
				message: `Found ${meetings.length} meetings. Starting extraction...`,
				seriesName
			});

			await clickTranscriptTab();

			for (let i = 0; i < meetings.length; i++) {
				if (batchCancelled) {
					dispatchProgress({ phase: 'cancelled', message: 'Batch cancelled by user' });
					break;
				}

				const meeting = meetings[i];
				dispatchProgress({
					phase: 'processing',
					total: meetings.length,
					current: i + 1,
					currentMeeting: meeting.text,
					message: `Processing ${i + 1}/${meetings.length}: ${meeting.text}`
				});

				const selectTimestamp = Date.now();

				try {
					// Step 1: Select the meeting (this triggers readcollabobject call)
					const hasTranscriptInDOM = await selectMeetingAndWait(meeting.index);

					// Step 2: Wait briefly for API metadata to be captured
					const apiMeta = await waitForAPIMetadata(selectTimestamp);

					// Step 3: Try API fetch first
					let transcript = null;
					if (apiMeta) {
						dispatchProgress({
							phase: 'api_attempt',
							total: meetings.length,
							current: i + 1,
							currentMeeting: meeting.text,
							message: `Trying API fetch for ${meeting.text}...`
						});
						transcript = await fetchTranscriptViaAPI(apiMeta);
					}

					// Step 4: Fall back to DOM if API didn't work
					if (!transcript && hasTranscriptInDOM) {
						await sleep(500); // settle time
						transcript = extractCurrentTranscript();
						if (transcript.hasTranscript) {
							domFallbackCount++;
						}
					} else if (transcript && transcript.hasTranscript) {
						apiSuccessCount++;
					}

					if (transcript && transcript.hasTranscript) {
						results.push({
							meetingDate: meeting.text,
							index: i,
							...transcript,
							apiMetadata: apiMeta || null
						});

						const sourceLabel = transcript.source === 'api' ? 'API' : 'DOM';
						dispatchProgress({
							phase: 'extracted',
							total: meetings.length,
							current: i + 1,
							currentMeeting: meeting.text,
							entryCount: transcript.entryCount,
							source: transcript.source,
							message: `[${sourceLabel}] ${transcript.entryCount} entries from ${meeting.text}`
						});
					} else {
						results.push({
							meetingDate: meeting.text,
							index: i,
							hasTranscript: false,
							vtt: '',
							txt: '',
							entryCount: 0,
							source: 'none',
							apiMetadata: apiMeta || null
						});

						dispatchProgress({
							phase: 'skipped',
							total: meetings.length,
							current: i + 1,
							currentMeeting: meeting.text,
							message: `No transcript for ${meeting.text}`
						});
					}
				} catch (err) {
					console.error(`[Batch Transcript] Error processing meeting ${i}:`, err);
					results.push({
						meetingDate: meeting.text,
						index: i,
						hasTranscript: false,
						vtt: '',
						txt: '',
						entryCount: 0,
						error: err.message
					});
				}
			}

			const withTranscript = results.filter((r) => r.hasTranscript);
			const totalEntries = withTranscript.reduce((sum, r) => sum + r.entryCount, 0);

			dispatchProgress({
				phase: 'complete',
				total: meetings.length,
				transcriptsFound: withTranscript.length,
				totalEntries,
				apiSuccessCount,
				domFallbackCount,
				message: `Done! ${withTranscript.length}/${meetings.length} had transcripts (${totalEntries} entries) — ${apiSuccessCount} via API, ${domFallbackCount} via DOM`
			});

			return {
				success: true,
				seriesName,
				results,
				sharePointTokens: getSharePointTokens(),
				summary: {
					total: meetings.length,
					withTranscript: withTranscript.length,
					totalEntries,
					apiSuccessCount,
					domFallbackCount
				}
			};
		} catch (err) {
			console.error('[Batch Transcript] Fatal error:', err);
			return { error: err.message };
		} finally {
			batchRunning = false;
		}
	};

	const cancelBatch = () => {
		batchCancelled = true;
	};

	// === Command listener ===
	document.addEventListener('teamsBatchTranscriptCommand', async (e) => {
		const { command } = e.detail || {};
		let result;

		switch (command) {
			case 'start':
				result = await downloadAllTranscripts();
				break;
			case 'cancel':
				cancelBatch();
				result = { cancelled: true };
				break;
			case 'status':
				result = { running: batchRunning };
				break;
			case 'ping':
				result = { available: true };
				break;
			default:
				result = { error: `Unknown command: ${command}` };
		}

		document.dispatchEvent(new CustomEvent('teamsBatchTranscriptResponse', {
			detail: { command, result }
		}));
	});

	console.log('[Teams Chat Exporter] Batch transcript download initialized (API + DOM hybrid)');
})();
