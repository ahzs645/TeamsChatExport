/**
 * Batch Transcript Download
 * Two-phase approach for maximum speed:
 *
 * Phase 1 (Sequential — fast click-through):
 *   Click each meeting in the dropdown quickly (~2s each) just to trigger
 *   the readcollabobject API call. Collect all metadata without waiting
 *   for DOM to stabilize.
 *
 * Phase 2 (Parallel — API fetch):
 *   Fire all transcript API fetches in parallel (concurrency-limited).
 *   For meetings where API fails, fall back to slow DOM extraction.
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
	const LOAD_MORE_WAIT = 1500;
	const API_FETCH_TIMEOUT = 15000;
	const API_CONCURRENCY = 4;
	const CLICK_THROUGH_WAIT = 2000; // wait per meeting during fast click-through
	const METADATA_POLL_INTERVAL = 200;
	const STABILIZE_INTERVAL = 500;
	const STABILIZE_CHECKS = 2;
	const DOM_WAIT_TIMEOUT = 15000;

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

	const getGroupCount = () => document.querySelectorAll('[role="group"]').length;

	const isNoTranscriptVisible = () => {
		const el = document.querySelector(NO_TRANSCRIPT_SELECTOR);
		if (!el) return false;
		return getComputedStyle(el).display !== 'none';
	};

	/**
	 * Quick-select a meeting by index. Does NOT wait for DOM to stabilize.
	 * Just clicks the option to trigger the readcollabobject API call.
	 */
	const quickSelectMeeting = async (index) => {
		await openDropdown();
		const options = getVisibleOptions();
		if (index >= options.length) {
			closeDropdown();
			throw new Error(`Meeting index ${index} out of range (${options.length} options)`);
		}
		options[index].click();
		// Brief wait for the readcollabobject request to fire
		await sleep(CLICK_THROUGH_WAIT);
	};

	/**
	 * Full select + wait for DOM to stabilize (used for DOM fallback only).
	 */
	const selectMeetingAndWaitForDOM = async (index) => {
		const baselineCount = getGroupCount();
		await openDropdown();
		const options = getVisibleOptions();
		if (index >= options.length) {
			closeDropdown();
			return false;
		}
		options[index].click();
		await sleep(1000);
		await clickTranscriptTab();

		const startTime = Date.now();
		let droppedToZero = baselineCount === 0;
		let lastCount = -1;
		let consecutiveStable = 0;

		while (Date.now() - startTime < DOM_WAIT_TIMEOUT) {
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
		return getGroupCount() > 0;
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

	/**
	 * Find metadata entries captured after a given timestamp.
	 * Returns all new entries (there may be multiple if readcollabobject was called twice).
	 */
	const findNewMetadata = (allMetadata, knownKeys, capturedAfter) => {
		return Object.entries(allMetadata)
			.filter(([key, v]) =>
				!knownKeys.has(key) &&
				new Date(v.capturedAt).getTime() >= capturedAfter
			)
			.sort(([, a], [, b]) => new Date(a.capturedAt).getTime() - new Date(b.capturedAt).getTime());
	};

	/**
	 * Wait for at least one new metadata entry to appear.
	 */
	const waitForNewMetadata = async (knownKeys, capturedAfter, timeoutMs = 3000) => {
		const start = Date.now();
		while (Date.now() - start < timeoutMs) {
			const all = getAPIMetadata();
			const newEntries = findNewMetadata(all, knownKeys, capturedAfter);
			if (newEntries.length > 0) {
				return { key: newEntries[0][0], ...newEntries[0][1] };
			}
			await sleep(METADATA_POLL_INTERVAL);
		}
		return null;
	};

	const getTokenForUrl = (transcriptUrl) => {
		const tokens = getSharePointTokens();
		if (!transcriptUrl || Object.keys(tokens).length === 0) return null;
		try {
			const host = new URL(transcriptUrl).hostname;
			const tokenEntry = tokens[host];
			if (tokenEntry && tokenEntry.token) {
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
	 * Fetch transcript content directly via SharePoint API.
	 */
	const fetchTranscriptViaAPI = async (apiMeta) => {
		if (!apiMeta || !apiMeta.resources) return null;

		const transcriptResource =
			apiMeta.resources['TranscriptV2'] ||
			apiMeta.resources['transcriptV2'] ||
			Object.values(apiMeta.resources).find((r) =>
				r.location && r.location.includes('transcript')
			);

		if (!transcriptResource || !transcriptResource.location) return null;

		const transcriptUrl = transcriptResource.location;
		const token = getTokenForUrl(transcriptUrl);
		if (!token) return null;

		try {
			const controller = new AbortController();
			const timeout = setTimeout(() => controller.abort(), API_FETCH_TIMEOUT);

			const response = await fetch(transcriptUrl, {
				headers: { 'Authorization': token },
				signal: controller.signal
			});
			clearTimeout(timeout);

			if (!response.ok) return null;

			const contentType = response.headers.get('content-type') || '';
			const text = await response.text();
			if (!text || text.length < 10) return null;

			let vtt = '';
			let txt = '';
			let entryCount = 0;

			if (text.startsWith('WEBVTT')) {
				vtt = text;
				txt = convertVttToTxt(text);
				entryCount = (text.match(/-->/g) || []).length;
			} else if (contentType.includes('json') || text.startsWith('{') || text.startsWith('[')) {
				try {
					const data = JSON.parse(text);
					const entries = Array.isArray(data) ? data : (data.entries || data.captions || []);
					if (entries.length > 0) {
						const utils = window.__teamsTranscriptUtils;
						if (utils) {
							vtt = utils.buildVttTranscript(entries);
							txt = utils.buildTxtTranscript(entries);
						} else {
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
					return null;
				}
			} else {
				txt = text;
				vtt = 'WEBVTT\n\nNOTE Raw transcript from API\n\n1\n00:00:00.000 --> 99:59:59.000\n' + text;
				entryCount = text.split('\n').filter((l) => l.trim()).length;
			}

			if (entryCount === 0 && txt.length < 10) return null;

			return { hasTranscript: true, vtt, txt, entryCount, source: 'api' };
		} catch (err) {
			return null;
		}
	};

	const convertVttToTxt = (vttContent) => {
		const lines = vttContent.split('\n');
		const txtLines = [];
		let currentTimestamp = '';
		for (const line of lines) {
			if (line.includes('-->')) {
				currentTimestamp = line.split('-->')[0].trim().split('.')[0] || '00:00:00';
			} else if (line.trim() && !line.startsWith('WEBVTT') && !line.startsWith('NOTE') && !line.match(/^\d+$/)) {
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

	// ========== DOM EXTRACTION (fallback) ==========

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
					startOffset: timestamp, endOffset: timestamp,
					speakerDisplayName: speaker, text: listitem.textContent.trim(),
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

	// ========== CONCURRENCY HELPER ==========

	/**
	 * Run async tasks with a concurrency limit.
	 * tasks: array of () => Promise
	 */
	const runWithConcurrency = async (tasks, limit) => {
		const results = new Array(tasks.length);
		let nextIndex = 0;

		const runNext = async () => {
			while (nextIndex < tasks.length) {
				if (batchCancelled) return;
				const i = nextIndex++;
				results[i] = await tasks[i]();
			}
		};

		const workers = [];
		for (let w = 0; w < Math.min(limit, tasks.length); w++) {
			workers.push(runNext());
		}
		await Promise.all(workers);
		return results;
	};

	// ========== MAIN ORCHESTRATOR ==========

	const downloadAllTranscripts = async () => {
		if (batchRunning) {
			return { error: 'Batch download already in progress' };
		}

		batchRunning = true;
		batchCancelled = false;

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
				message: `Found ${meetings.length} meetings. Phase 1: collecting metadata...`,
				seriesName
			});

			// ===== PHASE 1: Fast click-through to collect API metadata =====
			// Click each meeting quickly to trigger readcollabobject calls.
			// We track which metadata keys existed before each click so we can
			// associate new metadata with the correct meeting.

			const meetingMetadata = new Array(meetings.length).fill(null);
			const knownKeys = new Set(Object.keys(getAPIMetadata()));

			for (let i = 0; i < meetings.length; i++) {
				if (batchCancelled) break;

				const meeting = meetings[i];
				dispatchProgress({
					phase: 'collecting',
					total: meetings.length,
					current: i + 1,
					currentMeeting: meeting.text,
					message: `Collecting metadata ${i + 1}/${meetings.length}: ${meeting.text}`
				});

				const beforeClick = Date.now();

				try {
					await quickSelectMeeting(meeting.index);

					// Poll briefly for new metadata
					const meta = await waitForNewMetadata(knownKeys, beforeClick, 2000);
					if (meta) {
						meetingMetadata[i] = meta;
						knownKeys.add(meta.key);
					}
				} catch (err) {
					console.warn(`[Batch] Failed to click meeting ${i}:`, err.message);
				}
			}

			if (batchCancelled) {
				dispatchProgress({ phase: 'cancelled', message: 'Batch cancelled by user' });
				return { cancelled: true };
			}

			const metaCollected = meetingMetadata.filter(Boolean).length;
			dispatchProgress({
				phase: 'fetching',
				total: meetings.length,
				current: 0,
				message: `Phase 2: fetching ${metaCollected} transcripts via API (${API_CONCURRENCY} parallel)...`
			});

			// ===== PHASE 2: Parallel API fetches =====
			const results = new Array(meetings.length).fill(null);
			let apiSuccessCount = 0;
			let fetchedSoFar = 0;

			// Build tasks for meetings that have metadata
			const apiTasks = meetings.map((meeting, i) => async () => {
				if (batchCancelled) return;

				const meta = meetingMetadata[i];
				if (!meta) return; // no metadata — will try DOM later

				const transcript = await fetchTranscriptViaAPI(meta);
				fetchedSoFar++;

				if (transcript && transcript.hasTranscript) {
					apiSuccessCount++;
					results[i] = {
						meetingDate: meeting.text,
						index: i,
						...transcript,
						apiMetadata: meta
					};
					dispatchProgress({
						phase: 'extracted',
						total: meetings.length,
						current: fetchedSoFar,
						currentMeeting: meeting.text,
						entryCount: transcript.entryCount,
						source: 'api',
						message: `[API] ${transcript.entryCount} entries from ${meeting.text}`
					});
				} else {
					// API failed — mark for DOM fallback
					fetchedSoFar; // already incremented
					dispatchProgress({
						phase: 'api_failed',
						total: meetings.length,
						current: fetchedSoFar,
						currentMeeting: meeting.text,
						message: `API failed for ${meeting.text}, will try DOM`
					});
				}
			});

			await runWithConcurrency(apiTasks, API_CONCURRENCY);

			if (batchCancelled) {
				dispatchProgress({ phase: 'cancelled', message: 'Batch cancelled by user' });
				return { cancelled: true };
			}

			// ===== PHASE 3: DOM fallback for meetings that API missed =====
			const needDom = meetings
				.map((m, i) => ({ meeting: m, i }))
				.filter(({ i }) => !results[i] && !batchCancelled);

			let domFallbackCount = 0;

			if (needDom.length > 0) {
				dispatchProgress({
					phase: 'dom_fallback',
					total: meetings.length,
					current: 0,
					message: `Phase 3: DOM fallback for ${needDom.length} meetings...`
				});

				await clickTranscriptTab();

				for (const { meeting, i } of needDom) {
					if (batchCancelled) break;

					dispatchProgress({
						phase: 'dom_processing',
						total: meetings.length,
						current: i + 1,
						currentMeeting: meeting.text,
						message: `[DOM] Processing ${meeting.text}...`
					});

					try {
						const hasDOM = await selectMeetingAndWaitForDOM(meeting.index);
						if (hasDOM) {
							await sleep(500);
							const transcript = extractCurrentTranscript();
							if (transcript.hasTranscript) {
								domFallbackCount++;
								results[i] = {
									meetingDate: meeting.text,
									index: i,
									...transcript,
									apiMetadata: meetingMetadata[i] || null
								};
								dispatchProgress({
									phase: 'extracted',
									total: meetings.length,
									current: i + 1,
									currentMeeting: meeting.text,
									entryCount: transcript.entryCount,
									source: 'dom',
									message: `[DOM] ${transcript.entryCount} entries from ${meeting.text}`
								});
								continue;
							}
						}
					} catch (err) {
						console.warn(`[Batch] DOM fallback error for meeting ${i}:`, err.message);
					}

					// No transcript available
					results[i] = {
						meetingDate: meeting.text,
						index: i,
						hasTranscript: false,
						vtt: '', txt: '', entryCount: 0,
						source: 'none',
						apiMetadata: meetingMetadata[i] || null
					};
					dispatchProgress({
						phase: 'skipped',
						total: meetings.length,
						current: i + 1,
						currentMeeting: meeting.text,
						message: `No transcript for ${meeting.text}`
					});
				}
			}

			// Fill in any remaining nulls (cancelled meetings, etc.)
			for (let i = 0; i < results.length; i++) {
				if (!results[i]) {
					results[i] = {
						meetingDate: meetings[i].text,
						index: i,
						hasTranscript: false,
						vtt: '', txt: '', entryCount: 0,
						source: 'none',
						apiMetadata: meetingMetadata[i] || null
					};
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
				message: `Done! ${withTranscript.length}/${meetings.length} had transcripts (${totalEntries} entries) — ${apiSuccessCount} API, ${domFallbackCount} DOM`
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

	console.log('[Teams Chat Exporter] Batch transcript download initialized (fast parallel mode)');
})();
