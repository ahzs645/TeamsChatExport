/**
 * Transcript API Fetcher
 * Passively intercepts readcollabobject responses via Response.prototype.json
 * to extract transcript metadata and URLs from Teams meeting recap data.
 *
 * Teams caches a reference to native window.fetch at bundle load time,
 * so patching window.fetch won't catch these calls. However, patching
 * Response.prototype.json works because Teams still has to read the response
 * body, and prototype methods can't be pre-cached.
 *
 * SharePoint auth tokens are captured by chrome.webRequest in background.js
 * and forwarded here via CustomEvent from the content script.
 *
 * Collected data is exposed on window.__transcriptAPIData and via a hidden
 * DOM element for the content script to read.
 */
(() => {
	if (Response.prototype.__teamsChatExporterPatched) return;
	Response.prototype.__teamsChatExporterPatched = true;

	const origJson = Response.prototype.json;
	const origText = Response.prototype.text;
	const HIDDEN_DIV_ID = 'transcript-api-data';

	// Store all captured transcript metadata keyed by thread+calendarEvent
	window.__transcriptAPIData = {};

	// Store captured SharePoint auth tokens
	window.__sharePointTokens = {};

	const updateHiddenDiv = () => {
		let div = document.getElementById(HIDDEN_DIV_ID);
		if (!div) {
			div = document.createElement('div');
			div.id = HIDDEN_DIV_ID;
			div.style.display = 'none';
			document.body.appendChild(div);
		}
		div.setAttribute('data-transcript-api', JSON.stringify(window.__transcriptAPIData));
		div.setAttribute('data-sp-tokens', JSON.stringify(window.__sharePointTokens));
		div.setAttribute('data-updated', Date.now().toString());
	};

	/**
	 * Extract transcript metadata from a readcollabobject response.
	 */
	const extractFromCollabObject = (url, data) => {
		if (!data || !data.resources) return;

		// Parse the URL to get calendarEventId and threadId
		// Format: /readcollabobject/V2/{organizerId}@{tenantId}/{calendarEventId}/{threadId}
		const urlParts = url.split('/');
		const collabIdx = urlParts.findIndex((p) => p === 'readcollabobject');
		let calendarEventId = '';
		let urlThreadId = '';
		if (collabIdx >= 0) {
			calendarEventId = urlParts[collabIdx + 3] || '';
			urlThreadId = urlParts[collabIdx + 4] || '';
		}

		const key = calendarEventId || urlThreadId || `collab-${Date.now()}`;

		const entry = {
			calendarEventId,
			threadId: urlThreadId,
			capturedAt: new Date().toISOString(),
			resources: {}
		};

		for (const resource of data.resources) {
			const type = resource.type || resource.resourceType || '';
			const metadata = resource.metadata || {};
			const location = resource.location || '';

			if (type.toLowerCase().includes('transcript')) {
				entry.resources[type] = {
					location,
					callId: metadata.callId || '',
					startTime: metadata.startTime || '',
					endTime: metadata.endTime || '',
					organizerId: metadata.organizerId || '',
					threadId: metadata.threadId || urlThreadId,
					transcriptId: metadata.transcriptId || '',
					itemId: metadata.itemId || '',
					driveId: metadata.driveId || ''
				};
			}

			// Also capture recording metadata for context
			if (type.toLowerCase().includes('recording') || type.toLowerCase().includes('video')) {
				entry.resources[type] = {
					location,
					startTime: metadata.startTime || '',
					endTime: metadata.endTime || '',
					duration: metadata.duration || ''
				};
			}
		}

		// Only store if we found transcript data
		if (Object.keys(entry.resources).length > 0) {
			window.__transcriptAPIData[key] = entry;
			updateHiddenDiv();
			console.log(`[Teams Chat Exporter] Captured transcript metadata for ${key}:`,
				Object.keys(entry.resources).join(', '));
		}
	};

	// Store transcript content captured directly from cdnmedia/transcripts responses
	window.__capturedTranscriptContent = window.__capturedTranscriptContent || {};

	const TRANSCRIPT_DATA_DIV_ID = 'captured-transcript-content';

	const updateTranscriptContentDiv = () => {
		let div = document.getElementById(TRANSCRIPT_DATA_DIV_ID);
		if (!div) {
			div = document.createElement('div');
			div.id = TRANSCRIPT_DATA_DIV_ID;
			div.style.display = 'none';
			document.body.appendChild(div);
		}
		div.setAttribute('data-transcripts', JSON.stringify(window.__capturedTranscriptContent));
		div.setAttribute('data-updated', Date.now().toString());
	};

	// === Video Drive Item Capture ===
	// Captures driveId and itemId from v2.1/drives/.../items/... API calls
	// so we can later request @content.downloadUrl for direct video download.
	window.__videoDriveItem = window.__videoDriveItem || {};
	const VIDEO_DRIVE_DIV_ID = 'video-drive-data';

	const updateVideoDriveDiv = () => {
		let div = document.getElementById(VIDEO_DRIVE_DIV_ID);
		if (!div) {
			div = document.createElement('div');
			div.id = VIDEO_DRIVE_DIV_ID;
			div.style.display = 'none';
			document.body.appendChild(div);
		}
		div.setAttribute('data-drive-item', JSON.stringify(window.__videoDriveItem));
		div.setAttribute('data-updated', Date.now().toString());
	};

	const captureDriveItem = (url, responseData) => {
		const match = url.match(/\/v2\.1\/drives\/([^/?]+)\/items\/([^/?]+)/);
		if (!match) return;
		const driveId = match[1];
		const itemId = match[2];
		const siteMatch = url.match(/^https?:\/\/[^/]+(\/personal\/[^/]+|\/sites\/[^/]+)/);
		const siteBase = siteMatch ? siteMatch[1] : '';
		const host = (() => { try { return new URL(url).hostname; } catch(e) { return ''; } })();
		const apiBase = `https://${host}${siteBase}/_api/v2.1/drives/${driveId}/items/${itemId}`;

		window.__videoDriveItem = {
			driveId,
			itemId,
			siteBase,
			host,
			capturedAt: Date.now(),
			apiBase,
			// If the response already has @content.downloadUrl, capture it
			downloadUrl: responseData?.['@content.downloadUrl'] || null,
			fileName: responseData?.name || null,
			fileSize: responseData?.size || null
		};
		updateVideoDriveDiv();
		console.log(`[Teams Chat Exporter] Captured drive item: ${itemId} (${responseData?.name || 'unknown'})`);

		// If we don't have the download URL yet, fetch it
		if (!window.__videoDriveItem.downloadUrl) {
			const token = window.__sharePointTokens?.[host]?.token;
			if (token) {
				fetch(apiBase, {
					headers: { 'Authorization': token, 'Accept': 'application/json' }
				}).then(r => r.json()).then(data => {
					if (data['@content.downloadUrl']) {
						window.__videoDriveItem.downloadUrl = data['@content.downloadUrl'];
						window.__videoDriveItem.fileName = data.name || window.__videoDriveItem.fileName;
						window.__videoDriveItem.fileSize = data.size || window.__videoDriveItem.fileSize;
						updateVideoDriveDiv();
						console.log(`[Teams Chat Exporter] Got download URL for ${data.name} (${Math.round((data.size || 0) / 1024 / 1024)}MB)`);
					}
				}).catch(() => {});
			}
		}
	};

	// Patch Response.prototype.json to intercept readcollabobject, cdnmedia/transcripts, and drive items
	Response.prototype.json = function () {
		const responseUrl = this.url || '';

		// Capture drive item IDs for direct video download
		if (responseUrl.includes('/v2.1/drives/') && responseUrl.includes('/items/')) {
			const cloned2 = this.clone();
			origJson.call(cloned2).then((data) => {
				try { captureDriveItem(responseUrl, data); } catch (e) {}
			}).catch(() => { captureDriveItem(responseUrl, null); });
		}

		if (responseUrl.includes('readcollabobject')) {
			const cloned = this.clone();
			origJson.call(cloned).then((data) => {
				try {
					extractFromCollabObject(responseUrl, data);
				} catch (e) {
					console.error('[Teams Chat Exporter] Error extracting collab data:', e);
				}
			}).catch(() => {});
		}

		// Intercept Stream player's transcript fetch (cdnmedia/transcripts)
		if (responseUrl.includes('cdnmedia/transcripts') || responseUrl.includes('/transcripts')) {
			const cloned = this.clone();
			origJson.call(cloned).then((data) => {
				try {
					if (data && (Array.isArray(data) || data.entries || data.captions || data.value)) {
						const entries = Array.isArray(data) ? data : (data.entries || data.captions || data.value || []);
						if (entries.length > 0) {
							const key = `stream-${Date.now()}`;
							window.__capturedTranscriptContent[key] = {
								entries,
								capturedAt: new Date().toISOString(),
								source: 'cdnmedia',
								entryCount: entries.length,
								url: responseUrl.split('?')[0]
							};
							updateTranscriptContentDiv();
							console.log(`[Teams Chat Exporter] Captured ${entries.length} transcript entries from cdnmedia/transcripts`);
						}
					}
				} catch (e) {
					console.error('[Teams Chat Exporter] Error capturing transcript content:', e);
				}
			}).catch(() => {});
		}

		return origJson.call(this);
	};

	// Also patch Response.prototype.text to catch transcript responses read as text (e.g. VTT)
	Response.prototype.text = function () {
		const responseUrl = this.url || '';

		if (responseUrl.includes('cdnmedia/transcripts') || responseUrl.includes('/transcripts')) {
			const cloned = this.clone();
			origText.call(cloned).then((text) => {
				try {
					if (!text || text.length < 10) return;

					let entries = [];
					// Try parsing as JSON first
					if (text.startsWith('{') || text.startsWith('[')) {
						try {
							const data = JSON.parse(text);
							entries = Array.isArray(data) ? data : (data.entries || data.captions || data.value || []);
						} catch (e) {}
					}

					// If we got entries from JSON, or if it's VTT format, store it
					const key = `stream-${Date.now()}`;
					if (entries.length > 0) {
						window.__capturedTranscriptContent[key] = {
							entries,
							capturedAt: new Date().toISOString(),
							source: 'cdnmedia-json',
							entryCount: entries.length,
							url: responseUrl.split('?')[0]
						};
					} else if (text.startsWith('WEBVTT')) {
						window.__capturedTranscriptContent[key] = {
							rawVtt: text,
							capturedAt: new Date().toISOString(),
							source: 'cdnmedia-vtt',
							entryCount: (text.match(/-->/g) || []).length,
							url: responseUrl.split('?')[0]
						};
					}

					if (window.__capturedTranscriptContent[key]) {
						updateTranscriptContentDiv();
						console.log(`[Teams Chat Exporter] Captured transcript via .text() (${window.__capturedTranscriptContent[key].source}, ${window.__capturedTranscriptContent[key].entryCount} entries)`);
					}
				} catch (e) {
					console.error('[Teams Chat Exporter] Error capturing transcript text:', e);
				}
			}).catch(() => {});
		}

		return origText.call(this);
	};

	// === Receive SharePoint tokens from background.js via content script ===
	// background.js captures Authorization headers via chrome.webRequest
	// and forwards them here through a CustomEvent dispatched by content.js.
	document.addEventListener('teamsSharePointToken', (e) => {
		const { host, token, capturedAt } = e.detail || {};
		if (host && token) {
			window.__sharePointTokens[host] = {
				token,
				capturedAt: capturedAt || Date.now(),
				host
			};
			updateHiddenDiv();
			console.log(`[Teams Chat Exporter] Received SharePoint token for ${host} (via webRequest)`);
		}
	});

	// === Fresh video download URL request ===
	// Content script dispatches 'tceRequestFreshVideoUrl', page world fetches and replies.
	document.addEventListener('tceRequestFreshVideoUrl', async () => {
		const driveItem = window.__videoDriveItem;
		if (!driveItem || !driveItem.apiBase) {
			document.dispatchEvent(new CustomEvent('tceFreshVideoUrlReady', {
				detail: { error: 'No drive item data available' }
			}));
			return;
		}

		const token = window.__sharePointTokens?.[driveItem.host]?.token;
		if (!token) {
			document.dispatchEvent(new CustomEvent('tceFreshVideoUrlReady', {
				detail: { error: 'No auth token available' }
			}));
			return;
		}

		try {
			const resp = await fetch(driveItem.apiBase, {
				headers: { 'Authorization': token, 'Accept': 'application/json' }
			});
			if (!resp.ok) {
				document.dispatchEvent(new CustomEvent('tceFreshVideoUrlReady', {
					detail: { error: `API returned ${resp.status}` }
				}));
				return;
			}
			const data = await resp.json();
			const downloadUrl = data['@content.downloadUrl'];
			if (downloadUrl) {
				// Also update the cached data
				driveItem.downloadUrl = downloadUrl;
				driveItem.fileName = data.name || driveItem.fileName;
				driveItem.fileSize = data.size || driveItem.fileSize;
				updateVideoDriveDiv();

				document.dispatchEvent(new CustomEvent('tceFreshVideoUrlReady', {
					detail: {
						downloadUrl,
						fileName: data.name,
						fileSize: data.size
					}
				}));
				console.log(`[Teams Chat Exporter] Fresh download URL obtained for ${data.name}`);
			} else {
				document.dispatchEvent(new CustomEvent('tceFreshVideoUrlReady', {
					detail: { error: 'No download URL in API response' }
				}));
			}
		} catch (e) {
			document.dispatchEvent(new CustomEvent('tceFreshVideoUrlReady', {
				detail: { error: e.message }
			}));
		}
	});

	console.log('[Teams Chat Exporter] API fetcher initialized (Response.prototype.json + webRequest token relay)');
})();
