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

	// Patch Response.prototype.json to intercept readcollabobject responses
	Response.prototype.json = function () {
		const responseUrl = this.url || '';

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

		return origJson.call(this);
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

	console.log('[Teams Chat Exporter] API fetcher initialized (Response.prototype.json + webRequest token relay)');
})();
