# Teams chat export troubleshooting

## Check the build in the browser

1. In a source checkout, run `node scripts/build.mjs`. It refreshes the generated `build/` folder from the current source. The source `manifest.json` supplies the version, and GitHub releases package the same build folder.
2. In the Chrome or Edge profile that is signed in to Teams, open the extension manager and check the **Teams Chat Exporter** version and loaded folder. For a checkout, **Load unpacked** should point to `TeamsChatExport/build`. Reloading a copy from a previously extracted ZIP keeps that copy’s version and code.
3. Reload the extension in the extension manager, then **refresh the open Teams tab**. Refreshing the tab loads the updated content script and any helpers that start when the page loads.
4. Open the target chat and run **Extract this chat** once. Use the latest matching `Active chat:` entry in the Teams tab’s developer console to identify the run. If several Chrome profiles or Teams tabs are open, check the signed-in profile and chat before comparing logs.

The page-size fix shipped in **v6.7.5**. It caps an older saved page size of 500 at 200 in the exporter itself. If the latest console entry still reports `pageSize: 500`, check the loaded copy and refresh the Teams tab again. The extension card and the script executing in an already open tab can reflect different builds.

## Tell API export from scrolling

The exporter tries the chat API first when it has the active conversation ID and authorization from the existing Teams web session. It uses that sign-in without a separate Entra app registration. This API-first flow and its scrolling fallback both existed by v4.1. A successful export can come from either path.

| Console entry | What it means |
| --- | --- |
| `Skipped API export` | The active conversation ID or authorization was unavailable. The log reports which inputs were present. |
| `HTTP response` | A request reached an endpoint. The entry includes the host, region, API route version, page number, page size, HTTP status, and a rejection category when available. |
| `Success with endpoint, got N messages` | The API returned usable messages for the export. |
| `Using chat pane fallback` or `Scrolling to load older messages…` | The exporter switched to the chat pane after the API path was skipped or returned zero usable messages. |
| `Extracted N messages via DOM` | The scrolling fallback supplied the exported messages. |

On the Teams cloud layout, the scrollable element uses `data-tid="message-pane-list-viewport"` and message rows use `data-tid="chat-pane-message"`. Teams recycles rows as the viewport moves. The fallback saves each window of rendered messages, deduplicates them, and restores the original scroll position when the same chat is still open. It has bounded scroll and history-load attempts, so the result can contain less than the entire history.

## Diagnose failures

### The popup can’t connect

The manifest matches both `teams.cloud.microsoft` and `teams.microsoft.com`. The cloud app can show an active chat at the site root, so detect it by hostname even if the URL lacks `/v2/` or a chat ID. On `teams.cloud.microsoft`, the popup can retry injecting the content script using Chrome’s `scripting` permission when the tab has no responding listener. Check the loaded version, browser profile, and site access, then refresh Teams. The popup can connect and still fall back to scrolling if the API fails.

### The API returns HTTP 400 or 404

- For HTTP 400 with `history-parameters`, check `pageSize` first. The earlier **Fetch full history** preset requested 500 messages per page. A live run at 500 returned HTTP 400 on both the `amer` and `ca` chat-service routes. A successful request in the supplied HAR used `pageSize=200` and `startTime=1`.
- Current builds cap requests at 200 messages per page, including earlier saved preferences. **Fetch full history** increases the request limit to 200 pages. The popup’s “Up to” number is an upper bound, and paging can finish earlier when Teams runs out of messages.
- Rejection categories group known server error wording into chat ID, client version, authorization, or history parameters. The `history-parameters` label narrows the cause to request settings. The 500-versus-200 comparison made page size the leading explanation in the investigated run.
- When a cloud chat gets HTTP 404 on an `/api/mt/` route, check what the site used for chat history in a fresh capture. The investigated run got HTTP 404 on the attempted v2 routes. The HAR showed a successful v1 history route on the cloud host itself. Verify that route against the affected conversation type before treating it as a fix.

### Compare with Teams’ native request

The HAR showed a same-origin `GET` on `teams.cloud.microsoft` at `/api/chatsvc/ca/v1/users/ME/conversations/{chatId}/messages`. The request used `view=msnp24Equivalent|supportsMessageProperties`, `pageSize=200`, and `startTime=1`. The first response contained 200 messages and returned an older-page link in `_metadata.backwardLink`.

Teams can use different regions for different services. In the capture, `/api/mt/amer/` appeared alongside `/api/chatsvc/ca/`. The exporter now tracks those regions separately, prefers the cloud chat-service history route, and tries `ca` as a fallback when it only sees the other service’s region. It follows the server’s older-page link for the same origin and conversation path and stops if a link repeats.

To investigate another failure, capture the exact target chat and export run in the Teams tab’s **Network** panel. Compare the method, host, route pattern, status, pagination field names, and page size. Keep the HAR local and record only those details. HARs and console output can contain session credentials, conversation identifiers, and message content. Keep raw captures and response bodies out of the repo.

The extension also logs `[Teams native chat history]` during a short trace when the scrolling fallback starts on cloud Teams. `matchedChat: true` means the request URL contained the active conversation ID. `matchedChat: false` marks a generic history or batch request that may belong to another operation. `Trace started` alone confirms the trace was armed. In testing, Teams rendered more history without a matching request appearing in the trace.

## What was verified

- **Live API evidence** — Teams returned 200 messages through the cloud `ca` chat-service route in the HAR. That request was for a different conversation type from the chat being debugged.
- **Live fallback evidence** — A later export sent 500 and received HTTP 400 before scrolling, even after the capped build had been generated. That request confirms the page was still executing the older page-size behavior.
- **Chromium fixture** — The v6.7.5 build used a saved 500 setting, sent 200 instead, and exported both mock API pages without scrolling. The fixture also tested popup recovery and the virtualized scrolling path.
- **Remaining live check** — Run the current build in the target chat and confirm `pageSize: 200`, a successful API response, and `Success with endpoint` for that run. The mock result verifies the extension workflow. The live result verifies that Teams accepts the request for that chat.

## Run the checks locally

From the repo root, run these checks after changes to chat export or packaging.

```sh
node scripts/build.mjs
node scripts/teams-injected/chat-api-test.mjs
node scripts/teams-injected/transcript-test.mjs
```

With Playwright and its Chromium browser available to Node, run the unpacked-extension fixture and the wider injected UI regression suite. The fixture uses mock chat data.

```sh
node scripts/teams-injected/chat-test.mjs
node scripts/teams-injected/run.mjs "$(mktemp -d)"
```
