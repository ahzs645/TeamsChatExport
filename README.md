# Teams Chat Exporter

A Chrome extension that exports Microsoft Teams chats, meeting transcripts and meeting recordings from Teams on the web. It works with your normal Teams sign-in: no Azure AD app registration or admin consent is needed.

## Features

### Chats
- **Extract this chat**: open a chat or channel in Teams, click the toolbar icon and choose *Extract this chat*. Progress is shown on the page and the viewer opens when it finishes.
- **Viewer**: browse every chat you have extracted, search across them, and pick which participant is you.
- **Export** from the viewer as **JSON**, **CSV**, **TXT** or a self-contained **HTML** file. You can also load a previously exported JSON file back into the viewer.

### Meeting transcripts
On a meeting recording (Teams recap, SharePoint or Stream `stream.aspx` page):
- **Copy transcript** to the clipboard.
- **Download subtitles (.vtt)** with timestamps, or **Download text (.txt)** for plain reading.
- **Download all meeting transcripts**: opens a panel on the page that walks through every meeting in a recurring series and saves each transcript.

If the transcript is not ready yet, play the recording for a few seconds so Teams loads it.

### Meeting recordings
- **Download video** saves the recording as a video file using the best method available on the page.
- **Other methods** lists the alternatives (original file, fast parallel download, saving what the browser has already played, or recording playback at high speed) with a short note on when each one works.

### Settings
Under *Advanced settings* in the popup:
- **Fetch full history** fetches as many messages as possible for long chats.
- **Messages per request** and **Maximum requests** fine-tune how much history is fetched. The popup shows the resulting upper limit ("Up to N messages").

The popup follows your system's light or dark theme.

## Installation

The extension is not on the Chrome Web Store; load it unpacked:

1. Download or clone this repository.
2. Open `chrome://extensions/` in Chrome (or another Chromium browser such as Edge).
3. Turn on **Developer mode** (top right).
4. Click **Load unpacked** and select the repository folder (the one that contains `manifest.json`).
5. Pin **Teams Chat Exporter** from the extensions menu so its icon is always visible.

After installing or updating, reload any Teams tabs that were already open.

## Usage

1. Open [Teams on the web](https://teams.microsoft.com) and go to the chat, channel or meeting recording you want.
2. Click the **Teams Chat Exporter** icon in the toolbar. The badge under the title tells you what the extension sees on the page (for example *Teams chat: Project Sync* or *Meeting recording*).
3. Use the action shown for that page. Use **Open viewer** at any time to see what you have already extracted.

## Limitations

- Works with Teams on the web only (`teams.microsoft.com`, `teams.cloud.microsoft`) and with recordings on SharePoint or Stream. The desktop app is not supported.
- Chat extraction relies on Teams' web interface and APIs, which change from time to time; some message types or formatting may not come through perfectly.
- Video download depends on what the recording allows. Some methods need the video to have played briefly, and recording playback needs the tab to stay in the foreground.

## Development

There is no build step. Edit the files and click the reload button for the extension on `chrome://extensions/`.

| File | Purpose |
| --- | --- |
| `manifest.json` | Extension configuration (Manifest V3) |
| `popup.html`, `popup.css`, `popup.js` | Toolbar popup |
| `tokens.css` | Shared colours, spacing and focus styles (light and dark) |
| `content.js` | Runs on Teams/SharePoint pages: reads the current chat, injects transcript and video controls, answers popup messages |
| `src/modules/` | Chat extraction engine and Teams API/variant helpers |
| `results.html`, `results.js`, `style.css` | Viewer and JSON/CSV/TXT/HTML export |
| `transcript*.js`, `batchTranscriptDownload.js` | Transcript capture and batch download |
| `videoDownload/` | Video download methods and the coordinator that picks one |
| `background.js` | Service worker: stores extractions and opens the viewer |

## License

MIT License.
