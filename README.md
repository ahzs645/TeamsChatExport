# Microsoft Teams Chat Exporter Chrome Extension

A Chrome extension that allows you to export Microsoft Teams chat conversations to HTML format directly from the Teams web interface.

## Features

- Export chat conversations to HTML format
- Preserves message formatting and attachments
- No Azure AD registration or admin access required
- Simple and intuitive user interface

## Installation

1. Download or clone this repository
2. Open Chrome and navigate to `chrome://extensions/`
3. Enable "Developer mode" in the top right corner
4. Click "Load unpacked" and select the `chrome-extension` folder

## Usage

1. Navigate to the chat you want to export in Microsoft Teams web
2. Click the extension icon in your browser toolbar
3. Click "Export Chats"
4. The chat will be exported as an HTML file to your downloads folder

## Limitations

- Can only export the currently visible chat
- Requires the Teams web interface to be open
- Some complex message formatting might not be preserved perfectly

## Development

The extension consists of the following files:

- `manifest.json`: Extension configuration
- `popup.html`: User interface
- `popup.js`: UI interaction handling
- `content.js`: Chat data extraction
- `background.js`: Main export functionality

## License

MIT License - see the LICENSE file for details 