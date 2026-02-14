# BetterLetter → Docman Helper (Chrome Extension)

## Install (production-ready, unpacked)
1. Open Chrome and go to `chrome://extensions`.
2. Enable **Developer mode** (top-right).
3. Click **Load unpacked** and select the `chrome-extension` folder.

## Use
1. Stay logged into BetterLetter at `https://app.betterletter.ai`.
2. Click the extension icon.
3. Enter a partial practice name and click **Fetch & Login**.
4. The extension will:
   - Find the practice in the list
   - Open the practice page
   - Open **EHR Settings**
   - Read Docman credentials
   - Open Docman and submit login

## Notes
- If you are logged out of BetterLetter, log in and try again.
- The extension does not store credentials; it reads them and submits to Docman.
- Practice matching is case-insensitive and uses partial match.

## Files
- `manifest.json`: Extension manifest
- `background.js`: Orchestrates tab flow
- `content-betterletter.js`: Extracts ODS/Docman creds from BetterLetter
- `content-docman.js`: Fills Docman login form
- `popup.html` / `popup.js` / `popup.css`: UI
