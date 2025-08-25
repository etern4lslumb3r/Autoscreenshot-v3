# Autoscreenshot v3

A Tkinter GUI app that periodically captures screenshots, detects visual changes, and uploads snapshots into a target Google Doc. Designed for Windows; works best when launched without a console via the provided `.pyw` launcher.

## Features
- __Google Docs integration__: append screenshots to a specified document.
- __Change detection__: only upload when local changes exceed thresholds.
- __Flexible region__: capture full screen or a user-defined bounding box (Left, Top, Width, Height).
- __Quick region tools__: "Select Region on Screen" and "Clear Region Fields".
- __Advanced thresholds__: editable numeric fields synced with sliders.
- __Non-resizable window__: fixed-size UI for consistency.

## Requirements
- Python 3.9+ (tested on 3.13)
- See `requirements.txt` for Python packages:
  - opencv-python, numpy, pyautogui, pillow
  - google-api-python-client, google-auth, google-auth-oauthlib, google-auth-httplib2
  - pywin32 (Windows only)

Install dependencies:
```powershell
pip install -r requirements.txt
```

## Google API Setup
1. Enable APIs in Google Cloud Console:
   - Google Drive API
   - Google Docs API
2. Create OAuth 2.0 Client ID (Desktop app) and download the JSON.
3. Save the file as `credentials.json` in the project root.
4. First run will prompt a browser to authorize your Google account and create `token.json`.

Both `credentials.json` and the generated `token.json` must sit beside `autoscreenshot.py`.

## Running the App
- __No console (recommended on Windows)__:
  - Double-click `run_autoscreenshot.pyw`
  - or run: `pythonw run_autoscreenshot.pyw`
- __With console__ (for debugging):
  - `python autoscreenshot.py`

## Usage
1. __Google Doc ID__
   - Paste a full Doc URL or just the ID into the "Google Doc ID" field. Pasting a URL extracts the ID silently.
2. __Region (Bounding Box)__
   - Leave all fields blank for full-screen capture, or fill in `Left (x)`, `Top (y)`, `Width`, `Height`.
   - Click __Select Region on Screen__ to draw a rectangle interactively.
   - Click __Clear Region Fields__ to blank all four inputs quickly.
3. __Interval__
   - Set "Check Interval (seconds)" for how often screenshots are checked.
4. __Advanced Settings__
   - __Mean Deviation Threshold__: controls sensitivity vs. average change. Entry is editable and synced with the slider.
   - __Min Mean Threshold__: clamps how low the mean can go; prevents hypersensitivity at very low noise.
5. __Start/Stop__
   - Click __Start Monitoring__ to begin. Inputs related to region and IDs are disabled during monitoring.
   - Click __Stop Monitoring__ to end. App waits for background tasks to finish cleanly.

## Notes
- Tooltips are disabled to keep the UI uncluttered.
- The main window is not resizable by design.
- When region fields are empty, the full screen is captured.

## Troubleshooting
- __Missing packages__: run `pip install -r requirements.txt` again and ensure you're using the same interpreter that launches the app.
- __Google auth issues__: delete `token.json` and restart to re-auth. Ensure the right Google account is used and that Docs/Drive APIs are enabled.
- __Permission/Screen capture issues__: on Windows, run the IDE/terminal with sufficient permissions. Close any apps that might block screen capture overlays.

## Project Files
- `autoscreenshot.py` — Main Tkinter application.
- `run_autoscreenshot.pyw` — Windows launcher that hides the console window.
- `requirements.txt` — Python dependencies.
- `credentials.json` — OAuth client secrets (user-provided; not committed).
- `token.json` — OAuth token (auto-generated on first auth; not committed).
