BijoyToUnicode Word Add‑in

Convert Bangla text from Bijoy (e.g., SutonnyMJ) to Unicode directly inside Microsoft Word.

## Prerequisites
- Node.js 18+ and npm
- Microsoft Word (Desktop)
- Windows (recommended for Bijoy/SutonnyMJ fonts)

## Install
```bash
npm install
```

## Run (start Word with the add‑in)
This launches a local dev server and sideloads the add‑in into Word using `manifest.xml`.
```bash
npm start
```

If Word is already open, close it before running `npm start` for a clean sideload.

## Development server only (optional)
If you only want the webpack dev server (without auto‑launching Word):
```bash
npm run dev-server
```

## Build
Production build outputs to the `dist` folder used by the add‑in pages.
```bash
npm run build
```

## Validate manifest
```bash
npm run validate
```

## Stop the debugging session
```bash
npm run stop
```
 

## Using the Add‑in in Word
1. Run `npm start` to sideload.
2. In Word, open the task pane for this add‑in.
3. Paste or type Bijoy text in the input area — conversion runs automatically and the Unicode output appears.
4. Use the “Convert” action to analyze the selection and convert words detected with SutonnyMJ font to Unicode.

Notes
- The add‑in avoids altering non‑Bijoy text and performs conservative, word‑level conversion when needed.
- Ensure the SutonnyMJ (or other Bijoy) fonts are installed if your document uses them.

## Troubleshooting
- If Word doesn’t show the latest changes, stop and restart: `npm run stop` then `npm start`.
- If sideload fails, close all Word instances and retry `npm start`.
- To regenerate dev certificates (rare):
  ```bash
  npx office-addin-dev-certs install
  ```

## Scripts reference
- `npm start` — start debugging with `manifest.xml` (launches Word and sideloads the add‑in)
- `npm run dev-server` — webpack dev server only
- `npm run build` — production build
- `npm run validate` — validate `manifest.xml`
- `npm run stop` — stop the debugging session 
- `npm run stop` — stop the debugging session 
- `npx kill-port 3000` — to terminate the server running on port 3000

