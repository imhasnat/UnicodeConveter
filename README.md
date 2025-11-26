# Bijoy to Unicode Converter

A Microsoft Word Add-in to convert Bangla text from Bijoy (e.g., SutonnyMJ) to Unicode directly inside Microsoft Word.

## About

This add-in allows you to convert Bengali text from Bijoy encoding to Unicode format seamlessly within Microsoft Word. It supports automatic conversion and provides tools to apply Arabic fonts to Arabic text in your documents.

**Author:** Hasnat Mahbub  
**License:** MIT  
**Repository:** [https://github.com/imhasnat/UnicodeConveter](https://github.com/imhasnat/UnicodeConveter)

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

### Basic Usage

1. Run `npm start` to sideload the add-in.
2. In Word, go to the **Home** tab and click **"Open Converter"** in the **Bijoy Converter** group to open the task pane.
3. **Convert text in the task pane:**
   - Paste or type Bijoy text in the **Input Text (Bijoy)** area
   - Conversion runs automatically and the Unicode output appears in the **Output Text (Unicode)** area
   - Copy the converted text from the output area to use in your document

4. **Convert text in your document:**
   - Select text in your Word document that uses SutonnyMJ font
   - Click the **"Convert SutonnyMJ"** button in the task pane
   - The add-in will convert words detected with SutonnyMJ font to Unicode

5. **Apply Arabic font (optional):**
   - Select text in your Word document that contains Arabic characters
   - Click **"Apply Font to Arabic Text"** to apply the Arabic font to Arabic text in the selection

### Notes
- The add‑in avoids altering non‑Bijoy text and performs conservative, word‑level conversion when needed.
- Ensure the SutonnyMJ (or other Bijoy) fonts are installed if your document uses them.
- The task pane conversion works with any Bijoy text, while the "Convert SutonnyMJ" button specifically targets text with SutonnyMJ font in your document.

## Troubleshooting
- If Word doesn’t show the latest changes, stop and restart: `npm run stop` then `npm start`.
- If sideload fails, close all Word instances and retry `npm start`.
- To regenerate dev certificates (rare):
  ```bash
  npx office-addin-dev-certs install
  ```

## Scripts Reference
- `npm start` — start debugging with `manifest.xml` (launches Word and sideloads the add‑in)
- `npm run dev-server` — webpack dev server only
- `npm run build` — production build
- `npm run validate` — validate `manifest.xml`
- `npm run stop` — stop the debugging session
- `npx kill-port 3000` — to terminate the server running on port 3000

## References
https://bsbk.portal.gov.bd/apps/bangla-converter/index.html which is Developed and Customized by: Md. Elias Hossain, Programmer, Ministry of Land.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](https://github.com/imhasnat/UnicodeConveter/issues).

