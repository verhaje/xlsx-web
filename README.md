# xlsx_reader

Instant, private Excel previews — right in your browser.

[![GitHub Pages](https://img.shields.io/badge/GitHub-Pages-blue?logo=github)](https://verhaje.github.io/xlsx-web/)

Live demo: https://verhaje.github.io/xlsx-web/

Feel the relief of opening a spreadsheet without uploading it. xlsx_reader is a tiny, client-side viewer and formula engine that shows XLSX sheets quickly, accurately, and privately.

Why you'll love it
- 🚀 **Fast:** Fully client-side, small JS bundle for instant previews.
- 🔒 **Private:** Files never leave your machine — no server uploads.
- 🧮 **Accurate formulas:** Supports a large subset of Excel formulas, cross-sheet references, and locale mappings.
- 🛠️ **Extensible:** Clear plugin points for locale maps, builtins, and the renderer.

Quick start (try it now)

1. Install dependencies

```bash
npm install
```

2. Build the viewer

```bash
npm run build
```

3. Open the app

Open `index.html` in your browser and drop an .xlsx on the page — enjoy a private preview in seconds.

Development & testing

- Run the full test suite

```bash
npm test
```

- Development watch (auto-rebuild)

```bash
npm run watch
```

Where to find more details

- Technical docs and contributing guidelines: `docs/`
- Formula engine internals: `src/formula/engine/`

Contributing

We welcome contributors! Please read `docs/` before opening PRs — it explains architecture, formula contracts, and testing expectations.

License

This project is open-source — see the repository license for details.
