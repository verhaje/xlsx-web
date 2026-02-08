# Copilot / AI Agent Instructions — xlsx_reader

Short: help an AI contributor be productive quickly in this small XLSX viewer project.

- Focus areas (quick): `src/js/formula/engine/`, `src/js/renderer.js`, `src/js/main.js`, `src/js/parser.js`, `src/js/app.bundle.js`, `tests/`

Architecture (big picture)
- The app is a browser-based XLSX viewer. `index.html` loads `src/js/app.bundle.js` (built bundle).
- `src/js/main.js` is the entry: it reads workbook XML from a ZIP (JSZip), builds shared strings and styles, creates the formula engine and calls `renderSheet()`.
- `src/js/renderer.js` renders DOM table UI for a sheet. It builds a `cellMap` keyed by `${row}-${col}` and iterates rows/cols to create table rows/tds. Rendering requests formula evaluation via the formula engine when a `<c>` contains an `<f>` node.
- `src/js/formula/engine/` implements formula parsing, built-in functions, A1 range expansion and `evaluateFormula(formulaText, context)`; it expects the `context.resolveCell` callback to fetch other cell values (may be cross-sheet refs).
- `src/js/parser.js` contains XML helpers and cell-ref utilities: `parseCellRef`, `expandRange`, `parseSqref`, `extractCellValue`, and shared-string loading.
- Styles and cell-format conversion live in `src/js/styles.js`; `resize.js` attaches column/row resizers to the rendered table.

Key data flows & conventions
- Cell maps and cache keys:
  - `cellMap` entries are addressed by key string `${row}-${col}` where `row` and `col` are 1-based integers.
  - `sheetCache` stores per-sheet info under key = sheet name; value includes `{ cellMap, sheetDoc, maxRow, maxCol }`.
  - Evaluation cache uses keys like `${sheetName}::${row}-${col}` stored under `sheetCache.get('__values')` to avoid duplicate resolves during a render.
- Formula evaluation:
  - Renderer calls `formulaEngine.evaluateFormula(fText, { resolveCell: async (r) => ..., sharedStrings, zip, sheetDoc })`.
  - `resolveCell` receives references like `A1` or `Sheet Name!A1` and should return the computed value (string/number or error token starting with `#`).
- Range behavior:
  - `expandRange(range)` and `parseSqref(sqref)` expand A1 ranges into arrays of `{row, col}` coords used by aggregating functions (SUM, COUNTIFS, etc.).
- DOM metadata:
  - Renderer sets `td.dataset.formula` when a formula is present and `td.dataset.value` for the computed value; selection UI reads these for the formula bar.

Project-specific patterns to follow
- Keep changes confined to small functions where possible — the codebase bundles with `esbuild` and uses a single entry `src/js/main.js`.
- When adding parsing features, extend `src/js/parser.js` (it centralizes XML helpers and cell-ref logic). Example helpers: `parseMergeCells()` and `parseSharedFormulas()` (see recent additions).
- When altering rendering or evaluation, prefer updating `renderSheet()` in `src/js/renderer.js` to preserve DOM and cache behavior; use `sheetCache` to avoid repeated ZIP reads.
- Formula engine contract: do not change `evaluateFormula` behavior unless you update all callers. The engine expects `resolveCell` to be async and to provide sheetDoc/sharedStrings/zip context if needed.
- Cell reference transformations: the code uses A1-style parsing and `columnIndexToName`/`columnNameToIndex` helpers. Respect absolute (`$A$1`) vs relative refs when shifting shared formulas.

Build / test / run workflows (practical)
- Install deps and build:
```bash
npm install
npm run build
```
- For development with auto-rebuild (recommended when editing UI/renderer):
```bash
npm run watch
```
- Tests are Node-based and use ES module mode. Use the provided npm script to run the whole suite:
```bash
npm test
```
- To run individual suites during development (helps isolate issues):
```bash
node --experimental-vm-modules tests/formula-tests.js
node --experimental-vm-modules tests/formula-crosssheet-tests.js
```
- Remember to rebuild (`npm run build`) after editing source files if testing via the browser `index.html` (it loads `src/js/app.bundle.js`). Node tests import `src/js/*.js` directly and don't need the bundle.

Editing & integration tips
- After changing parsing logic or renderer, run `npm test` first to validate engine behavior; then run `npm run build` and open `index.html` for manual UI checks.
- Use `sheetCache` and `__values` caching pattern in `renderer.js` when adding any cross-sheet or repeated evaluations.
- To add DOM features like merged cells (colspan/rowspan), set `td.colSpan` and `td.rowSpan` and ensure you skip creating `td` for covered cells, keeping `data-row-index`/`data-col-index` to preserve selection semantics.

Where to look for examples
- Formula parsing/evaluation examples: `src/js/formula/engine/` (many builtins and range-handling code)
- Rendering + evaluation invocation: `src/js/renderer.js` (search for `evaluateFormula` calls and `dataset.formula` usage)
- ZIP/XML helpers and ref utilities: `src/js/parser.js`
- Tests demonstrating function expectations: `tests/formula-*.js` files (run these to verify behavior)
- Build entry: `package.json` scripts (`build`, `watch`, `test`)

What AI agents should avoid changing without explicit guidance
- Changing the public `evaluateFormula` signature without updating all calls in `renderer.js` and tests.
- Replacing the cell key convention (`${row}-${col}`) — many caches and maps rely on it.
- Removing the `sharedStrings` indexing logic in `parseSharedStrings()` — tests and real sheets depend on it.

If you're unsure about a change
- Run `npm test`. If tests fail, inspect stack traces and change only the minimal code necessary.
- For UI regressions, run `npm run build` and open `index.html` in the browser (the viewer uses the bundle `src/js/app.bundle.js`).

Feedback
- If anything in these instructions is unclear or you'd like more examples (small PRs demonstrating the patterns), tell me which area to expand (parsing, renderer, engine, tests) and I'll update this document.
