**Agents Overview**

The agent is running on **windows**. Use only windows like commands and don't use grep.

- **Purpose:** Describe the main programmatic "agents" (components) in this repository and how to run/validate them.

**Formula Engine**: src/js/formula/engine/
- **Role:** Tokenizes, parses, and evaluates Excel formulas (async evaluation, cycle detection, builtins, locale maps, cross-sheet references).
- **Cross-Sheet Support:** Handles `Sheet2!A1` and `'Sheet Name'!B2` syntax for referencing cells in other worksheets.

**Renderer**: src/js/renderer.js
- **Role:** Renders sheets to the DOM, handles formatting, and requests formula evaluation for cells.

**App Entry**: src/js/main.js
- **Role:** Initializes the app, loads locale maps, wires `formulaEngine` into the renderer.

**Locale Maps**: src/formula/locales/
- **Role:** JSON files mapping localized Excel function names to canonical names (en, fr, de, es, pt).

**Tests**: tests/
- `tests/formula-tests.js` — Core formula unit tests (50+ tests).
- `tests/formula-locale-tests.js` — Localization tests for FR/DE (can be expanded).
- `tests/formula-crosssheet-tests.js` — Cross-sheet reference tests (14 tests).

**Commands**
Build the application:

```bash
npm install
npm run build
```

Run unit tests locally (Node >=14):

```bash
npm test
```

Or run individually:

```bash
node --experimental-vm-modules tests/formula-tests.js
node --experimental-vm-modules tests/formula-locale-tests.js
node --experimental-vm-modules tests/formula-crosssheet-tests.js
```

Development mode with auto-rebuild:

```bash
npm run watch
```

**Notes & Conventions**
- The formula engine exposes `createFormulaEngine()` and `evaluateFormula()` for evaluator integration.
- Locale maps can be extended by adding files under `src/formula/locales/` and loading them in `main.js`.
- The project uses only `jszip` as a third-party library; other parsing and evaluation logic is handwritten.

**Next steps**
- Add more locale maps as needed.
- Expand tests to cover additional Excel functions and edge cases.
