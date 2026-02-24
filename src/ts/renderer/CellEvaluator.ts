// CellEvaluator.ts - Cell evaluation with cross-sheet caching

import type JSZip from 'jszip';
import type { SheetInfo, SheetModel, SheetCacheEntry, IFormulaEngine } from '../types';
import { XmlParser } from '../core/XmlParser';
import { CellReference } from '../core/CellReference';
import { SheetParser } from '../workbook/SheetParser';
import { FormulaShifter } from './FormulaShifter';

interface CellFormulaInfo {
  formulaText: string | null;
  sharedInfo: { si: string; anchorRef?: string; baseFormula?: string } | null;
}

/**
 * Manages cell evaluation, cross-sheet references, and caching during rendering.
 */
export class CellEvaluator {
  private zip: JSZip;
  private sheet: SheetInfo;
  private sheets: SheetInfo[];
  private sharedStrings: string[];
  private formulaEngine: IFormulaEngine;

  /** Per-sheet cached cell maps and sheet docs */
  private sheetCache: Map<string, any>;
  /** Promise coalescing for sheet building */
  private buildingSheets: Map<string, Promise<SheetCacheEntry>>;
  /** Cells currently being evaluated (circular reference detection) */
  private evaluatingCells: Set<string>;
  /** Promise coalescing for cell evaluation */
  private pendingCells: Map<string, Promise<any>>;
  /** Shared range cache for resolved range arrays */
  private rangeCache: Map<string, any[]>;
  /** Promise coalescing for in-flight resolveRange calls */
  private pendingRanges: Map<string, Promise<any[]>>;
  /** Sheets that currently have an in-flight _resolveRangeImpl (cycle detection) */
  private activeRangeSheets: Set<string>;
  /** Pre-built sheet name → SheetInfo lookup (avoids O(n) Array.find per call) */
  private sheetLookup: Map<string, SheetInfo>;

  constructor(
    zip: JSZip,
    sheet: SheetInfo,
    sheets: SheetInfo[],
    sharedStrings: string[],
    formulaEngine: IFormulaEngine,
    sheetCache: Map<string, any>,
    rangeCache?: Map<string, any[]>
  ) {
    this.zip = zip;
    this.sheet = sheet;
    this.sheets = sheets;
    this.sharedStrings = sharedStrings;
    this.formulaEngine = formulaEngine;
    this.sheetCache = sheetCache;
    this.buildingSheets = new Map();
    this.evaluatingCells = new Set();
    this.pendingCells = new Map();
    this.rangeCache = rangeCache || new Map();
    this.pendingRanges = new Map();
    this.activeRangeSheets = new Set();

    // Build sheet lookup map once instead of using Array.find() per call
    this.sheetLookup = new Map();
    if (Array.isArray(sheets)) {
      for (const s of sheets) {
        this.sheetLookup.set(s.name, s);
      }
    }
  }

  /**
   * Get the formula text for a cell element, resolving shared formulas if needed.
   */
  static getCellFormula(cellEl: Element, cellRef: string, sheetModel: SheetModel): string | null {
    const fNode = cellEl.getElementsByTagName('f')[0];
    if (!fNode) return null;
    const fType = fNode.getAttribute('t');
    if (fType === 'shared' && sheetModel?.sharedCells) {
      const sharedInfo = sheetModel.sharedCells.get(cellRef);
      if (sharedInfo?.baseFormula && sharedInfo?.anchorRef) {
        return FormulaShifter.deriveSharedFormula(sharedInfo.baseFormula, sharedInfo.anchorRef, cellRef);
      }
    }
    return fNode.textContent || '';
  }

  /**
   * Get a cache key for a shared formula instance.
   */
  static getSharedCacheKey(
    sheetName: string,
    sharedInfo: { si: string; anchorRef?: string } | null,
    targetRef: string
  ): string | null {
    if (!sharedInfo?.anchorRef) return null;
    const anchor = CellReference.parse(sharedInfo.anchorRef);
    const target = CellReference.parse(targetRef);
    if (!anchor.row || !anchor.col || !target.row || !target.col) return null;
    return `${sheetName}::${sharedInfo.si}::${target.row - anchor.row},${target.col - anchor.col}`;
  }

  /**
   * Evaluate a cell by A1 reference, resolving cross-sheet refs and using caches.
   */
  async evaluateCellByRef(refA1: string, contextSheetName?: string): Promise<any> {
    let targetSheetName = contextSheetName || this.sheet.name;
    let a1 = refA1;

    if (typeof refA1 === 'string' && refA1.includes('!')) {
      const parts = refA1.split('!');
      targetSheetName = parts.slice(0, -1).join('!');
      a1 = parts[parts.length - 1];
    }
    targetSheetName = CellEvaluator.stripSheetQuotes(targetSheetName);

    const targetSheetObj: SheetInfo | undefined = this.sheetLookup.get(targetSheetName);

    const info = targetSheetObj
      ? await this.buildCellMapForSheet(targetSheetObj)
      : { sheetModel: null } as any;

    // Parse once, pass pre-parsed row/col to mapMergedRef to avoid double-parse
    let parsed = CellReference.parse(a1);
    if (!parsed.col || !parsed.row) return '';
    a1 = CellReference.mapMergedRef(a1, info.sheetModel, parsed.row, parsed.col);
    // Re-parse only if mapMergedRef actually changed the ref (merged cell)
    if (a1 !== CellReference.toA1(parsed.row, parsed.col)) {
      parsed = CellReference.parse(a1);
      if (!parsed.col || !parsed.row) return '';
    }

    const key = `${targetSheetName}::${parsed.row}-${parsed.col}`;

    // Single get for the values cache (avoids has+get double lookup)
    let valuesCache: Map<string, any> | undefined = this.sheetCache.get('__values');
    if (valuesCache) {
      const cached = valuesCache.get(key);
      if (cached !== undefined) return cached;
    }

    // Circular / transitive-dependency guard — MUST come before the
    // pendingCells check.  When two sheets reference each other via
    // range-based lookups (e.g. IST→VLOOKUP(DPT!A:AAB) and
    // DPT→VLOOKUP(IST!A:ZZ)), the range resolution of the second sheet
    // discovers formula cells that are already mid-evaluation on the
    // first sheet.  Previously, pendingCells.has() fired first and
    // returned the in-flight promise — which transitively awaited itself,
    // creating a deadlock that hung the UI with the loading spinner.
    //
    // Returning the XML-cached <v> value (instead of '#REF!') matches
    // Excel's behaviour for iterative / circular references and keeps
    // the computed results correct for the initial file load.
    if (this.evaluatingCells.has(key)) {
      const cachedInfo = this.sheetCache.get(targetSheetName);
      if (cachedInfo?.cellMap) {
        const cellEl = cachedInfo.cellMap.get(`${parsed.row}-${parsed.col}`);
        if (cellEl) {
          const fallback = SheetParser.extractCellValue(cellEl, this.sharedStrings);
          return fallback;
        }
      }
      return '';
    }

    // Prevent thundering herd — safe now that circular deps are caught above
    if (this.pendingCells.has(key)) return this.pendingCells.get(key);
    // Ensure values map
    if (!valuesCache) { valuesCache = new Map(); this.sheetCache.set('__values', valuesCache); }
    this.evaluatingCells.add(key);

    const promise = (async () => {
      const cmap: Map<string, Element> = info.cellMap || new Map();
      const cellEl = cmap.get(`${parsed.row}-${parsed.col}`);
      if (!cellEl) {
        valuesCache!.set(key, '');
        return '';
      }

      let val: any;
      const cellRef = CellReference.toA1(parsed.row, parsed.col);
      const formulaText = CellEvaluator.getCellFormula(cellEl, cellRef, info.sheetModel);

      if (formulaText !== null && this.formulaEngine) {
        const sharedInfo = info.sheetModel?.sharedCells?.get(cellRef) || null;
        const cacheKey = CellEvaluator.getSharedCacheKey(targetSheetName, sharedInfo, cellRef);
        if (cacheKey && this.sheetCache.get('__sharedValues')?.has(cacheKey)) {
          val = this.sheetCache.get('__sharedValues').get(cacheKey);
        } else {
          val = await this.formulaEngine.evaluateFormula(formulaText, {
            resolveCell: async (r: string) => await this.evaluateCellByRef(r, targetSheetName),
            resolveCellsBatch: async (refs: string[]) => await this.evaluateCellsBatch(refs, targetSheetName),
            resolveRange: async (sn: string | undefined, sr: number, sc: number, er: number, ec: number) =>
              await this.resolveRange(sn, sr, sc, er, ec, targetSheetName),
            sharedStrings: this.sharedStrings,
            zip: this.zip,
            rangeCache: this.rangeCache,
            getSheetMaxRow: (sheetName?: string) => this.getMaxRowForSheet(sheetName || targetSheetName),
          });
          if (cacheKey) {
            if (!this.sheetCache.has('__sharedValues')) this.sheetCache.set('__sharedValues', new Map());
            this.sheetCache.get('__sharedValues').set(cacheKey, val);
          }
        }
      } else {
        val = SheetParser.extractCellValue(cellEl, this.sharedStrings);
      }

      valuesCache!.set(key, val);
      return val;
    })();

    this.pendingCells.set(key, promise);
    try {
      return await promise;
    } finally {
      this.evaluatingCells.delete(key);
      this.pendingCells.delete(key);
    }
  }

  // ---- Private helpers ----

  private async buildCellMapForSheet(sheetObj: SheetInfo): Promise<SheetCacheEntry> {
    if (!sheetObj?.target) return { maxRow: 0, maxCol: 0 };
    if (this.sheetCache.has(sheetObj.name)) return this.sheetCache.get(sheetObj.name);
    if (this.buildingSheets.has(sheetObj.name)) return this.buildingSheets.get(sheetObj.name)!;

    const promise = (async (): Promise<SheetCacheEntry> => {
      try {
        const doc = await XmlParser.readZipXml(this.zip, sheetObj.target!);
        const data = doc.getElementsByTagName('sheetData')[0];
        const rows = data ? Array.from(data.getElementsByTagName('row')) : [];
        const cmap = new Map<string, Element>();
        let mRow = 0;
        let mCol = 0;
        rows.forEach((r) => {
          const rowIndex = parseInt(r.getAttribute('r') || '0', 10);
          const actualRow = Number.isNaN(rowIndex) || rowIndex === 0 ? mRow + 1 : rowIndex;
          mRow = Math.max(mRow, actualRow);
          const cells = Array.from(r.getElementsByTagName('c'));
          cells.forEach((cell) => {
            const ref = cell.getAttribute('r') || '';
            const { col } = CellReference.parse(ref);
            if (col === 0) return;
            mCol = Math.max(mCol, col);
            cmap.set(CellReference.cellKey(actualRow, col), cell);
          });
        });
        const model = SheetParser.buildSheetModel(doc);
        const info: SheetCacheEntry = { cellMap: cmap, sheetDoc: doc, maxRow: mRow, maxCol: mCol, sheetModel: model };
        this.sheetCache.set(sheetObj.name, info);
        return info;
      } catch {
        return { maxRow: 0, maxCol: 0 };
      } finally {
        this.buildingSheets.delete(sheetObj.name);
      }
    })();

    this.buildingSheets.set(sheetObj.name, promise);
    return promise;
  }

  /**
   * Get the maximum populated row for a sheet by name.
   * Returns the cached maxRow from the sheet cache if available, otherwise 1000 as a fallback.
   */
  private getMaxRowForSheet(sheetName: string): number {
    const cleaned = CellEvaluator.stripSheetQuotes(sheetName);
    const entry = this.sheetCache.get(cleaned) as SheetCacheEntry | undefined;
    if (entry && entry.maxRow && entry.maxRow > 0) return entry.maxRow;
    return 200; // conservative fallback for not-yet-loaded sheets
  }

  private static stripSheetQuotes(name: string): string {
    if (name.length >= 2 && name.startsWith("'") && name.endsWith("'")) return name.slice(1, -1);
    return name;
  }

  /**
   * Parse a ref string into sheet name and A1 reference.
   * Optimized to avoid .split()/.slice()/.join() allocations for the common case.
   */
  private splitRef(refA1: string, contextSheetName?: string): { sheetName: string; a1: string } {
    if (typeof refA1 === 'string') {
      const bangIdx = refA1.lastIndexOf('!');
      if (bangIdx !== -1) {
        const rawSheet = refA1.slice(0, bangIdx);
        const a1 = refA1.slice(bangIdx + 1);
        return { sheetName: CellEvaluator.stripSheetQuotes(rawSheet), a1 };
      }
    }
    return { sheetName: contextSheetName || this.sheet.name, a1: refA1 };
  }

  /**
   * Resolve a contiguous range of cells by numeric coordinates.
   * This bypasses A1 string expansion entirely — instead of creating thousands
   * of "A1", "A2"… strings via RangeUtils.expandRange, then parsing them back
   * in evaluateCellsBatch, we iterate row/col directly.
   * Eliminates ~20s of CellReference.parse overhead and ~10s of GC pressure
   * on sheets with large cross-sheet range references.
   */
  async resolveRange(
    sheetName: string | undefined,
    startRow: number, startCol: number,
    endRow: number, endCol: number,
    contextSheetName?: string
  ): Promise<any[]> {
    const targetSheetName = CellEvaluator.stripSheetQuotes(sheetName || contextSheetName || this.sheet.name);

    // Circular range-dependency guard.
    // When two sheets mutually reference each other via range-based lookups
    // (e.g. IST→VLOOKUP(DPT!A:AAB) and DPT→VLOOKUP(IST!A:ZZ)), the range
    // resolution chains deadlock: _resolveRangeImpl('SheetA') awaits
    // formula cells whose evaluation triggers _resolveRangeImpl('SheetB'),
    // which awaits formula cells that request the still-pending SheetA
    // range.  Detect this by tracking which sheets have an active range
    // resolution and falling back to XML-cached <v> values when a cycle
    // is detected.
    if (this.activeRangeSheets.has(targetSheetName)) {
      return this._resolveRangeFallback(targetSheetName, startRow, startCol, endRow, endCol);
    }

    // Coalesce identical in-flight resolveRange calls.
    // When 32 formulas concurrently request the same cross-sheet range (e.g.
    // 'Cars'!A:AAB), only the first actually resolves;
    // the rest await the same Promise.  This prevents the thundering-herd
    // pattern that created 32 × 70K-element arrays and caused >1s hang.
    const coalescingKey = `${targetSheetName}::${startRow}-${startCol}:${endRow}-${endCol}`;
    const pendingRange = this.pendingRanges.get(coalescingKey);
    if (pendingRange) return pendingRange;

    this.activeRangeSheets.add(targetSheetName);
    const promise = this._resolveRangeImpl(targetSheetName, startRow, startCol, endRow, endCol);
    this.pendingRanges.set(coalescingKey, promise);
    try {
      return await promise;
    } finally {
      this.activeRangeSheets.delete(targetSheetName);
      this.pendingRanges.delete(coalescingKey);
    }
  }

  private async _resolveRangeImpl(
    targetSheetName: string,
    startRow: number, startCol: number,
    endRow: number, endCol: number,
  ): Promise<any[]> {
    const targetSheetObj = this.sheetLookup.get(targetSheetName);

    const info: SheetCacheEntry = targetSheetObj
      ? await this.buildCellMapForSheet(targetSheetObj)
      : { maxRow: 0, maxCol: 0 } as any;
    const cmap: Map<string, Element> = info.cellMap || new Map();
    const sheetModel = info.sheetModel ?? null;
    const coveredMap = sheetModel?.coveredMap;
    const hasMerges = !!(coveredMap && coveredMap.size > 0);

    // Ensure values cache
    if (!this.sheetCache.has('__values')) this.sheetCache.set('__values', new Map());
    const valuesCache: Map<string, any> = this.sheetCache.get('__values');

    // Clamp loop bounds to the sheet's actual data extent.
    // For whole-column ranges (e.g. A:AAB), the caller may request up to 5000
    // rows, but the sheet may only have 46 populated rows.  Iterating beyond
    // the data extent creates millions of string-key allocations for cells that
    // are guaranteed to be empty, causing enormous GC pressure (35% of a 17s
    // hang in the trace).
    const effectiveEndRow = (info.maxRow && info.maxRow > 0) ? Math.min(endRow, info.maxRow) : endRow;
    const effectiveEndCol = (info.maxCol && info.maxCol > 0) ? Math.min(endCol, info.maxCol) : endCol;

    const totalCells = (endRow - startRow + 1) * (endCol - startCol + 1);
    const vals = new Array(totalCells);
    // Pre-fill with empty strings; cells beyond the data extent are empty by
    // definition and never need Map lookups or string-key construction.
    if (effectiveEndRow < endRow || effectiveEndCol < endCol) {
      vals.fill('');
    }
    const needsAsync: { vi: number; row: number; col: number; key: string; cellEl: Element; cellRef: string }[] = [];

    let vi = 0;
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        // Skip cells outside the sheet's actual data bounds (already pre-filled
        // with '' above).  This avoids millions of Map lookups and string
        // allocations for cells that can never contain data.
        if (r > effectiveEndRow || c > effectiveEndCol) {
          vi++;
          continue;
        }
        let row = r;
        let col = c;

        // Merged-cell remap
        if (hasMerges) {
          const anchorKey = coveredMap!.get(`${row}-${col}`);
          if (anchorKey) {
            const dashIdx = anchorKey.indexOf('-');
            row = parseInt(anchorKey.slice(0, dashIdx), 10);
            col = parseInt(anchorKey.slice(dashIdx + 1), 10);
          }
        }

        const numKey = `${row}-${col}`;
        const key = `${targetSheetName}::${numKey}`;

        // Fast-path: cached
        const cachedVal = valuesCache.get(key);
        if (cachedVal !== undefined) {
          vals[vi] = cachedVal;
          vi++;
          continue;
        }

        const cellEl = cmap.get(numKey);
        if (!cellEl) {
          valuesCache.set(key, '');
          vals[vi] = '';
          vi++;
          continue;
        }

        const cellRef = CellReference.toA1(row, col);
        const formulaText = CellEvaluator.getCellFormula(cellEl, cellRef, info.sheetModel as any);

        if (formulaText !== null && this.formulaEngine) {
          needsAsync.push({ vi, row, col, key, cellEl, cellRef });
        } else {
          const val = SheetParser.extractCellValue(cellEl, this.sharedStrings);
          valuesCache.set(key, val);
          vals[vi] = val;
        }
        vi++;
      }
    }

    // Resolve formula cells in small batches to avoid flooding the microtask
    // queue with thousands of concurrent async chains (which block the main
    // thread for >1s and prevent the browser from painting).
    if (needsAsync.length > 0) {
      const BATCH = 64;
      for (let i = 0; i < needsAsync.length; i += BATCH) {
        const batch = needsAsync.slice(i, i + BATCH);
        await Promise.all(batch.map(async (item) => {
          const cachedVal = valuesCache.get(item.key);
          if (cachedVal !== undefined) {
            vals[item.vi] = cachedVal;
            return;
          }
          const ref = targetSheetName !== this.sheet.name
            ? `${targetSheetName}!${item.cellRef}` : item.cellRef;
          const val = await this.evaluateCellByRef(ref, targetSheetName);
          vals[item.vi] = val;
        }));
      }
    }

    return vals;
  }

  /**
   * Fallback range resolver used when a circular range dependency is detected.
   * Returns cell values from the values cache or the XML-cached <v> elements
   * without evaluating any formulas, which breaks the deadlock cycle.
   */
  private async _resolveRangeFallback(
    targetSheetName: string,
    startRow: number, startCol: number,
    endRow: number, endCol: number,
  ): Promise<any[]> {
    const targetSheetObj = this.sheetLookup.get(targetSheetName);
    const info: SheetCacheEntry = targetSheetObj
      ? await this.buildCellMapForSheet(targetSheetObj)
      : { maxRow: 0, maxCol: 0 } as SheetCacheEntry;
    const cmap: Map<string, Element> = info.cellMap || new Map();
    const coveredMap = info.sheetModel?.coveredMap;
    const hasMerges = !!(coveredMap && coveredMap.size > 0);

    const valuesCache: Map<string, any> = this.sheetCache.get('__values') || new Map();

    const effectiveEndRow = (info.maxRow && info.maxRow > 0) ? Math.min(endRow, info.maxRow) : endRow;
    const effectiveEndCol = (info.maxCol && info.maxCol > 0) ? Math.min(endCol, info.maxCol) : endCol;

    const totalCells = (endRow - startRow + 1) * (endCol - startCol + 1);
    const vals = new Array(totalCells);
    if (effectiveEndRow < endRow || effectiveEndCol < endCol) vals.fill('');

    let vi = 0;
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        if (r > effectiveEndRow || c > effectiveEndCol) { vi++; continue; }

        let row = r;
        let col = c;
        if (hasMerges) {
          const anchorKey = coveredMap!.get(`${row}-${col}`);
          if (anchorKey) {
            const dashIdx = anchorKey.indexOf('-');
            row = parseInt(anchorKey.slice(0, dashIdx), 10);
            col = parseInt(anchorKey.slice(dashIdx + 1), 10);
          }
        }

        const key = `${targetSheetName}::${row}-${col}`;
        const cachedVal = valuesCache.get(key);
        if (cachedVal !== undefined) { vals[vi] = cachedVal; vi++; continue; }

        const cellEl = cmap.get(`${row}-${col}`);
        if (!cellEl) { vals[vi] = ''; vi++; continue; }

        // Extract the XML-cached value (no formula evaluation)
        const val = SheetParser.extractCellValue(cellEl, this.sharedStrings);
        valuesCache.set(key, val);
        vals[vi] = val;
        vi++;
      }
    }
    return vals;
  }

  /**
   * Fast inline A1 parser.  Avoids the overhead of CellReference.parse() + Map
   * look-up on every cell in a potentially 100 K-element range batch.
   * Returns { row, col } with 1-based indices, or { row: 0, col: 0 } on failure.
   */
  private static readonly _fastResult = { row: 0, col: 0 };
  private static fastParseA1(a1: string): { row: number; col: number } {
    let col = 0;
    let i = 0;
    const len = a1.length;
    // Skip leading '$' (absolute column)
    if (i < len && a1.charCodeAt(i) === 36) i++;
    // Parse column letters
    while (i < len) {
      const ch = a1.charCodeAt(i);
      if (ch >= 65 && ch <= 90) { col = col * 26 + (ch - 64); i++; } // A-Z
      else if (ch >= 97 && ch <= 122) { col = col * 26 + (ch - 96); i++; } // a-z
      else break;
    }
    if (col === 0) { CellEvaluator._fastResult.row = 0; CellEvaluator._fastResult.col = 0; return CellEvaluator._fastResult; }
    // Skip '$' before row digits
    if (i < len && a1.charCodeAt(i) === 36) i++;
    // Parse row digits
    let row = 0;
    while (i < len) {
      const ch = a1.charCodeAt(i);
      if (ch >= 48 && ch <= 57) { row = row * 10 + (ch - 48); i++; }
      else break;
    }
    CellEvaluator._fastResult.row = row;
    CellEvaluator._fastResult.col = col;
    return CellEvaluator._fastResult;
  }

  /**
   * Efficient batch cell resolution. Groups refs by sheet, resolves non-formula
   * cells synchronously from the cached cellMap, and only goes async for formula cells.
   * This avoids creating a Promise per cell in large ranges (e.g., 10K cells).
   *
   * Performance-critical: uses inline A1 parsing and numeric keys to avoid
   * the overhead of CellReference.parse / Map cache thrashing on large ranges.
   */
  async evaluateCellsBatch(refs: string[], contextSheetName?: string): Promise<any[]> {
    if (refs.length === 0) return [];

    // Ensure values cache exists
    if (!this.sheetCache.has('__values')) this.sheetCache.set('__values', new Map());
    const valuesCache: Map<string, any> = this.sheetCache.get('__values');

    // Group refs by sheet to minimize buildCellMapForSheet calls.
    // Inlined splitRef logic to avoid allocating {sheetName, a1} per ref.
    const groups = new Map<string, { indices: number[]; a1s: string[] }>();
    const results = new Array(refs.length);
    const needsAsync: { index: number; key: string; cellEl: Element; cellRef: string; formulaText: string; sheetName: string; info: SheetCacheEntry }[] = [];
    const defaultSheet = contextSheetName || this.sheet.name;

    for (let i = 0; i < refs.length; i++) {
      const ref = refs[i];
      let sheetName: string;
      let a1: string;
      const bangIdx = ref.lastIndexOf('!');
      if (bangIdx !== -1) {
        const rawSheet = ref.slice(0, bangIdx);
        a1 = ref.slice(bangIdx + 1);
        sheetName = (rawSheet.length >= 2 && rawSheet.charCodeAt(0) === 39 && rawSheet.charCodeAt(rawSheet.length - 1) === 39)
          ? rawSheet.slice(1, -1) : rawSheet;
      } else {
        sheetName = defaultSheet;
        a1 = ref;
      }
      let group = groups.get(sheetName);
      if (!group) { group = { indices: [], a1s: [] }; groups.set(sheetName, group); }
      group.indices.push(i);
      group.a1s.push(a1);
    }

    // Resolve each group
    for (const [sheetName, group] of groups) {
      // Find sheet object via pre-built lookup (O(1) instead of O(n))
      const targetSheetObj: SheetInfo | undefined = this.sheetLookup.get(sheetName);

      const info: SheetCacheEntry = targetSheetObj
        ? await this.buildCellMapForSheet(targetSheetObj)
        : { maxRow: 0, maxCol: 0 } as any;
      const cmap: Map<string, Element> = info.cellMap || new Map();
      const sheetModel = info.sheetModel ?? null;
      const coveredMap = sheetModel?.coveredMap;
      const hasMerges = !!(coveredMap && coveredMap.size > 0);

      for (let j = 0; j < group.indices.length; j++) {
        const idx = group.indices[j];
        const a1 = group.a1s[j];

        // --- Inline fast parse: avoid CellReference.parse() + Map overhead ---
        let { row, col } = CellEvaluator.fastParseA1(a1);
        if (!col || !row) { results[idx] = ''; continue; }

        // Merged-cell remap using numeric key directly (no string re-parse)
        if (hasMerges) {
          const anchorKey = coveredMap!.get(`${row}-${col}`);
          if (anchorKey) {
            const dashIdx = anchorKey.indexOf('-');
            row = parseInt(anchorKey.slice(0, dashIdx), 10);
            col = parseInt(anchorKey.slice(dashIdx + 1), 10);
          }
        }

        const numKey = `${row}-${col}`;
        const key = `${sheetName}::${numKey}`;

        // Fast-path: value already cached
        const cachedVal = valuesCache.get(key);
        if (cachedVal !== undefined) {
          results[idx] = cachedVal;
          continue;
        }

        const cellEl = cmap.get(numKey);
        if (!cellEl) {
          valuesCache.set(key, '');
          results[idx] = '';
          continue;
        }

        const cellRef = CellReference.toA1(row, col);
        const formulaText = CellEvaluator.getCellFormula(cellEl, cellRef, info.sheetModel as any);

        if (formulaText !== null && this.formulaEngine) {
          // Formula cell — must evaluate async (deferred)
          needsAsync.push({ index: idx, key, cellEl, cellRef, formulaText, sheetName, info });
        } else {
          // Plain value cell — synchronous extraction (no Promise overhead)
          const val = SheetParser.extractCellValue(cellEl, this.sharedStrings);
          valuesCache.set(key, val);
          results[idx] = val;
        }
      }
    }

    // Evaluate formula cells (typically a small fraction of a large range)
    if (needsAsync.length > 0) {
      // Process all formula cells concurrently in one batch.
      // The pendingCells Map already coalesces duplicate evaluations, and
      // the valuesCache prevents redundant work, so a single Promise.all is
      // safe and avoids chunking overhead that previously generated thousands
      // of microtasks.
      await Promise.all(needsAsync.map(async (item) => {
        // Check if another concurrent path already resolved it
        const cachedVal = valuesCache.get(item.key);
        if (cachedVal !== undefined) {
          results[item.index] = cachedVal;
          return;
        }
        const val = await this.evaluateCellByRef(
          item.sheetName !== this.sheet.name ? `${item.sheetName}!${item.cellRef}` : item.cellRef,
          item.sheetName
        );
        results[item.index] = val;
      }));
    }

    return results;
  }
}
