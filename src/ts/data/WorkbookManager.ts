// WorkbookManager.ts - Central workbook state and background sheet loading
//
// Manages the in-memory state of an opened workbook:
// - CellStores per sheet
// - Shared DependencyGraph
// - RecalcEngine
// - Background loading of non-active sheets
//
// Designed to be created once per file open and passed to renderer/editor.

import type JSZip from 'jszip';
import { CellReference } from '../core/CellReference';
import { XmlParser } from '../core/XmlParser';
import { SheetParser } from '../workbook/SheetParser';
import { CellStore } from './CellStore';
import { DependencyGraph } from './DependencyGraph';
import { RecalcEngine } from './RecalcEngine';
import { CellEvaluator } from '../renderer/CellEvaluator';
import type { SheetInfo, SheetModel, IFormulaEngine, SheetCacheEntry } from '../types';

export interface WorkbookManagerOptions {
  zip: JSZip;
  sheets: SheetInfo[];
  sharedStrings: string[];
  formulaEngine: IFormulaEngine;
}

/**
 * Sheet load status for tracking background loading.
 */
export type SheetLoadStatus = 'pending' | 'loading' | 'loaded' | 'error';

/**
 * WorkbookManager is the central coordinator for in-memory workbook state.
 */
export class WorkbookManager {
  readonly zip: JSZip;
  readonly sheets: SheetInfo[];
  readonly sharedStrings: string[];
  readonly formulaEngine: IFormulaEngine;

  /** CellStores per sheet, keyed by sheet name */
  readonly cellStores: Map<string, CellStore> = new Map();

  /** Sheet models (merge info, shared formulas) per sheet */
  readonly sheetModels: Map<string, SheetModel> = new Map();

  /** The shared dependency graph */
  readonly graph: DependencyGraph = new DependencyGraph();

  /** The recalculation engine */
  readonly recalcEngine: RecalcEngine;

  /** Sheet loading status */
  private loadStatus: Map<string, SheetLoadStatus> = new Map();

  /** Callbacks for when a sheet finishes loading */
  private loadCallbacks: Map<string, Array<(store: CellStore) => void>> = new Map();

  /** Active sheet name */
  private _activeSheet: string = '';

  /** sheetCache for CellEvaluator compatibility */
  readonly sheetCache: Map<string, any> = new Map();

  /** DOM update callback, set by the editor */
  onCellUpdated?: (sheetName: string, row: number, col: number, value: any) => void;

  /** Whether the dependency graph has been built for all loaded sheets */
  private _graphReady = false;
  /** Promise for deferred graph build (resolves when complete) */
  private _graphBuildPromise: Promise<void> | null = null;

  /** Guard: whether loadRemainingSheets is currently running */
  private _loadingRemaining = false;

  /** Promise that resolves when loadRemainingSheets finishes */
  private _loadRemainingPromise: Promise<void> | null = null;

  constructor(options: WorkbookManagerOptions) {
    this.zip = options.zip;
    this.sheets = options.sheets;
    this.sharedStrings = options.sharedStrings;
    this.formulaEngine = options.formulaEngine;

    this.recalcEngine = new RecalcEngine({
      cellStores: this.cellStores,
      graph: this.graph,
      formulaEngine: this.formulaEngine,
      sharedStrings: this.sharedStrings,
      onCellUpdated: (sheet, row, col, value) => {
        this.onCellUpdated?.(sheet, row, col, value);
      },
    });

    // Initialize load status
    for (const sheet of this.sheets) {
      this.loadStatus.set(sheet.name, 'pending');
    }
  }

  get activeSheet(): string { return this._activeSheet; }

  /**
   * Load a single sheet's data into the CellStore.
   * This is the core loading method — it reads the XML and populates the store.
   */
  async loadSheet(sheet: SheetInfo): Promise<CellStore> {
    if (!sheet.target) throw new Error(`Sheet "${sheet.name}" has no target path`);

    // Return existing if already loaded
    const existing = this.cellStores.get(sheet.name);
    if (existing && this.loadStatus.get(sheet.name) === 'loaded') return existing;

    this.loadStatus.set(sheet.name, 'loading');

    try {
      const sheetDoc = await XmlParser.readZipXml(this.zip, sheet.target);
      const sheetModel = SheetParser.buildSheetModel(sheetDoc);
      const sheetData = sheetDoc.getElementsByTagName('sheetData')[0];

      this.sheetModels.set(sheet.name, sheetModel);

      const store = new CellStore(sheet.name);
      this.cellStores.set(sheet.name, store);

      if (sheetData) {
        const rows = Array.from(sheetData.getElementsByTagName('row'));
        let maxRow = 0;

        rows.forEach((row) => {
          const rowIndex = parseInt(row.getAttribute('r') || '0', 10);
          const actualRow = Number.isNaN(rowIndex) || rowIndex === 0 ? maxRow + 1 : rowIndex;
          maxRow = Math.max(maxRow, actualRow);

          const cells = Array.from(row.getElementsByTagName('c'));
          cells.forEach((cellEl) => {
            const ref = cellEl.getAttribute('r') || '';
            const parsed = CellReference.parse(ref);
            if (parsed.col === 0) return;

            // Extract formula
            const cellRef = CellReference.toA1(actualRow, parsed.col);
            const formulaText = CellEvaluator.getCellFormula(cellEl, cellRef, sheetModel);

            // Extract value
            const value = SheetParser.extractCellValue(cellEl, this.sharedStrings);

            // Get style index (convert to number to save ~20 bytes/cell vs string)
            const sAttr = cellEl.getAttribute('s');
            const styleIndex = sAttr != null ? parseInt(sAttr, 10) : undefined;

            // Determine type from cell element
            const cellType = cellEl.getAttribute('t');
            let dataType: 'string' | 'number' | 'boolean' | undefined;
            if (cellType === 's' || cellType === 'inlineStr') dataType = 'string';
            else if (cellType === 'b') dataType = 'boolean';

            store.load(actualRow, parsed.col, value, formulaText, {
              styleIndex,
              type: dataType,
            });
          });
        });

        // Store lightweight sheetCache entry (cellMap and sheetDoc are NOT retained
        // — the CellStore already has all extracted data, and SheetRenderer builds
        // its own local cache for CellEvaluator during rendering).
        this.sheetCache.set(sheet.name, {
          maxRow: store.maxRow,
          maxCol: store.maxCol,
          sheetModel,
        });
      }

      this.loadStatus.set(sheet.name, 'loaded');

      // Dependency graph building is deferred — it will be built in
      // the background after the initial render for much faster startup.
      // See loadActiveSheet() → _buildGraphDeferred().

      // Notify any waiting callbacks
      const callbacks = this.loadCallbacks.get(sheet.name);
      if (callbacks) {
        callbacks.forEach((cb) => cb(store));
        this.loadCallbacks.delete(sheet.name);
      }

      return store;
    } catch (err) {
      this.loadStatus.set(sheet.name, 'error');
      throw err;
    }
  }

  /**
   * Load the active (first-displayed) sheet and start background loading of others.
   * Background loading and graph building are deferred until after the caller
   * has had a chance to render (the returned promise resolves immediately after
   * the active sheet is loaded — background work starts via setTimeout so it
   * does not compete with the initial render).
   */
  async loadActiveSheet(sheet: SheetInfo): Promise<CellStore> {
    this._activeSheet = sheet.name;
    const store = await this.loadSheet(sheet);

    // Schedule background loading AFTER the current microtask queue drains,
    // so the renderer can paint the active sheet first.
    // The dependency graph build is chained AFTER all sheets finish loading
    // to avoid partially-built graphs and wasted work.
    setTimeout(() => {
      this._loadRemainingPromise = this.loadRemainingSheets(sheet.name).then(() => {
        this._scheduleDeferredGraphBuild();
      });
    }, 50);

    return store;
  }

  /**
   * Set the active sheet (when user clicks a tab).
   */
  setActiveSheet(sheetName: string): void {
    this._activeSheet = sheetName;
  }

  /**
   * Ensure a sheet has an in-memory store and cache entry.
   * Used for locally-created sheets with no XML target.
   */
  ensureSheetStore(sheetName: string): CellStore {
    const existing = this.cellStores.get(sheetName);
    if (existing) return existing;

    const store = new CellStore(sheetName);
    this.cellStores.set(sheetName, store);
    this.sheetModels.set(sheetName, {
      mergedRanges: [],
      anchorMap: new Map(),
      coveredMap: new Map(),
      sharedFormulas: new Map(),
      sharedCells: new Map(),
    });
    this.sheetCache.set(sheetName, {
      maxRow: 0,
      maxCol: 0,
      sheetModel: this.sheetModels.get(sheetName),
    });
    this.loadStatus.set(sheetName, 'loaded');
    return store;
  }

  /**
   * Remove a sheet from in-memory workbook state.
   */
  removeSheet(sheetName: string): void {
    const index = this.sheets.findIndex((s) => s.name === sheetName);
    if (index >= 0) this.sheets.splice(index, 1);

    this.cellStores.delete(sheetName);
    this.sheetModels.delete(sheetName);
    this.sheetCache.delete(sheetName);
    this.loadStatus.delete(sheetName);
    this.loadCallbacks.delete(sheetName);

    if (this._activeSheet === sheetName) {
      this._activeSheet = this.sheets[0]?.name || '';
    }

    // Only rebuild the graph if it was previously ready.
    // If the initial background load is still running, the graph will be
    // built once it completes — scheduling another build here would cause
    // contention and redundant work.
    if (this._graphReady) {
      this._graphReady = false;
      this._graphBuildPromise = null;
      this._scheduleDeferredGraphBuild();
    }
  }

  /**
   * Rename a sheet in workbook metadata and in-memory stores.
   */
  renameSheet(oldName: string, newName: string): void {
    const trimmed = newName.trim();
    if (!trimmed || oldName === trimmed) return;
    if (this.sheets.some((s) => s.name === trimmed)) {
      throw new Error(`A sheet named "${trimmed}" already exists.`);
    }

    const sheet = this.sheets.find((s) => s.name === oldName);
    if (!sheet) return;
    sheet.name = trimmed;

    if (this.cellStores.has(oldName)) {
      const store = this.cellStores.get(oldName)!;
      this.cellStores.delete(oldName);
      this.cellStores.set(trimmed, store);
    }

    if (this.sheetModels.has(oldName)) {
      const model = this.sheetModels.get(oldName)!;
      this.sheetModels.delete(oldName);
      this.sheetModels.set(trimmed, model);
    }

    if (this.sheetCache.has(oldName)) {
      const cache = this.sheetCache.get(oldName);
      this.sheetCache.delete(oldName);
      this.sheetCache.set(trimmed, cache);
    }

    if (this.loadStatus.has(oldName)) {
      const status = this.loadStatus.get(oldName)!;
      this.loadStatus.delete(oldName);
      this.loadStatus.set(trimmed, status);
    }

    if (this.loadCallbacks.has(oldName)) {
      const callbacks = this.loadCallbacks.get(oldName)!;
      this.loadCallbacks.delete(oldName);
      this.loadCallbacks.set(trimmed, callbacks);
    }

    if (this._activeSheet === oldName) {
      this._activeSheet = trimmed;
    }

    // Only rebuild the graph if it was previously ready.
    if (this._graphReady) {
      this._graphReady = false;
      this._graphBuildPromise = null;
      this._scheduleDeferredGraphBuild();
    }
  }

  /**
   * Load remaining sheets in the background.
   * Guarded so only one invocation runs at a time — subsequent calls
   * return the existing promise if one is in flight.
   */
  private async loadRemainingSheets(excludeSheet: string): Promise<void> {
    if (this._loadingRemaining) {
      // Already running — wait for the existing run to finish
      if (this._loadRemainingPromise) return this._loadRemainingPromise;
      return;
    }
    this._loadingRemaining = true;

    try {
      for (const sheet of this.sheets) {
        if (sheet.name === excludeSheet) continue;
        if (this.loadStatus.get(sheet.name) !== 'pending') continue;

        try {
          await this.loadSheet(sheet);
        } catch (err) {
          console.warn(`Background loading of sheet "${sheet.name}" failed:`, err);
        }

        // Yield to event loop between sheets so the UI stays responsive.
        // Use a slightly longer delay (25ms) to give the renderer enough
        // breathing room instead of flooding the task queue with
        // back-to-back setTimeout(0) tasks.
        await new Promise<void>((resolve) => setTimeout(resolve, 25));
      }
    } finally {
      this._loadingRemaining = false;
    }
  }

  /**
   * Schedule the deferred dependency graph build.
   * Called after all sheets have finished loading so the graph is built once,
   * not incrementally while sheets are still being decompressed.
   */
  private _scheduleDeferredGraphBuild(): void {
    if (this._graphBuildPromise) return; // already scheduled
    this._graphBuildPromise = new Promise<void>((resolve) => {
      // Use requestIdleCallback when available so graph building only runs
      // when the browser is idle; fall back to setTimeout(…, 100) to give
      // the renderer plenty of time to finish.
      const schedule = typeof requestIdleCallback === 'function'
        ? (fn: () => void) => requestIdleCallback(fn, { timeout: 2000 })
        : (fn: () => void) => setTimeout(fn, 100);

      schedule(async () => {
        try {
          await this._buildGraphDeferred();
        } catch (err) {
          console.warn('Deferred graph build failed:', err);
        }
        resolve();
      });
    });
  }

  /**
   * Build the dependency graph for all loaded sheets, yielding between sheets
   * to avoid blocking the UI.
   */
  private async _buildGraphDeferred(): Promise<void> {
    this.graph.clear();
    for (const [sheetName] of this.cellStores) {
      await this.recalcEngine.buildSheetGraphDeferred(sheetName);
    }
    this._graphReady = true;
  }

  /**
   * Ensure the dependency graph is built (blocks until complete).
   * Called by operations that need the graph (e.g. cell edits).
   */
  async ensureGraphReady(): Promise<void> {
    if (this._graphReady) return;
    if (this._graphBuildPromise) {
      await this._graphBuildPromise;
      return;
    }
    // No build scheduled yet — build synchronously
    this.recalcEngine.buildFullGraph();
    this._graphReady = true;
  }

  /**
   * Check if a sheet is loaded.
   */
  isSheetLoaded(sheetName: string): boolean {
    return this.loadStatus.get(sheetName) === 'loaded';
  }

  /**
   * Get the load status of a sheet.
   */
  getSheetStatus(sheetName: string): SheetLoadStatus {
    return this.loadStatus.get(sheetName) || 'pending';
  }

  /**
   * Wait for a sheet to be loaded.
   */
  async waitForSheet(sheetName: string): Promise<CellStore> {
    const existing = this.cellStores.get(sheetName);
    if (existing && this.loadStatus.get(sheetName) === 'loaded') return existing;

    return new Promise<CellStore>((resolve) => {
      if (!this.loadCallbacks.has(sheetName)) this.loadCallbacks.set(sheetName, []);
      this.loadCallbacks.get(sheetName)!.push(resolve);
    });
  }

  /**
   * Get a CellStore, loading the sheet if necessary.
   */
  async getOrLoadStore(sheetName: string): Promise<CellStore | undefined> {
    const existing = this.cellStores.get(sheetName);
    if (existing) return existing;

    const sheet = this.sheets.find((s) => s.name === sheetName);
    if (!sheet) return undefined;

    return this.loadSheet(sheet);
  }

  /**
   * Resolve a cell reference to its value from the in-memory stores.
   * Used as a fast resolveCell callback for formula evaluation.
   */
  resolveCellValue(ref: string, defaultSheet?: string): any {
    let targetSheet = defaultSheet || this._activeSheet;
    let a1 = ref;

    if (ref.includes('!')) {
      const parts = ref.split('!');
      targetSheet = parts.slice(0, -1).join('!').replace(/^'+|'+$/g, '');
      a1 = parts[parts.length - 1];
    }

    const store = this.cellStores.get(targetSheet);
    if (!store) return '';

    const parsed = CellReference.parse(a1);
    if (!parsed.row || !parsed.col) return '';

    return store.getValue(parsed.row, parsed.col);
  }
}
