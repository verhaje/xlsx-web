// RecalcEngine.ts - Recalculation of dependent cells when values change
//
// When a cell value changes, this engine:
// 1. Looks up all dependent cells via the DependencyGraph
// 2. Topologically sorts them
// 3. Re-evaluates each formula cell
// 4. Updates the CellStore and DOM

import { CellReference } from '../core/CellReference';
import { CellStore } from './CellStore';
import { DependencyGraph } from './DependencyGraph';
import type { IFormulaEngine, FormulaContext } from '../types';

export interface RecalcContext {
  /** All CellStores keyed by sheet name */
  cellStores: Map<string, CellStore>;
  /** The shared dependency graph */
  graph: DependencyGraph;
  /** The formula engine instance */
  formulaEngine: IFormulaEngine;
  /** Shared strings from the workbook (for resolveCell) */
  sharedStrings: string[];
  /** Callback to update a DOM cell after recalculation */
  onCellUpdated?: (sheetName: string, row: number, col: number, value: any) => void;
}

/**
 * RecalcEngine handles cascading recalculation when cells are edited.
 */
export class RecalcEngine {
  private ctx: RecalcContext;

  constructor(ctx: RecalcContext) {
    this.ctx = ctx;
  }

  /**
   * Update the context (e.g. when a new workbook is loaded).
   */
  updateContext(partial: Partial<RecalcContext>): void {
    Object.assign(this.ctx, partial);
  }

  /**
   * Set a cell value (not a formula) and trigger recalculation of dependents.
   */
  async setCellValue(sheetName: string, row: number, col: number, value: string | number | boolean): Promise<void> {
    const store = this.ctx.cellStores.get(sheetName);
    if (!store) return;

    // Update the cell
    store.set(row, col, value);

    // Clear this cell's formula dependencies (it's no longer a formula)
    this.ctx.graph.clearCell(sheetName, row, col);

    // Recalculate dependents
    await this.recalcDependents(sheetName, row, col);
  }

  /**
   * Set a cell formula and trigger recalculation.
   * The formulaText should include the leading '='.
   */
  async setCellFormula(sheetName: string, row: number, col: number, formulaText: string): Promise<any> {
    const store = this.ctx.cellStores.get(sheetName);
    if (!store) return '';

    // Strip the leading '='
    const bareFormula = formulaText.startsWith('=') ? formulaText.slice(1) : formulaText;

    // Update the formula in the store
    store.setFormula(row, col, bareFormula);

    // Update dependency graph
    this.ctx.graph.setFormulaDependencies(sheetName, row, col, bareFormula);

    // Evaluate this cell's formula
    const value = await this.evaluateCell(sheetName, row, col, bareFormula);
    store.setComputedValue(row, col, value);

    // Notify DOM update for this cell
    this.ctx.onCellUpdated?.(sheetName, row, col, value);

    // Recalculate dependents
    await this.recalcDependents(sheetName, row, col);

    return value;
  }

  /**
   * Recalculate all cells that depend on the given cell.
   */
  private async recalcDependents(sheetName: string, row: number, col: number): Promise<void> {
    const order = this.ctx.graph.getRecalcOrder(sheetName, row, col);

    for (const qualKey of order) {
      const { sheetName: depSheet, row: depRow, col: depCol } = DependencyGraph.parseQualifiedKey(qualKey);
      const depStore = this.ctx.cellStores.get(depSheet);
      if (!depStore) continue;

      const cellData = depStore.get(depRow, depCol);
      if (!cellData?.formula) continue;

      // Re-evaluate the formula
      const value = await this.evaluateCell(depSheet, depRow, depCol, cellData.formula);
      depStore.setComputedValue(depRow, depCol, value);

      // Notify DOM update
      this.ctx.onCellUpdated?.(depSheet, depRow, depCol, value);
    }
  }

  /**
   * Evaluate a single cell's formula.
   */
  private async evaluateCell(sheetName: string, row: number, col: number, formulaText: string): Promise<any> {
    const resolveCell = async (ref: string): Promise<any> => {
      return this.resolveCellRef(ref, sheetName);
    };

    try {
      return await this.ctx.formulaEngine.evaluateFormula(formulaText, {
        resolveCell,
        sharedStrings: this.ctx.sharedStrings,
      });
    } catch {
      return '#ERROR!';
    }
  }

  /**
   * Resolve a cell reference (possibly cross-sheet) to its value from CellStores.
   */
  private resolveCellRef(ref: string, defaultSheet: string): any {
    let targetSheet = defaultSheet;
    let a1 = ref;

    if (ref.includes('!')) {
      const parts = ref.split('!');
      targetSheet = parts.slice(0, -1).join('!').replace(/^'+|'+$/g, '');
      a1 = parts[parts.length - 1];
    }

    const store = this.ctx.cellStores.get(targetSheet);
    if (!store) return '';

    const parsed = CellReference.parse(a1);
    if (!parsed.row || !parsed.col) return '';

    return store.getValue(parsed.row, parsed.col);
  }

  /**
   * Build the dependency graph for all formula cells in all loaded sheets.
   * Should be called after initial sheet loading is complete.
   * Uses bulk loading (no per-cell clearCell) for much faster initial build.
   */
  buildFullGraph(): void {
    this.ctx.graph.clear();
    const allCells: Array<{ sheetName: string; row: number; col: number; formula: string }> = [];
    for (const [sheetName, store] of this.ctx.cellStores) {
      const formulaCells = store.getFormulaCells();
      for (const fc of formulaCells) {
        allCells.push({ sheetName, row: fc.row, col: fc.col, formula: fc.formula });
      }
    }
    this.ctx.graph.bulkSetFormulaDependencies(allCells);
  }

  /**
   * Build the dependency graph for a single sheet.
   * Uses bulk loading for efficient initial construction.
   */
  buildSheetGraph(sheetName: string): void {
    const store = this.ctx.cellStores.get(sheetName);
    if (!store) return;
    const formulaCells = store.getFormulaCells();
    const cells = formulaCells.map(fc => ({
      sheetName,
      row: fc.row,
      col: fc.col,
      formula: fc.formula,
    }));
    this.ctx.graph.bulkSetFormulaDependencies(cells);
  }

  /**
   * Build the dependency graph for a single sheet asynchronously,
   * yielding to the event loop periodically to avoid blocking the UI.
   * Suitable for background/deferred graph construction.
   */
  async buildSheetGraphDeferred(sheetName: string, chunkSize = 500): Promise<void> {
    const store = this.ctx.cellStores.get(sheetName);
    if (!store) return;
    const formulaCells = store.getFormulaCells();
    const total = formulaCells.length;

    for (let i = 0; i < total; i += chunkSize) {
      const end = Math.min(i + chunkSize, total);
      const chunk: Array<{ sheetName: string; row: number; col: number; formula: string }> = [];
      for (let j = i; j < end; j++) {
        const fc = formulaCells[j];
        chunk.push({ sheetName, row: fc.row, col: fc.col, formula: fc.formula });
      }
      this.ctx.graph.bulkSetFormulaDependencies(chunk);
      // Yield to event loop between chunks — use a longer delay (25ms)
      // to give the renderer and other UI work breathing room, instead of
      // setTimeout(0) which generates back-to-back microtask floods.
      if (end < total) {
        await new Promise<void>(resolve => setTimeout(resolve, 25));
      }
    }
  }
}
