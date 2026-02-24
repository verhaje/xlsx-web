// CellStore.ts - In-memory cell data model
//
// Stores cell values, formulas, and types for all sheets. Designed to be the
// single source of truth for cell data, replacing XML element lookups after
// initial load. Future-proof: contains enough metadata to reconstruct XLSX
// sheet XML for saving.

import { CellReference } from '../core/CellReference';
import type { CellRef } from '../types';

// ---- Cell Data Types ----

/**
 * The type of data stored in a cell.
 */
export type CellDataType = 'string' | 'number' | 'boolean' | 'formula' | 'date' | 'error' | 'empty';

/**
 * Represents a single cell's data in memory.
 *
 * Memory optimization notes:
 * - `styleIndex` is stored as number|undefined (was string) — saves ~20 bytes per cell
 * - `dirty` is not stored here; dirty cells are tracked in a separate Set in CellStore
 * - Optional fields (formula, numberFormat) are omitted when not needed
 */
export interface CellData {
  /** The raw value (for formulas, this is the computed result). */
  value: any;
  /** Formula text (without leading '='), null if not a formula cell. */
  formula: string | null;
  /** The type of data. */
  type: CellDataType;
  /** Original style index from the XLSX ('s' attribute on <c>), as a number. */
  styleIndex?: number;
  /** Number format string (e.g. '#,##0.00', 'yyyy-mm-dd') for future XLSX save. */
  numberFormat?: string;
}

/**
 * CellStore manages all in-memory cell data for a single sheet.
 * Keys follow the project convention: `${row}-${col}` (1-based).
 */
export class CellStore {
  /** sheet name (for cross-sheet refs and save) */
  readonly sheetName: string;

  /** All cell data, keyed by "row-col" (1-based). */
  private cells: Map<string, CellData> = new Map();

  /** Track dirty cells separately to avoid a boolean on every CellData. */
  private _dirtyCells: Set<string> = new Set();

  /** Track formula cell keys for fast getFormulaCells() without full scan. */
  private _formulaKeys: Set<string> = new Set();

  /** Track max extents for iteration. */
  private _maxRow = 0;
  private _maxCol = 0;

  constructor(sheetName: string) {
    this.sheetName = sheetName;
  }

  // ---- Accessors ----

  get maxRow(): number { return this._maxRow; }
  get maxCol(): number { return this._maxCol; }

  /**
   * Get cell data by row-col key.
   */
  getByKey(key: string): CellData | undefined {
    return this.cells.get(key);
  }

  /**
   * Get cell data by row and col (1-based).
   */
  get(row: number, col: number): CellData | undefined {
    return this.cells.get(CellReference.cellKey(row, col));
  }

  /**
   * Get cell data by A1-style reference.
   */
  getByRef(a1: string): CellData | undefined {
    const ref = CellReference.parse(a1);
    if (!ref.row || !ref.col) return undefined;
    return this.get(ref.row, ref.col);
  }

  /**
   * Check if a cell exists.
   */
  has(row: number, col: number): boolean {
    return this.cells.has(CellReference.cellKey(row, col));
  }

  /**
   * Get the computed value of a cell. Returns '' for missing cells.
   */
  getValue(row: number, col: number): any {
    const cell = this.get(row, col);
    return cell ? cell.value : '';
  }

  /**
   * Get the formula of a cell, or null if it's not a formula cell.
   */
  getFormula(row: number, col: number): string | null {
    const cell = this.get(row, col);
    return cell ? cell.formula : null;
  }

  // ---- Mutators ----

  /**
   * Set a cell's data. Automatically detects formulas (starting with '=').
   * Returns the previous CellData (or undefined) for undo support.
   */
  set(row: number, col: number, input: string | number | boolean, options?: {
    styleIndex?: number;
    numberFormat?: string;
  }): CellData | undefined {
    const key = CellReference.cellKey(row, col);
    const prev = this.cells.get(key);
    this._maxRow = Math.max(this._maxRow, row);
    this._maxCol = Math.max(this._maxCol, col);

    const cellData = CellStore.createCellData(input, options);
    this._dirtyCells.add(key);
    if (cellData.formula) {
      this._formulaKeys.add(key);
    } else {
      this._formulaKeys.delete(key);
    }
    this.cells.set(key, cellData);
    return prev;
  }

  /**
   * Set a cell's computed value (after formula evaluation).
   * Does NOT mark the cell dirty — used by the recalc engine.
   */
  setComputedValue(row: number, col: number, value: any): void {
    const key = CellReference.cellKey(row, col);
    const cell = this.cells.get(key);
    if (cell) {
      cell.value = value;
    }
  }

  /**
   * Set a formula cell. The formula text should NOT include the leading '='.
   */
  setFormula(row: number, col: number, formulaText: string, computedValue?: any): CellData | undefined {
    const key = CellReference.cellKey(row, col);
    const prev = this.cells.get(key);
    this._maxRow = Math.max(this._maxRow, row);
    this._maxCol = Math.max(this._maxCol, col);

    this.cells.set(key, {
      value: computedValue ?? '',
      formula: formulaText,
      type: 'formula',
    });
    this._dirtyCells.add(key);
    this._formulaKeys.add(key);
    return prev;
  }

  /**
   * Bulk-load cell data from parsed sheet (used during initial load).
   * Does NOT mark cells as dirty.
   */
  load(row: number, col: number, value: any, formula: string | null, options?: {
    styleIndex?: number;
    type?: CellDataType;
  }): void {
    const key = CellReference.cellKey(row, col);
    this._maxRow = Math.max(this._maxRow, row);
    this._maxCol = Math.max(this._maxCol, col);

    let type: CellDataType = options?.type || 'empty';
    if (formula) {
      type = 'formula';
      this._formulaKeys.add(key);
    } else if (type === 'empty') {
      type = CellStore.inferType(value);
    }

    const cell: CellData = {
      value,
      formula,
      type,
    };
    if (options?.styleIndex !== undefined) cell.styleIndex = options.styleIndex;
    this.cells.set(key, cell);
  }

  /**
   * Delete a cell.
   */
  delete(row: number, col: number): CellData | undefined {
    const key = CellReference.cellKey(row, col);
    const prev = this.cells.get(key);
    this.cells.delete(key);
    this._dirtyCells.delete(key);
    this._formulaKeys.delete(key);
    return prev;
  }

  /**
   * Iterate all non-empty cells.
   */
  forEach(callback: (key: string, data: CellData, row: number, col: number) => void): void {
    this.cells.forEach((data, key) => {
      const [r, c] = key.split('-').map(Number);
      callback(key, data, r, c);
    });
  }

  /**
   * Get all formula cells (used for dependency graph building).
   */
  getFormulaCells(): Array<{ row: number; col: number; key: string; formula: string }> {
    const result: Array<{ row: number; col: number; key: string; formula: string }> = [];
    for (const key of this._formulaKeys) {
      const data = this.cells.get(key);
      if (data?.formula) {
        const dashIdx = key.indexOf('-');
        const r = parseInt(key.slice(0, dashIdx), 10);
        const c = parseInt(key.slice(dashIdx + 1), 10);
        result.push({ row: r, col: c, key, formula: data.formula });
      }
    }
    return result;
  }

  /**
   * Get all dirty cells (for future save functionality).
   */
  getDirtyCells(): Array<{ row: number; col: number; key: string; data: CellData }> {
    const result: Array<{ row: number; col: number; key: string; data: CellData }> = [];
    for (const key of this._dirtyCells) {
      const data = this.cells.get(key);
      if (data) {
        const [r, c] = key.split('-').map(Number);
        result.push({ row: r, col: c, key, data });
      }
    }
    return result;
  }

  /**
   * Check if a cell is dirty.
   */
  isDirty(row: number, col: number): boolean {
    return this._dirtyCells.has(CellReference.cellKey(row, col));
  }

  /**
   * Clear all dirty flags (e.g. after saving).
   */
  clearDirtyFlags(): void {
    this._dirtyCells.clear();
  }

  /**
   * Get the total number of cells.
   */
  get size(): number {
    return this.cells.size;
  }

  // ---- Static Helpers ----

  /**
   * Create CellData from a user input value.
   */
  static createCellData(input: string | number | boolean, options?: {
    styleIndex?: number;
    numberFormat?: string;
  }): CellData {
    if (typeof input === 'string' && input.startsWith('=')) {
      const cell: CellData = {
        value: '',
        formula: input.slice(1), // strip leading '='
        type: 'formula',
      };
      if (options?.styleIndex !== undefined) cell.styleIndex = options.styleIndex;
      if (options?.numberFormat) cell.numberFormat = options.numberFormat;
      return cell;
    }

    const type = CellStore.inferType(input);
    let value: any = input;

    // Attempt numeric coercion for string inputs that look numeric
    if (typeof input === 'string' && type === 'number') {
      value = Number(input);
    }

    const cell: CellData = {
      value,
      formula: null,
      type,
    };
    if (options?.styleIndex !== undefined) cell.styleIndex = options.styleIndex;
    if (options?.numberFormat) cell.numberFormat = options.numberFormat;
    return cell;
  }

  /**
   * Infer CellDataType from a runtime value.
   */
  static inferType(value: any): CellDataType {
    if (value === null || value === undefined || value === '') return 'empty';
    if (typeof value === 'number') return 'number';
    if (typeof value === 'boolean') return 'boolean';
    if (typeof value === 'string') {
      if (value.startsWith('#') && value.endsWith('!')) return 'error';
      const n = Number(value);
      if (!Number.isNaN(n) && value.trim() !== '') return 'number';
      return 'string';
    }
    return 'string';
  }
}
