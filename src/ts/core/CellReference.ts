// CellReference.ts - Cell reference parsing and conversion utilities

import type { CellRef, RangeRef } from '../types';

/**
 * Utility class for working with A1-style cell references.
 */
export class CellReference {
  /**
   * Parse an A1-style cell reference (e.g. "B3") into { row, col }.
   */
  static parse(ref: string): CellRef {
    const match = /([A-Z]+)(\d+)/i.exec(ref);
    if (!match) return { col: 0, row: 0 };
    const colLetters = match[1].toUpperCase();
    const row = parseInt(match[2], 10);
    return { col: CellReference.columnNameToIndex(colLetters), row };
  }

  /**
   * Convert a column letter string ("A", "AA", etc.) to a 1-based column index.
   */
  static columnNameToIndex(name: string): number {
    let index = 0;
    for (let i = 0; i < name.length; i += 1) {
      index = index * 26 + (name.charCodeAt(i) - 64);
    }
    return index;
  }

  /**
   * Convert a 1-based column index to a column letter string.
   */
  static columnIndexToName(index: number): string {
    let result = '';
    let current = index;
    while (current > 0) {
      const remainder = (current - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      current = Math.floor((current - 1) / 26);
    }
    return result;
  }

  /**
   * Expand an A1 range string (e.g. "A1:C3") into an array of CellRef objects.
   */
  static expandRange(range: string): CellRef[] {
    const [start, end] = range.split(':');
    const s = CellReference.parse(start);
    const e = end ? CellReference.parse(end) : s;
    const cells: CellRef[] = [];
    for (let r = s.row; r <= e.row; r++) {
      for (let c = s.col; c <= e.col; c++) {
        cells.push({ row: r, col: c });
      }
    }
    return cells;
  }

  /**
   * Parse a space-separated sqref string into an array of CellRef objects.
   */
  static parseSqref(sqref: string): CellRef[] {
    const parts = sqref.trim().split(/\s+/);
    const cells: CellRef[] = [];
    parts.forEach((p) => {
      cells.push(...CellReference.expandRange(p));
    });
    return cells;
  }

  /**
   * Parse a range string into start and end CellRef.
   */
  static parseRangeRef(range: string): RangeRef {
    const [startRef, endRef] = range.split(':');
    const start = CellReference.parse(startRef);
    const end = endRef ? CellReference.parse(endRef) : start;
    return { start, end };
  }

  /**
   * Build a cell key string from row and col (1-based).
   */
  static cellKey(row: number, col: number): string {
    return `${row}-${col}`;
  }

  /**
   * Build an A1-style reference string from row and col (1-based).
   */
  static toA1(row: number, col: number): string {
    return `${CellReference.columnIndexToName(col)}${row}`;
  }

  /**
   * Map a merged-cell reference to its anchor cell reference.
   */
  static mapMergedRef(refA1: string, sheetModel: { coveredMap?: Map<string, string> } | null): string {
    if (!sheetModel?.coveredMap) return refA1;
    const parsed = CellReference.parse(refA1);
    if (!parsed.row || !parsed.col) return refA1;
    const anchorKey = sheetModel.coveredMap.get(`${parsed.row}-${parsed.col}`);
    if (!anchorKey) return refA1;
    const [row, col] = anchorKey.split('-').map(Number);
    return CellReference.toA1(row, col);
  }
}
