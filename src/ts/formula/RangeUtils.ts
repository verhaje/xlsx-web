// RangeUtils.ts - Helpers to parse and expand A1-style ranges (formula engine internal)

import type { RangeDimensions, ParsedRef } from '../types';

/**
 * Range parsing utilities used internally by the formula expression parser.
 * (Separate from CellReference to avoid circular deps between formula engine and core.)
 */
export class RangeUtils {
  /** Column-only regex: matches e.g. "A", "XFD" (no row digits). */
  private static readonly COL_ONLY_RE = /^[A-Z]+$/i;

  /** Cached regex for parseRef. */
  private static readonly REF_RE = /^([A-Z]+)(\d+)$/i;

  /** Maximum rows to expand for whole-column ranges to prevent OOM. */
  private static readonly MAX_COLUMN_RANGE_ROWS = 5000;

  /**
   * Convert a column letter string to a 1-based column number.
   */
  static colLetterToNum(col: string): number {
    const upper = col.toUpperCase();
    let num = 0;
    for (let i = 0; i < upper.length; i++) {
      num = num * 26 + (upper.charCodeAt(i) - 64);
    }
    return num;
  }

  /** Cache for colNumToLetter — avoids repeated string construction in hot loops */
  private static colLetterCache: string[] = [];

  /**
   * Convert a 1-based column number to a column letter string.
   * Results are cached for columns 1-702 (A..ZZ) which covers all practical sheets.
   */
  static colNumToLetter(col: number): string {
    if (col > 0 && col <= 702) {
      const cached = RangeUtils.colLetterCache[col];
      if (cached) return cached;
    }
    let s = '';
    let c = col;
    while (c > 0) {
      const rem = (c - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      c = Math.floor((c - 1) / 26);
    }
    if (col > 0 && col <= 702) {
      RangeUtils.colLetterCache[col] = s;
    }
    return s;
  }

  /**
   * Test whether a range string is a whole-column range (e.g. "A:F").
   */
  static isColumnRange(range: string): boolean {
    const parts = range.split(':');
    if (parts.length !== 2) return false;
    return RangeUtils.COL_ONLY_RE.test(parts[0]) && RangeUtils.COL_ONLY_RE.test(parts[1]);
  }

  /**
   * Parse a simple A1 reference into { col, row } (1-based).
   */
  static parseRef(ref: string): ParsedRef | null {
    if (!ref) return null;
    const m = RangeUtils.REF_RE.exec(ref);
    if (!m) return null;
    const col = m[1].toUpperCase();
    let colNum = 0;
    for (let i = 0; i < col.length; i++) {
      colNum = colNum * 26 + (col.charCodeAt(i) - 64);
    }
    const row = parseInt(m[2], 10);
    return { col: colNum, row };
  }

  /**
   * Expand a range string (e.g. "A1:C3") into an array of A1 reference strings.
   * For whole-column ranges like "A:F", uses `maxRow` (defaults to 1000 if unspecified).
   */
  static expandRange(range: string, maxRow?: number): string[] {
    if (!range) return [];
    const parts = range.split(':');
    if (parts.length === 1) return [parts[0]];
    const startPart = parts[0];
    const endPart = parts[1];

    // Whole-column range (e.g. "A:F")
    if (RangeUtils.COL_ONLY_RE.test(startPart) && RangeUtils.COL_ONLY_RE.test(endPart)) {
      const sc = RangeUtils.colLetterToNum(startPart);
      const ec = RangeUtils.colLetterToNum(endPart);
      const mr = Math.min(
        maxRow && maxRow > 0 ? maxRow : RangeUtils.MAX_COLUMN_RANGE_ROWS,
        RangeUtils.MAX_COLUMN_RANGE_ROWS
      );
      const refs: string[] = [];
      for (let r = 1; r <= mr; r++) {
        for (let c = sc; c <= ec; c++) {
          refs.push(`${RangeUtils.colNumToLetter(c)}${r}`);
        }
      }
      return refs;
    }

    const s = RangeUtils.parseRef(startPart);
    const e = RangeUtils.parseRef(endPart);
    if (!s || !e) return [startPart];
    const refs: string[] = [];
    for (let r = s.row; r <= e.row; r++) {
      for (let c = s.col; c <= e.col; c++) {
        refs.push(`${RangeUtils.colNumToLetter(c)}${r}`);
      }
    }
    return refs;
  }

  /**
   * Get the row/column dimensions of a range string.
   * For whole-column ranges like "A:F", uses `maxRow` (defaults to 1000 if unspecified).
   */
  static getRangeDimensions(rangeStr: string, maxRow?: number): RangeDimensions {
    const parts = rangeStr.split(':');
    if (parts.length === 1) return { rows: 1, cols: 1 };

    // Whole-column range
    if (RangeUtils.COL_ONLY_RE.test(parts[0]) && RangeUtils.COL_ONLY_RE.test(parts[1])) {
      const sc = RangeUtils.colLetterToNum(parts[0]);
      const ec = RangeUtils.colLetterToNum(parts[1]);
      const mr = Math.min(
        maxRow && maxRow > 0 ? maxRow : RangeUtils.MAX_COLUMN_RANGE_ROWS,
        RangeUtils.MAX_COLUMN_RANGE_ROWS
      );
      return { rows: mr, cols: ec - sc + 1 };
    }

    const s = RangeUtils.parseRef(parts[0]);
    const e = RangeUtils.parseRef(parts[1]);
    if (!s || !e) return { rows: 1, cols: 1 };
    return {
      rows: e.row - s.row + 1,
      cols: e.col - s.col + 1,
    };
  }
}
