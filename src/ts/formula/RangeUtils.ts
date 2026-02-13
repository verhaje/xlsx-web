// RangeUtils.ts - Helpers to parse and expand A1-style ranges (formula engine internal)

import type { RangeDimensions, ParsedRef } from '../types';

/**
 * Range parsing utilities used internally by the formula expression parser.
 * (Separate from CellReference to avoid circular deps between formula engine and core.)
 */
export class RangeUtils {
  /**
   * Parse a simple A1 reference into { col, row } (1-based).
   */
  static parseRef(ref: string): ParsedRef | null {
    if (!ref) return null;
    const m = /^([A-Z]+)(\d+)$/i.exec(ref);
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
   */
  static expandRange(range: string): string[] {
    if (!range) return [];
    const parts = range.split(':');
    if (parts.length === 1) return [parts[0]];
    const start = parts[0];
    const end = parts[1];
    const s = RangeUtils.parseRef(start);
    const e = RangeUtils.parseRef(end);
    if (!s || !e) return [start];
    const refs: string[] = [];
    for (let r = s.row; r <= e.row; r++) {
      for (let c = s.col; c <= e.col; c++) {
        let colStr = '';
        let cc = c;
        while (cc > 0) {
          const rem = (cc - 1) % 26;
          colStr = String.fromCharCode(65 + rem) + colStr;
          cc = Math.floor((cc - 1) / 26);
        }
        refs.push(`${colStr}${r}`);
      }
    }
    return refs;
  }

  /**
   * Get the row/column dimensions of a range string.
   */
  static getRangeDimensions(rangeStr: string): RangeDimensions {
    const parts = rangeStr.split(':');
    if (parts.length === 1) return { rows: 1, cols: 1 };
    const s = RangeUtils.parseRef(parts[0]);
    const e = RangeUtils.parseRef(parts[1]);
    if (!s || !e) return { rows: 1, cols: 1 };
    return {
      rows: e.row - s.row + 1,
      cols: e.col - s.col + 1,
    };
  }
}
