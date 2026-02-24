// CellReference.ts - Cell reference parsing and conversion utilities

import type { CellRef, RangeRef } from '../types';

/** Cached regex for A1-style cell reference parsing */
const A1_RE = /([A-Z]+)(\d+)/i;

/** LRU parse cache to avoid re-parsing the same A1 refs (hot path: 3-4× per cell) */
const PARSE_CACHE = new Map<string, CellRef>();
const PARSE_CACHE_MAX = 65536;

/** Sentinel for refs that fail to parse */
const ZERO_REF: CellRef = Object.freeze({ col: 0, row: 0 });

/**
 * Utility class for working with A1-style cell references.
 */
export class CellReference {
  /**
   * Parse an A1-style cell reference (e.g. "B3") into { row, col }.
   * Results are cached to avoid repeated regex + allocation overhead.
   */
  static parse(ref: string): CellRef {
    const cached = PARSE_CACHE.get(ref);
    if (cached !== undefined) return cached;

    const match = A1_RE.exec(ref);
    if (!match) {
      PARSE_CACHE.set(ref, ZERO_REF);
      return ZERO_REF;
    }
    const colLetters = match[1].toUpperCase();
    const row = parseInt(match[2], 10);
    const result: CellRef = { col: CellReference.columnNameToIndex(colLetters), row };

    // Evict oldest entries when cache is full
    if (PARSE_CACHE.size >= PARSE_CACHE_MAX) {
      const first = PARSE_CACHE.keys().next().value;
      if (first !== undefined) PARSE_CACHE.delete(first);
    }
    PARSE_CACHE.set(ref, result);
    return result;
  }

  /**
   * Clear the parse cache (useful for testing).
   */
  static clearParseCache(): void {
    PARSE_CACHE.clear();
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

  /** Cache for column index → name (covers columns 1-702, i.e. A..ZZ) */
  private static colNameCache: string[] = [];

  /**
   * Convert a 1-based column index to a column letter string.
   * Cached for columns 1-702.
   */
  static columnIndexToName(index: number): string {
    if (index > 0 && index <= 702) {
      const cached = CellReference.colNameCache[index];
      if (cached) return cached;
    }
    let result = '';
    let current = index;
    while (current > 0) {
      const remainder = (current - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      current = Math.floor((current - 1) / 26);
    }
    if (index > 0 && index <= 702) {
      CellReference.colNameCache[index] = result;
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
   * Overload with pre-parsed row/col avoids redundant CellReference.parse() calls.
   */
  static mapMergedRef(refA1: string, sheetModel: { coveredMap?: Map<string, string> } | null): string;
  static mapMergedRef(refA1: string, sheetModel: { coveredMap?: Map<string, string> } | null, row: number, col: number): string;
  static mapMergedRef(refA1: string, sheetModel: { coveredMap?: Map<string, string> } | null, row?: number, col?: number): string {
    if (!sheetModel?.coveredMap || sheetModel.coveredMap.size === 0) return refA1;
    // Use pre-parsed row/col if provided, otherwise parse
    if (row === undefined || col === undefined) {
      const parsed = CellReference.parse(refA1);
      row = parsed.row;
      col = parsed.col;
    }
    if (!row || !col) return refA1;
    const anchorKey = sheetModel.coveredMap.get(`${row}-${col}`);
    if (!anchorKey) return refA1;
    const dashIdx = anchorKey.indexOf('-');
    const aRow = parseInt(anchorKey.slice(0, dashIdx), 10);
    const aCol = parseInt(anchorKey.slice(dashIdx + 1), 10);
    return CellReference.toA1(aRow, aCol);
  }
}
