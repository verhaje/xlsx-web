// DependencyGraph.ts - Track cell dependencies for recalculation
//
// Maintains a directed graph: if cell A contains formula =B1+C1, then
// B1 and C1 are "precedents" of A, and A is a "dependent" of B1 and C1.
// When B1 changes, we look up its dependents to know what to recalc.

import { CellReference } from '../core/CellReference';
import { RangeUtils } from '../formula/RangeUtils';
import { Tokenizer } from '../formula/Tokenizer';
import type { Token } from '../types';

/**
 * Maximum number of individual cells a range will be expanded into for
 * dependency tracking. Ranges larger than this are skipped to avoid
 * excessive memory usage and O(n²) graph construction time. The formulas
 * still evaluate correctly; they just won't trigger fine-grained recalc.
 */
const MAX_RANGE_EXPANSION = 500;

/**
 * String interning pool for qualified keys. Identical key strings share
 * a single allocation, dramatically reducing memory when the same sheet
 * name appears in hundreds of thousands of keys.
 */
const _internPool: Map<string, string> = new Map();

function intern(s: string): string {
  const existing = _internPool.get(s);
  if (existing !== undefined) return existing;
  _internPool.set(s, s);
  return s;
}

/**
 * Qualified cell key: "SheetName::row-col"
 */
function qualifiedKey(sheetName: string, row: number, col: number): string {
  const key = `${sheetName}::${row}-${col}`;
  return intern(key);
}

/** Cached regex for parseRef-style matching inside extractReferences. */
const A1_RE = /^([A-Z]+)(\d+)$/i;

/**
 * DependencyGraph tracks which cells depend on which other cells.
 *
 * - `dependents`: cellKey → Set of cells that depend on it (need recalc when it changes)
 * - `precedents`: cellKey → Set of cells it references (what the formula reads)
 *
 * Both use qualified keys: "SheetName::row-col"
 */
export class DependencyGraph {
  /** Map from cell key to Set of cells that depend on it. */
  private dependents: Map<string, Set<string>> = new Map();
  /** Map from cell key to Set of cells it references. */
  private precedents: Map<string, Set<string>> = new Map();

  /**
   * Clear all dependencies and the interning pool.
   */
  clear(): void {
    this.dependents.clear();
    this.precedents.clear();
    _internPool.clear();
  }

  /**
   * Remove all dependencies for a specific cell (called before re-parsing its formula).
   */
  clearCell(sheetName: string, row: number, col: number): void {
    const cellKey = qualifiedKey(sheetName, row, col);

    // Remove this cell from all its precedents' dependent lists
    const precs = this.precedents.get(cellKey);
    if (precs) {
      for (const precKey of precs) {
        const deps = this.dependents.get(precKey);
        if (deps) {
          deps.delete(cellKey);
          if (deps.size === 0) this.dependents.delete(precKey);
        }
      }
      this.precedents.delete(cellKey);
    }

    // Also remove its dependent list (they'll re-register themselves)
    // We keep dependents intact since other cells still point to this
  }

  /**
   * Register that `cellKey` depends on `precedentKey`.
   * Uses Sets for O(1) duplicate checking (vs O(n) with arrays).
   */
  addDependency(cellKey: string, precedentKey: string): void {
    // cell depends on precedent
    let precs = this.precedents.get(cellKey);
    if (!precs) { precs = new Set(); this.precedents.set(cellKey, precs); }
    precs.add(precedentKey);

    // precedent has cell as dependent
    let deps = this.dependents.get(precedentKey);
    if (!deps) { deps = new Set(); this.dependents.set(precedentKey, deps); }
    deps.add(cellKey);
  }

  /**
   * Parse a formula and register all its cell/range references as dependencies.
   */
  setFormulaDependencies(sheetName: string, row: number, col: number, formulaText: string): void {
    const cellKey = qualifiedKey(sheetName, row, col);

    // Clear existing dependencies for this cell
    this.clearCell(sheetName, row, col);

    // Extract references from the formula and register them
    DependencyGraph._extractAndRegister(formulaText, sheetName, cellKey, this);
  }

  /**
   * Bulk-load formula dependencies for many cells at once during initial graph
   * construction. Skips `clearCell` per cell (assumes graph was freshly cleared)
   * and avoids intermediate array allocations from `extractReferences`.
   */
  bulkSetFormulaDependencies(
    cells: ReadonlyArray<{ sheetName: string; row: number; col: number; formula: string }>
  ): void {
    for (let i = 0; i < cells.length; i++) {
      const c = cells[i];
      const cellKey = qualifiedKey(c.sheetName, c.row, c.col);
      DependencyGraph._extractAndRegister(c.formula, c.sheetName, cellKey, this);
    }
  }

  /**
   * Get all cells that directly depend on the given cell.
   */
  getDependents(sheetName: string, row: number, col: number): string[] {
    const key = qualifiedKey(sheetName, row, col);
    const deps = this.dependents.get(key);
    return deps ? Array.from(deps) : [];
  }

  /**
   * Get all cells that need recalculation when the given cell changes.
   * Returns cells in correct evaluation order: if A depends on B, B comes first.
   * Detects circular references and excludes them.
   */
  getRecalcOrder(sheetName: string, row: number, col: number): string[] {
    const startKey = qualifiedKey(sheetName, row, col);

    // Step 1: collect all transitive dependents via BFS
    // Use index-based queue instead of shift() to avoid O(n²) array shifting
    const allDeps = new Set<string>();
    const queue: string[] = [startKey];
    let qHead = 0;
    while (qHead < queue.length) {
      const key = queue[qHead++];
      const deps = this.dependents.get(key);
      if (!deps) continue;
      for (const dep of deps) {
        if (dep === startKey) continue; // skip cycle back to start
        if (!allDeps.has(dep)) {
          allDeps.add(dep);
          queue.push(dep);
        }
      }
    }

    if (allDeps.size === 0) return [];

    // Step 2: topological sort the subgraph using Kahn's algorithm
    // Count in-degree for each node within the subgraph
    const inDegree = new Map<string, number>();
    for (const node of allDeps) inDegree.set(node, 0);

    for (const node of allDeps) {
      const precs = this.precedents.get(node);
      if (precs) {
        for (const prec of precs) {
          if (allDeps.has(prec)) {
            inDegree.set(node, (inDegree.get(node) || 0) + 1);
          }
        }
      }
    }

    // Enqueue nodes with in-degree 0 (within the subgraph)
    const sorted: string[] = [];
    const bfsQueue: string[] = [];
    let bfsHead = 0;
    for (const [node, deg] of inDegree) {
      if (deg === 0) bfsQueue.push(node);
    }

    while (bfsHead < bfsQueue.length) {
      const node = bfsQueue[bfsHead++];
      sorted.push(node);
      const deps = this.dependents.get(node);
      if (deps) {
        for (const dep of deps) {
          if (!allDeps.has(dep)) continue;
          const newDeg = (inDegree.get(dep) || 1) - 1;
          inDegree.set(dep, newDeg);
          if (newDeg === 0) bfsQueue.push(dep);
        }
      }
    }

    // If sorted.length < allDeps.size, there's a cycle — remaining nodes are skipped
    return sorted;
  }

  /**
   * Get the number of tracked cells.
   */
  get size(): number {
    return this.precedents.size;
  }

  /**
   * Parse a qualified key back to { sheetName, row, col }.
   */
  static parseQualifiedKey(key: string): { sheetName: string; row: number; col: number } {
    const sepIdx = key.indexOf('::');
    const sheetName = key.slice(0, sepIdx);
    const cellPart = key.slice(sepIdx + 2);
    const dashIdx = cellPart.indexOf('-');
    const r = parseInt(cellPart.slice(0, dashIdx), 10);
    const c = parseInt(cellPart.slice(dashIdx + 1), 10);
    return { sheetName, row: r, col: c };
  }

  /**
   * Build a qualified key.
   */
  static qualifiedKey(sheetName: string, row: number, col: number): string {
    return qualifiedKey(sheetName, row, col);
  }

  /**
   * Internal: extract references from formula text and directly register
   * them as dependencies in the graph. Avoids allocating an intermediate
   * array of ref strings.
   */
  private static _extractAndRegister(
    formulaText: string,
    defaultSheet: string,
    cellKey: string,
    graph: DependencyGraph,
  ): void {
    try {
      const tokens = Tokenizer.tokenize(formulaText);
      for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        const refSheet = token.sheet ? String(token.sheet).replace(/^'+|'+$/g, '') : defaultSheet;

        if (token.type === 'CELL') {
          const parsed = CellReference.parse(String(token.value));
          if (parsed.row && parsed.col) {
            graph.addDependency(cellKey, qualifiedKey(refSheet, parsed.row, parsed.col));
          }
        } else if (token.type === 'RANGE') {
          const rangeStr = String(token.value);
          // Skip whole-column ranges (e.g. "A:F") — expanding them would create
          // thousands of dependency entries. Formulas using column ranges will
          // still evaluate correctly; they just won't trigger fine-grained recalc.
          if (RangeUtils.isColumnRange(rangeStr)) {
            continue;
          }
          // Estimate size before expanding to avoid huge allocations
          const dims = RangeUtils.getRangeDimensions(rangeStr);
          if (dims.rows * dims.cols > MAX_RANGE_EXPANSION) {
            continue; // skip oversized ranges for dep tracking
          }
          const expanded = RangeUtils.expandRange(rangeStr);
          for (let j = 0; j < expanded.length; j++) {
            const parsed = CellReference.parse(expanded[j]);
            if (parsed.row && parsed.col) {
              graph.addDependency(cellKey, qualifiedKey(refSheet, parsed.row, parsed.col));
            }
          }
        }
      }
    } catch {
      // If tokenization fails, skip (formula may be invalid)
    }
  }

  /**
   * Extract all cell and range references from a formula text.
   * Returns qualified keys ("SheetName::row-col").
   */
  static extractReferences(formulaText: string, defaultSheet: string): string[] {
    const refs: string[] = [];
    try {
      const tokens = Tokenizer.tokenize(formulaText);
      for (const token of tokens) {
        const refSheet = token.sheet ? String(token.sheet).replace(/^'+|'+$/g, '') : defaultSheet;

        if (token.type === 'CELL') {
          const parsed = CellReference.parse(String(token.value));
          if (parsed.row && parsed.col) {
            refs.push(qualifiedKey(refSheet, parsed.row, parsed.col));
          }
        } else if (token.type === 'RANGE') {
          const rangeStr = String(token.value);
          if (RangeUtils.isColumnRange(rangeStr)) {
            continue;
          }
          // Cap expansion to MAX_RANGE_EXPANSION cells
          const dims = RangeUtils.getRangeDimensions(rangeStr);
          if (dims.rows * dims.cols > MAX_RANGE_EXPANSION) {
            continue;
          }
          const expanded = RangeUtils.expandRange(rangeStr);
          for (const a1 of expanded) {
            const parsed = CellReference.parse(a1);
            if (parsed.row && parsed.col) {
              refs.push(qualifiedKey(refSheet, parsed.row, parsed.col));
            }
          }
        }
      }
    } catch {
      // If tokenization fails, return empty (formula may be invalid)
    }
    return refs;
  }
}
