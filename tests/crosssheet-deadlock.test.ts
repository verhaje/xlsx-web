// @vitest-environment jsdom

/**
 * Regression test for the cross-sheet circular VLOOKUP deadlock.
 *
 * When two sheets mutually reference each other via whole-column VLOOKUP ranges
 * (e.g. IST→VLOOKUP(DPT!A:C) and DPT→VLOOKUP(IST!A:C)), the range resolution
 * of the second sheet discovers formula cells already mid-evaluation on the
 * first sheet.  Previously, the pendingCells check fired before the
 * evaluatingCells guard, returning an in-flight promise that transitively
 * awaited itself — creating a deadlock that hung the UI indefinitely.
 *
 * The fix moved the evaluatingCells check before pendingCells so that circular
 * dependencies return the XML-cached <v> value instead of the pending promise.
 */

import { describe, it, expect, vi as viMock } from 'vitest';
import { CellEvaluator } from '../src/ts/renderer/CellEvaluator';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';
import type { SheetInfo, SheetCacheEntry, SheetModel, IFormulaEngine } from '../src/ts/types';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Create a minimal SheetModel with no merges or shared formulas. */
function emptySheetModel(): SheetModel {
  return {
    mergedRanges: [],
    anchorMap: new Map(),
    coveredMap: new Map(),
    sharedFormulas: new Map(),
    sharedCells: new Map(),
  };
}

/**
 * Build a DOM Element representing an OOXML `<c>` cell node.
 *
 * @param ref  A1-style reference, e.g. "A1"
 * @param value  The cached `<v>` value (string or number)
 * @param formula  Optional formula text (creates a `<f>` child)
 * @param type  Optional cell type attribute (e.g. "s" for shared string)
 */
function makeCell(
  ref: string,
  value: string,
  formula?: string,
  type?: string,
): Element {
  const typeAttr = type ? ` t="${type}"` : '';
  const fNode = formula ? `<f>${formula}</f>` : '';
  const vNode = value !== '' ? `<v>${value}</v>` : '';
  const xml = `<c r="${ref}"${typeAttr}>${fNode}${vNode}</c>`;
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement;
}

/**
 * Build a sheetCache entry manually so CellEvaluator never touches the zip.
 * `buildCellMapForSheet` returns early when the sheet name already exists in
 * the cache.
 */
function buildCacheEntry(
  cells: { ref: string; row: number; col: number; value: string; formula?: string }[],
): SheetCacheEntry {
  const cellMap = new Map<string, Element>();
  let maxRow = 0;
  let maxCol = 0;
  for (const c of cells) {
    cellMap.set(`${c.row}-${c.col}`, makeCell(c.ref, c.value, c.formula));
    maxRow = Math.max(maxRow, c.row);
    maxCol = Math.max(maxCol, c.col);
  }
  return {
    cellMap,
    sheetDoc: null,
    maxRow,
    maxCol,
    sheetModel: emptySheetModel(),
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('Cross-sheet circular reference deadlock', () => {
  const SHEET_A = 'SheetA';
  const SHEET_B = 'SheetB';

  const sheetsInfo: SheetInfo[] = [
    { name: SHEET_A, relId: 'rId1', target: 'xl/worksheets/sheet1.xml' },
    { name: SHEET_B, relId: 'rId2', target: 'xl/worksheets/sheet2.xml' },
  ];

  /**
   * Shared setup:
   *
   *   SheetA!A1 = "keyA"               (plain value)
   *   SheetA!B1 = VLOOKUP(A1,SheetB!A:C,2,FALSE)  → depends on SheetB
   *   SheetA!C1 = 100                  (cached <v> value)
   *
   *   SheetB!A1 = "keyA"               (plain value)
   *   SheetB!B1 = VLOOKUP(A1,SheetA!A:C,3,FALSE)  → depends on SheetA
   *   SheetB!C1 = 200                  (cached <v> value)
   *
   * Mutual dependency: SheetA!B1 triggers resolveRange for SheetB!A:C,
   * which evaluates SheetB!B1 (formula), which triggers resolveRange for
   * SheetA!A:C, which finds SheetA!B1 already mid-evaluation.
   *
   * Without the fix, this creates a promise deadlock.
   */
  function buildSheetCacheAndEvaluator(): {
    evaluator: CellEvaluator;
    sheetCache: Map<string, any>;
  } {
    // Pre-build cell maps so zip is never accessed
    const sheetAEntry = buildCacheEntry([
      { ref: 'A1', row: 1, col: 1, value: 'keyA' },
      { ref: 'B1', row: 1, col: 2, value: '200', formula: "VLOOKUP(A1,SheetB!A:C,2,FALSE)" },
      { ref: 'C1', row: 1, col: 3, value: '100' },
    ]);

    const sheetBEntry = buildCacheEntry([
      { ref: 'A1', row: 1, col: 1, value: 'keyA' },
      { ref: 'B1', row: 1, col: 2, value: '100', formula: "VLOOKUP(A1,SheetA!A:C,3,FALSE)" },
      { ref: 'C1', row: 1, col: 3, value: '200' },
    ]);

    const sheetCache = new Map<string, any>();
    sheetCache.set(SHEET_A, sheetAEntry);
    sheetCache.set(SHEET_B, sheetBEntry);

    const formulaEngine = createFormulaEngine();
    // Use null as zip (never accessed because sheets are pre-cached)
    const evaluator = new CellEvaluator(
      null as any,  // zip
      sheetsInfo[0],
      sheetsInfo,
      [],           // sharedStrings
      formulaEngine,
      sheetCache,
    );

    return { evaluator, sheetCache };
  }

  it('resolves mutual cross-sheet VLOOKUP without deadlock', async () => {
    const { evaluator } = buildSheetCacheAndEvaluator();

    // This must complete within the test timeout (default 5s).
    // Before the fix, this would hang forever due to promise deadlock.
    const result = await evaluator.evaluateCellByRef('B1', SHEET_A);

    // SheetA!B1 = VLOOKUP("keyA", SheetB!A:C, 2, FALSE)
    // SheetB!A1="keyA", SheetB!B1 is a formula (circular → falls back to
    // cached <v> = "100"), SheetB!C1=200
    // VLOOKUP finds "keyA" in row 1, returns col 2 value.
    // The exact fallback value depends on how the circular ref resolves.
    // The critical assertion is that this doesn't deadlock.
    expect(result).toBeDefined();
  }, 5000);

  it('resolves both directions of the circular dependency', async () => {
    const { evaluator } = buildSheetCacheAndEvaluator();

    // Evaluate cells from both sheets concurrently
    const [resultA, resultB] = await Promise.all([
      evaluator.evaluateCellByRef('B1', SHEET_A),
      evaluator.evaluateCellByRef('B1', SHEET_B),
    ]);

    // Both must resolve without deadlock
    expect(resultA).toBeDefined();
    expect(resultB).toBeDefined();
  }, 5000);

  it('non-formula cells are unaffected by circular guard', async () => {
    const { evaluator } = buildSheetCacheAndEvaluator();

    // Plain value cells should still resolve normally
    const a1 = await evaluator.evaluateCellByRef('A1', SHEET_A);
    expect(a1).toBe('keyA');

    const c1 = await evaluator.evaluateCellByRef('C1', SHEET_B);
    expect(c1).toBe('200');
  });

  it('evaluating all formula cells in a sheet with mutual deps completes', async () => {
    const { evaluator } = buildSheetCacheAndEvaluator();

    // Simulate the renderer evaluating every formula cell in SheetA.
    // B1 triggers a resolveRange on SheetB, which discovers SheetB!B1's
    // formula referencing back into SheetA.  The evaluatingCells guard
    // prevents the deadlock.
    const [a1, b1, c1] = await Promise.all([
      evaluator.evaluateCellByRef('A1', SHEET_A),
      evaluator.evaluateCellByRef('B1', SHEET_A),
      evaluator.evaluateCellByRef('C1', SHEET_A),
    ]);

    expect(a1).toBe('keyA');
    expect(b1).toBeDefined();   // resolved via VLOOKUP (uses cached fallback for circular)
    expect(c1).toBe('100');
  }, 5000);

  it('resolveRange with mutual dependencies completes via fallback', async () => {
    const { evaluator } = buildSheetCacheAndEvaluator();

    // Resolve a full range on SheetA that includes a formula cell
    // referencing SheetB (which in turn references SheetA back).
    // The activeRangeSheets guard in resolveRange prevents the
    // range-level deadlock by falling back to cached <v> values.
    const range = await evaluator.resolveRange(
      SHEET_A,  // sheetName
      1, 1,     // startRow, startCol
      1, 3,     // endRow, endCol (A1:C1)
      SHEET_A,  // contextSheetName
    );

    expect(range).toHaveLength(3);
    expect(range[0]).toBe('keyA'); // A1 - plain value
    // B1 resolved through VLOOKUP (may use cached fallback for circular parts)
    expect(range[1]).toBeDefined();
    expect(range[2]).toBe('100');  // C1 - plain value
  }, 5000);

  it('second evaluation uses values cache (no re-deadlock)', async () => {
    const { evaluator } = buildSheetCacheAndEvaluator();

    // First pass populates the caches
    await evaluator.evaluateCellByRef('B1', SHEET_A);

    // Second pass should hit values cache and return immediately
    const result2 = await evaluator.evaluateCellByRef('B1', SHEET_A);
    expect(result2).toBeDefined();
  });

  it('three-sheet chain resolves without deadlock', async () => {
    // SheetA!B1 → VLOOKUP into SheetB → SheetB!B1 → VLOOKUP into SheetC → SheetC!B1 → VLOOKUP back into SheetA
    const SHEET_C = 'SheetC';
    const threeSheets: SheetInfo[] = [
      { name: SHEET_A, relId: 'rId1', target: 'xl/worksheets/sheet1.xml' },
      { name: SHEET_B, relId: 'rId2', target: 'xl/worksheets/sheet2.xml' },
      { name: SHEET_C, relId: 'rId3', target: 'xl/worksheets/sheet3.xml' },
    ];

    const sheetCache = new Map<string, any>();
    sheetCache.set(SHEET_A, buildCacheEntry([
      { ref: 'A1', row: 1, col: 1, value: 'k' },
      { ref: 'B1', row: 1, col: 2, value: '10', formula: "VLOOKUP(A1,SheetB!A:C,2,FALSE)" },
      { ref: 'C1', row: 1, col: 3, value: '30' },
    ]));
    sheetCache.set(SHEET_B, buildCacheEntry([
      { ref: 'A1', row: 1, col: 1, value: 'k' },
      { ref: 'B1', row: 1, col: 2, value: '20', formula: "VLOOKUP(A1,SheetC!A:C,2,FALSE)" },
      { ref: 'C1', row: 1, col: 3, value: '60' },
    ]));
    sheetCache.set(SHEET_C, buildCacheEntry([
      { ref: 'A1', row: 1, col: 1, value: 'k' },
      { ref: 'B1', row: 1, col: 2, value: '30', formula: "VLOOKUP(A1,SheetA!A:C,2,FALSE)" },
      { ref: 'C1', row: 1, col: 3, value: '90' },
    ]));

    const formulaEngine = createFormulaEngine();
    const evaluator = new CellEvaluator(
      null as any,
      threeSheets[0],
      threeSheets,
      [],
      formulaEngine,
      sheetCache,
    );

    // Must complete without deadlock
    const result = await evaluator.evaluateCellByRef('B1', SHEET_A);
    expect(result).toBeDefined();
  }, 5000);
});
