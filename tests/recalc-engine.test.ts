import { describe, it, expect, beforeEach } from 'vitest';
import { CellStore } from '../src/ts/data/CellStore';
import { DependencyGraph } from '../src/ts/data/DependencyGraph';
import { RecalcEngine } from '../src/ts/data/RecalcEngine';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';
import type { IFormulaEngine } from '../src/ts/types';

describe('RecalcEngine', () => {
  let store: CellStore;
  let graph: DependencyGraph;
  let engine: IFormulaEngine;
  let recalc: RecalcEngine;
  let cellStores: Map<string, CellStore>;
  let updatedCells: Array<{ sheet: string; row: number; col: number; value: any }>;

  beforeEach(() => {
    store = new CellStore('Sheet1');
    graph = new DependencyGraph();
    engine = createFormulaEngine({});
    cellStores = new Map([['Sheet1', store]]);
    updatedCells = [];

    recalc = new RecalcEngine({
      cellStores,
      graph,
      formulaEngine: engine,
      sharedStrings: [],
      onCellUpdated: (sheet, row, col, value) => {
        updatedCells.push({ sheet, row, col, value });
      },
    });
  });

  describe('setCellValue', () => {
    it('sets a plain value', async () => {
      await recalc.setCellValue('Sheet1', 1, 1, 42);
      expect(store.getValue(1, 1)).toBe(42);
    });

    it('recalculates dependent formula cells', async () => {
      // Set up: A1=10, A2=20, A3=SUM(A1,A2) which is =A1+A2
      store.load(1, 1, 10, null);
      store.load(2, 1, 20, null);
      store.load(3, 1, 30, 'A1+A2');
      graph.setFormulaDependencies('Sheet1', 3, 1, 'A1+A2');

      // Change A1 to 50
      await recalc.setCellValue('Sheet1', 1, 1, 50);

      // A3 should be recalculated to 50 + 20 = 70
      expect(store.getValue(3, 1)).toBe(70);
      expect(updatedCells.some(c => c.row === 3 && c.col === 1 && c.value === 70)).toBe(true);
    });

    it('cascades through multiple levels of dependents', async () => {
      // A1 = value, A2 = A1*2, A3 = A2+10
      store.load(1, 1, 5, null);
      store.load(2, 1, 10, 'A1*2');
      store.load(3, 1, 20, 'A2+10');
      graph.setFormulaDependencies('Sheet1', 2, 1, 'A1*2');
      graph.setFormulaDependencies('Sheet1', 3, 1, 'A2+10');

      // Change A1 to 100
      await recalc.setCellValue('Sheet1', 1, 1, 100);

      // A2 = 100*2 = 200
      expect(store.getValue(2, 1)).toBe(200);
      // A3 = 200+10 = 210
      expect(store.getValue(3, 1)).toBe(210);
    });
  });

  describe('setCellFormula', () => {
    it('sets and evaluates a formula', async () => {
      store.load(1, 1, 10, null);
      store.load(2, 1, 20, null);

      const result = await recalc.setCellFormula('Sheet1', 3, 1, '=A1+A2');
      expect(result).toBe(30);
      expect(store.getValue(3, 1)).toBe(30);
      expect(store.getFormula(3, 1)).toBe('A1+A2');
    });

    it('updates dependency graph for the new formula', async () => {
      store.load(1, 1, 10, null);
      store.load(2, 1, 20, null);

      await recalc.setCellFormula('Sheet1', 3, 1, '=A1+A2');

      // Changing A1 should now recalculate A3
      await recalc.setCellValue('Sheet1', 1, 1, 100);
      expect(store.getValue(3, 1)).toBe(120);
    });

    it('handles formula change on existing formula cell', async () => {
      store.load(1, 1, 10, null);
      store.load(2, 1, 20, null);
      store.load(1, 2, 5, null); // B1

      // Set A3 = A1 + A2
      await recalc.setCellFormula('Sheet1', 3, 1, '=A1+A2');
      expect(store.getValue(3, 1)).toBe(30);

      // Change A3 formula to =B1*3
      await recalc.setCellFormula('Sheet1', 3, 1, '=B1*3');
      expect(store.getValue(3, 1)).toBe(15);

      // Changing A1 should NOT affect A3 anymore
      updatedCells = [];
      await recalc.setCellValue('Sheet1', 1, 1, 999);
      const a3Update = updatedCells.find(c => c.row === 3 && c.col === 1);
      expect(a3Update).toBeUndefined();
    });
  });

  describe('buildFullGraph', () => {
    it('builds graph for all formula cells across sheets', () => {
      const store2 = new CellStore('Sheet2');
      cellStores.set('Sheet2', store2);

      store.load(1, 1, 10, null);
      store.load(2, 1, '', 'A1*2');
      store2.load(1, 1, '', 'Sheet1!A1+5');

      recalc.buildFullGraph();

      // A1 change should affect Sheet1!A2 and Sheet2!A1
      const deps = graph.getDependents('Sheet1', 1, 1);
      expect(deps.length).toBeGreaterThanOrEqual(1);
    });
  });

  describe('cross-sheet recalculation', () => {
    it('recalculates formulas referencing other sheets', async () => {
      const store2 = new CellStore('Sheet2');
      cellStores.set('Sheet2', store2);

      store.load(1, 1, 10, null); // Sheet1!A1 = 10
      store2.load(1, 1, 10, 'Sheet1!A1');  // Sheet2!A1 = Sheet1!A1
      graph.setFormulaDependencies('Sheet2', 1, 1, 'Sheet1!A1');

      // Re-init recalc with both stores
      recalc = new RecalcEngine({
        cellStores,
        graph,
        formulaEngine: engine,
        sharedStrings: [],
        onCellUpdated: (sheet, row, col, value) => {
          updatedCells.push({ sheet, row, col, value });
        },
      });

      // Change Sheet1!A1
      await recalc.setCellValue('Sheet1', 1, 1, 99);

      // Sheet2!A1 should be recalculated
      expect(store2.getValue(1, 1)).toBe(99);
    });
  });
});
