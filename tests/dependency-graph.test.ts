import { describe, it, expect, beforeEach } from 'vitest';
import { DependencyGraph } from '../src/ts/data/DependencyGraph';

describe('DependencyGraph', () => {
  let graph: DependencyGraph;

  beforeEach(() => {
    graph = new DependencyGraph();
  });

  describe('addDependency', () => {
    it('tracks a direct dependency', () => {
      graph.addDependency('Sheet1::1-1', 'Sheet1::2-1');
      const deps = graph.getDependents('Sheet1', 2, 1);
      expect(deps).toContain('Sheet1::1-1');
    });

    it('supports multiple dependents', () => {
      graph.addDependency('Sheet1::1-1', 'Sheet1::3-1'); // A1 depends on A3
      graph.addDependency('Sheet1::2-1', 'Sheet1::3-1'); // A2 depends on A3
      const deps = graph.getDependents('Sheet1', 3, 1);
      expect(deps).toHaveLength(2);
      expect(deps).toContain('Sheet1::1-1');
      expect(deps).toContain('Sheet1::2-1');
    });
  });

  describe('setFormulaDependencies', () => {
    it('parses and registers cell references from formula', () => {
      graph.setFormulaDependencies('Sheet1', 1, 1, 'B1+C1');
      const depsB = graph.getDependents('Sheet1', 1, 2); // B1
      const depsC = graph.getDependents('Sheet1', 1, 3); // C1
      expect(depsB).toContain('Sheet1::1-1');
      expect(depsC).toContain('Sheet1::1-1');
    });

    it('handles range references', () => {
      graph.setFormulaDependencies('Sheet1', 1, 1, 'SUM(A2:A4)');
      expect(graph.getDependents('Sheet1', 2, 1)).toContain('Sheet1::1-1');
      expect(graph.getDependents('Sheet1', 3, 1)).toContain('Sheet1::1-1');
      expect(graph.getDependents('Sheet1', 4, 1)).toContain('Sheet1::1-1');
    });

    it('clears old dependencies when re-parsing', () => {
      graph.setFormulaDependencies('Sheet1', 1, 1, 'B1');
      expect(graph.getDependents('Sheet1', 1, 2)).toContain('Sheet1::1-1');

      // Change formula
      graph.setFormulaDependencies('Sheet1', 1, 1, 'C1');
      expect(graph.getDependents('Sheet1', 1, 2)).not.toContain('Sheet1::1-1');
      expect(graph.getDependents('Sheet1', 1, 3)).toContain('Sheet1::1-1');
    });
  });

  describe('clearCell', () => {
    it('removes all dependencies of a cell', () => {
      graph.setFormulaDependencies('Sheet1', 1, 1, 'B1+C1');
      graph.clearCell('Sheet1', 1, 1);
      expect(graph.getDependents('Sheet1', 1, 2)).toHaveLength(0);
      expect(graph.getDependents('Sheet1', 1, 3)).toHaveLength(0);
    });
  });

  describe('getRecalcOrder', () => {
    it('returns dependents in topological order', () => {
      // A1 = B1 + C1 (A1 depends on B1 and C1)
      graph.setFormulaDependencies('Sheet1', 1, 1, 'B1+C1');
      // D1 = A1 * 2 (D1 depends on A1)
      graph.setFormulaDependencies('Sheet1', 1, 4, 'A1*2');

      // When B1 changes, recalc order should be A1 first, then D1
      const order = graph.getRecalcOrder('Sheet1', 1, 2);
      expect(order).toContain('Sheet1::1-1');
      expect(order).toContain('Sheet1::1-4');

      const a1Idx = order.indexOf('Sheet1::1-1');
      const d1Idx = order.indexOf('Sheet1::1-4');
      expect(a1Idx).toBeLessThan(d1Idx);
    });

    it('handles circular references gracefully', () => {
      // Create a cycle: A1 = B1, B1 = A1
      graph.addDependency('Sheet1::1-1', 'Sheet1::1-2');
      graph.addDependency('Sheet1::1-2', 'Sheet1::1-1');

      // Should not infinite loop
      const order = graph.getRecalcOrder('Sheet1', 1, 1);
      expect(order).toBeDefined();
    });

    it('does not include the changed cell itself', () => {
      graph.setFormulaDependencies('Sheet1', 1, 1, 'B1');
      const order = graph.getRecalcOrder('Sheet1', 1, 2);
      expect(order).not.toContain('Sheet1::1-2');
    });
  });

  describe('clear', () => {
    it('removes all tracked dependencies', () => {
      graph.setFormulaDependencies('Sheet1', 1, 1, 'B1+C1');
      graph.clear();
      expect(graph.getDependents('Sheet1', 1, 2)).toHaveLength(0);
      expect(graph.size).toBe(0);
    });
  });

  describe('extractReferences', () => {
    it('extracts cell references', () => {
      const refs = DependencyGraph.extractReferences('A1+B2*C3', 'Sheet1');
      expect(refs).toContain('Sheet1::1-1');
      expect(refs).toContain('Sheet1::2-2');
      expect(refs).toContain('Sheet1::3-3');
    });

    it('extracts range references', () => {
      const refs = DependencyGraph.extractReferences('SUM(A1:A3)', 'Sheet1');
      expect(refs).toContain('Sheet1::1-1');
      expect(refs).toContain('Sheet1::2-1');
      expect(refs).toContain('Sheet1::3-1');
    });

    it('handles cross-sheet references', () => {
      const refs = DependencyGraph.extractReferences('Sheet2!A1+B1', 'Sheet1');
      expect(refs).toContain('Sheet2::1-1');
      expect(refs).toContain('Sheet1::1-2');
    });

    it('returns empty array for invalid formulas', () => {
      const refs = DependencyGraph.extractReferences('', 'Sheet1');
      expect(refs).toEqual([]);
    });
  });

  describe('parseQualifiedKey', () => {
    it('parses a qualified key', () => {
      const result = DependencyGraph.parseQualifiedKey('Sheet1::3-5');
      expect(result).toEqual({ sheetName: 'Sheet1', row: 3, col: 5 });
    });

    it('handles sheet names with special chars', () => {
      const result = DependencyGraph.parseQualifiedKey('My Sheet::1-1');
      expect(result).toEqual({ sheetName: 'My Sheet', row: 1, col: 1 });
    });
  });
});
