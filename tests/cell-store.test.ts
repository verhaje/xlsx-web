import { describe, it, expect, beforeEach } from 'vitest';
import { CellStore } from '../src/ts/data/CellStore';
import type { CellData } from '../src/ts/data/CellStore';

describe('CellStore', () => {
  let store: CellStore;

  beforeEach(() => {
    store = new CellStore('Sheet1');
  });

  describe('constructor', () => {
    it('creates a store with the given sheet name', () => {
      expect(store.sheetName).toBe('Sheet1');
      expect(store.size).toBe(0);
      expect(store.maxRow).toBe(0);
      expect(store.maxCol).toBe(0);
    });
  });

  describe('set and get', () => {
    it('sets and gets a numeric value', () => {
      store.set(1, 1, 42);
      const cell = store.get(1, 1);
      expect(cell).toBeDefined();
      expect(cell!.value).toBe(42);
      expect(cell!.type).toBe('number');
      expect(cell!.formula).toBeNull();
      expect(store.isDirty(1, 1)).toBe(true);
    });

    it('sets and gets a string value', () => {
      store.set(1, 2, 'Hello');
      const cell = store.get(1, 2);
      expect(cell!.value).toBe('Hello');
      expect(cell!.type).toBe('string');
    });

    it('sets and gets a boolean value', () => {
      store.set(2, 1, true);
      const cell = store.get(2, 1);
      expect(cell!.value).toBe(true);
      expect(cell!.type).toBe('boolean');
    });

    it('detects formulas starting with =', () => {
      store.set(1, 1, '=SUM(A2:A3)');
      const cell = store.get(1, 1);
      expect(cell!.type).toBe('formula');
      expect(cell!.formula).toBe('SUM(A2:A3)');
      expect(cell!.value).toBe(''); // not yet evaluated
    });

    it('coerces numeric strings to numbers', () => {
      store.set(1, 1, '123.45');
      const cell = store.get(1, 1);
      expect(cell!.value).toBe(123.45);
      expect(cell!.type).toBe('number');
    });

    it('returns previous value on overwrite', () => {
      store.set(1, 1, 'old');
      const prev = store.set(1, 1, 'new');
      expect(prev).toBeDefined();
      expect(prev!.value).toBe('old');
    });

    it('tracks max row and column', () => {
      store.set(5, 10, 'x');
      expect(store.maxRow).toBe(5);
      expect(store.maxCol).toBe(10);
      store.set(3, 15, 'y');
      expect(store.maxRow).toBe(5);
      expect(store.maxCol).toBe(15);
    });
  });

  describe('getByKey', () => {
    it('gets cell by key string', () => {
      store.set(2, 3, 'value');
      const cell = store.getByKey('2-3');
      expect(cell!.value).toBe('value');
    });

    it('returns undefined for missing key', () => {
      expect(store.getByKey('999-999')).toBeUndefined();
    });
  });

  describe('getByRef', () => {
    it('gets cell by A1-style reference', () => {
      store.set(3, 2, 'test'); // B3
      const cell = store.getByRef('B3');
      expect(cell).toBeDefined();
      expect(cell!.value).toBe('test');
    });

    it('returns undefined for invalid ref', () => {
      expect(store.getByRef('')).toBeUndefined();
    });
  });

  describe('getValue', () => {
    it('returns the cell value', () => {
      store.set(1, 1, 42);
      expect(store.getValue(1, 1)).toBe(42);
    });

    it('returns empty string for missing cells', () => {
      expect(store.getValue(99, 99)).toBe('');
    });
  });

  describe('getFormula', () => {
    it('returns the formula for formula cells', () => {
      store.set(1, 1, '=A2+A3');
      expect(store.getFormula(1, 1)).toBe('A2+A3');
    });

    it('returns null for non-formula cells', () => {
      store.set(1, 1, 42);
      expect(store.getFormula(1, 1)).toBeNull();
    });
  });

  describe('setFormula', () => {
    it('sets a formula cell explicitly', () => {
      store.setFormula(1, 1, 'SUM(A2:A3)', 60);
      const cell = store.get(1, 1);
      expect(cell!.formula).toBe('SUM(A2:A3)');
      expect(cell!.value).toBe(60);
      expect(cell!.type).toBe('formula');
      expect(store.isDirty(1, 1)).toBe(true);
    });
  });

  describe('setComputedValue', () => {
    it('updates the value without marking dirty', () => {
      store.load(1, 1, '', 'SUM(A2:A3)');
      expect(store.isDirty(1, 1)).toBe(false);
      store.setComputedValue(1, 1, 60);
      expect(store.getValue(1, 1)).toBe(60);
      expect(store.isDirty(1, 1)).toBe(false);
    });
  });

  describe('load', () => {
    it('loads cell data without marking dirty', () => {
      store.load(1, 1, 42, null, { styleIndex: 5 });
      const cell = store.get(1, 1);
      expect(cell!.value).toBe(42);
      expect(store.isDirty(1, 1)).toBe(false);
      expect(cell!.styleIndex).toBe(5);
    });

    it('loads formula cells', () => {
      store.load(1, 1, 100, 'SUM(A2:A3)');
      const cell = store.get(1, 1);
      expect(cell!.type).toBe('formula');
      expect(cell!.formula).toBe('SUM(A2:A3)');
      expect(cell!.value).toBe(100);
    });
  });

  describe('delete', () => {
    it('deletes a cell and returns previous data', () => {
      store.set(1, 1, 'old');
      const prev = store.delete(1, 1);
      expect(prev!.value).toBe('old');
      expect(store.get(1, 1)).toBeUndefined();
    });

    it('returns undefined for missing cell', () => {
      expect(store.delete(99, 99)).toBeUndefined();
    });
  });

  describe('has', () => {
    it('returns true for existing cells', () => {
      store.set(1, 1, 'x');
      expect(store.has(1, 1)).toBe(true);
    });

    it('returns false for missing cells', () => {
      expect(store.has(99, 99)).toBe(false);
    });
  });

  describe('forEach', () => {
    it('iterates all cells', () => {
      store.set(1, 1, 'A');
      store.set(2, 3, 'B');
      const visited: Array<{ key: string; row: number; col: number }> = [];
      store.forEach((key, _data, row, col) => {
        visited.push({ key, row, col });
      });
      expect(visited).toHaveLength(2);
      expect(visited.map(v => v.key).sort()).toEqual(['1-1', '2-3']);
    });
  });

  describe('getFormulaCells', () => {
    it('returns only formula cells', () => {
      store.set(1, 1, '=A2+A3');
      store.set(2, 1, 42);
      store.set(3, 1, '=B1*2');
      const formulas = store.getFormulaCells();
      expect(formulas).toHaveLength(2);
      expect(formulas.map(f => f.formula).sort()).toEqual(['A2+A3', 'B1*2']);
    });
  });

  describe('getDirtyCells', () => {
    it('returns only dirty cells', () => {
      store.load(1, 1, 'loaded', null); // not dirty
      store.set(2, 1, 'edited');        // dirty
      const dirty = store.getDirtyCells();
      expect(dirty).toHaveLength(1);
      expect(dirty[0].data.value).toBe('edited');
    });
  });

  describe('clearDirtyFlags', () => {
    it('clears all dirty flags', () => {
      store.set(1, 1, 'edited');
      store.set(2, 1, 'also edited');
      store.clearDirtyFlags();
      expect(store.getDirtyCells()).toHaveLength(0);
    });
  });

  describe('inferType', () => {
    it.each([
      [null, 'empty'],
      [undefined, 'empty'],
      ['', 'empty'],
      [42, 'number'],
      ['123', 'number'],
      [true, 'boolean'],
      ['Hello', 'string'],
      ['#REF!', 'error'],
      ['#DIV/0!', 'error'],
    ])('infers type for %s as %s', (value, expected) => {
      expect(CellStore.inferType(value)).toBe(expected);
    });
  });
});
