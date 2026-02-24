import { describe, it, expect } from 'vitest';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';

const engine = createFormulaEngine({});

async function evalFormula(formula: string) {
  return engine.evaluateFormula(formula, { resolveCell: async () => 0 });
}

describe('Date/Time Functions', () => {
  describe('DATE serial values', () => {
    it.each([
      ['1900-01-01',            '=DATE(1900,1,1)',   1],
      ['1900-02-28',            '=DATE(1900,2,28)',  59],
      ['1900-02-29 (Excel bug)','=DATE(1900,2,29)',  60],
      ['1900-03-01',            '=DATE(1900,3,1)',   61],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toBe(expected);
    });
  });

  describe('Extractors from serial', () => {
    it.each([
      ['YEAR(61)',              '=YEAR(61)',   1900],
      ['MONTH(61)',             '=MONTH(61)',  3],
      ['DAY(61)',               '=DAY(61)',    1],
      ['YEAR(60) → 1900-02-29','=YEAR(60)',   1900],
      ['MONTH(60)',             '=MONTH(60)',  2],
      ['DAY(60)',               '=DAY(60)',    29],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toBe(expected);
    });
  });

  describe('Nested DATE extraction', () => {
    it.each([
      ['YEAR(DATE(2024,1,15))', '=YEAR(DATE(2024,1,15))',  2024],
      ['MONTH(DATE(2024,12,5))','=MONTH(DATE(2024,12,5))', 12],
      ['DAY(DATE(2024,12,5))',  '=DAY(DATE(2024,12,5))',   5],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toBe(expected);
    });
  });

  describe('TODAY()', () => {
    it('YEAR(TODAY()) returns current year', async () => {
      const now = new Date();
      expect(await evalFormula('=YEAR(TODAY())')).toBe(now.getFullYear());
    });
    it('MONTH(TODAY()) returns current month', async () => {
      const now = new Date();
      expect(await evalFormula('=MONTH(TODAY())')).toBe(now.getMonth() + 1);
    });
    it('DAY(TODAY()) returns current day', async () => {
      const now = new Date();
      expect(await evalFormula('=DAY(TODAY())')).toBe(now.getDate());
    });
  });
});
