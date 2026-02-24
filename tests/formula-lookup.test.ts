import { describe, it, expect } from 'vitest';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';

const engine = createFormulaEngine({});

const cells: Record<string, any> = {
  // Vertical table A1:C5
  A1: 1, B1: 'Apple',  C1: 10,
  A2: 2, B2: 'Banana', C2: 20,
  A3: 3, B3: 'Cherry', C3: 30,
  A4: 4, B4: 'Date',   C4: 40,
  A5: 5, B5: 'Elderberry', C5: 50,
  // Horizontal table E1:I3
  E1: 1,       F1: 2,        G1: 3,        H1: 4,      I1: 5,
  E2: 'Apple', F2: 'Banana', G2: 'Cherry', H2: 'Date', I2: 'Elderberry',
  E3: 10,      F3: 20,       G3: 30,       H3: 40,     I3: 50,
  // 1D arrays
  J1: 10, J2: 20, J3: 30, J4: 40, J5: 50,
  K1: 'alpha', L1: 'beta', M1: 'gamma', N1: 'delta', O1: 'epsilon',
  P1: 30, P2: 10, P3: 50, P4: 20, P5: 40,
  // Lookup values
  Z1: 3, Z2: 'Cherry', Z3: 99,
};

const resolveCell = async (ref: string) => cells[ref] ?? '';

async function evalFormula(formula: string) {
  return engine.evaluateFormula(formula, { resolveCell });
}

describe('Lookup Functions', () => {
  describe('MATCH', () => {
    it.each([
      ['exact number',                '=MATCH(30,J1:J5,0)',            3],
      ['exact first',                 '=MATCH(10,J1:J5,0)',            1],
      ['exact last',                  '=MATCH(50,J1:J5,0)',            5],
      ['exact not found',             '=MATCH(99,J1:J5,0)',            '#N/A'],
      ['exact string',                '=MATCH("gamma",K1:O1,0)',       3],
      ['exact string case insensitive','=MATCH("BETA",K1:O1,0)',       2],
      ['exact string not found',      '=MATCH("omega",K1:O1,0)',       '#N/A'],
      ['approx ascending',            '=MATCH(25,J1:J5,1)',            2],
      ['approx exact value',          '=MATCH(30,J1:J5,1)',            3],
      ['approx below min',            '=MATCH(5,J1:J5,1)',             '#N/A'],
      ['approx above max',            '=MATCH(100,J1:J5,1)',           5],
      ['default match_type',          '=MATCH(35,J1:J5)',              3],
      ['approx descending',           '=MATCH(25,J1:J5,-1)',           3],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('INDEX', () => {
    it.each([
      ['2D row 1 col 1',   '=INDEX(A1:C5,1,1)',  1],
      ['2D row 1 col 2',   '=INDEX(A1:C5,1,2)',  'Apple'],
      ['2D row 3 col 3',   '=INDEX(A1:C5,3,3)',  30],
      ['2D row 5 col 2',   '=INDEX(A1:C5,5,2)',  'Elderberry'],
      ['2D last cell',     '=INDEX(A1:C5,5,3)',   50],
      ['1D row 3',         '=INDEX(J1:J5,3,1)',   30],
      ['1D row 1',         '=INDEX(J1:J5,1,1)',   10],
      ['row out of bounds','=INDEX(A1:C5,10,1)',   '#REF!'],
      ['col out of bounds','=INDEX(A1:C5,1,10)',   '#REF!'],
      ['negative row',     '=INDEX(A1:C5,-1,1)',   '#VALUE!'],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('VLOOKUP', () => {
    it.each([
      ['exact number col 2',    '=VLOOKUP(2,A1:C5,2,FALSE)',       'Banana'],
      ['exact number col 3',    '=VLOOKUP(3,A1:C5,3,FALSE)',       30],
      ['exact first row',       '=VLOOKUP(1,A1:C5,2,FALSE)',       'Apple'],
      ['exact last row',        '=VLOOKUP(5,A1:C5,2,FALSE)',       'Elderberry'],
      ['exact not found',       '=VLOOKUP(99,A1:C5,2,FALSE)',      '#N/A'],
      ['exact using 0',         '=VLOOKUP(4,A1:C5,3,0)',           40],
      ['approx default',        '=VLOOKUP(2.5,A1:C5,2)',           'Banana'],
      ['approx TRUE',           '=VLOOKUP(3.9,A1:C5,2,TRUE)',      'Cherry'],
      ['approx exact value',    '=VLOOKUP(4,A1:C5,3,TRUE)',        40],
      ['approx above max',      '=VLOOKUP(100,A1:C5,2,TRUE)',      'Elderberry'],
      ['approx below min',      '=VLOOKUP(0,A1:C5,2,TRUE)',        '#N/A'],
      ['col index 0',           '=VLOOKUP(2,A1:C5,0,FALSE)',       '#VALUE!'],
      ['col index too large',   '=VLOOKUP(2,A1:C5,10,FALSE)',      '#REF!'],
      ['with cell ref',         '=VLOOKUP(Z1,A1:C5,2,FALSE)',      'Cherry'],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('HLOOKUP', () => {
    it.each([
      ['exact number row 2',    '=HLOOKUP(2,E1:I3,2,FALSE)',       'Banana'],
      ['exact number row 3',    '=HLOOKUP(3,E1:I3,3,FALSE)',       30],
      ['exact first col',       '=HLOOKUP(1,E1:I3,2,FALSE)',       'Apple'],
      ['exact last col',        '=HLOOKUP(5,E1:I3,2,FALSE)',       'Elderberry'],
      ['exact not found',       '=HLOOKUP(99,E1:I3,2,FALSE)',      '#N/A'],
      ['approx default',        '=HLOOKUP(2.5,E1:I3,2)',           'Banana'],
      ['approx above max',      '=HLOOKUP(100,E1:I3,3,TRUE)',      50],
      ['row index 0',           '=HLOOKUP(2,E1:I3,0,FALSE)',       '#VALUE!'],
      ['row index too large',   '=HLOOKUP(2,E1:I3,10,FALSE)',      '#REF!'],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('Combined / Nested', () => {
    it.each([
      ['INDEX-MATCH combo',         '=INDEX(C1:C5,MATCH(3,A1:A5,0),1)',                     30],
      ['INDEX-MATCH horizontal',    '=INDEX(E3:I3,1,MATCH(3,E1:I1,0))',                     30],
      ['IF with VLOOKUP',           '=IF(VLOOKUP(2,A1:C5,3,FALSE)>15,"big","small")',       'big'],
      ['IFERROR with VLOOKUP miss', '=IFERROR(VLOOKUP(99,A1:C5,2,FALSE),"Not found")',      'Not found'],
      ['IFERROR with VLOOKUP hit',  '=IFERROR(VLOOKUP(2,A1:C5,2,FALSE),"Not found")',       'Banana'],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('Edge cases', () => {
    it('VLOOKUP with known result', async () => {
      expect(await evalFormula('=VLOOKUP(1,A1:C5,2,FALSE)')).toBe('Apple');
    });
    it('INDEX single cell', async () => {
      expect(await evalFormula('=INDEX(A1:A1,1,1)')).toBe(1);
    });
    it('MATCH in single cell', async () => {
      expect(await evalFormula('=MATCH(1,A1:A1,0)')).toBe(1);
    });
  });
});
