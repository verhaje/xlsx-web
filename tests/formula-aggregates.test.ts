import { describe, it, expect } from 'vitest';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';

const engine = createFormulaEngine({});

const cells: Record<string, any> = {
  // Headers
  A1: 'Product', B1: 'Region', C1: 'Sales', D1: 'Quantity',
  // Data rows
  A2: 'Apple',  B2: 'North', C2: 100, D2: 10,
  A3: 'Banana', B3: 'South', C3: 200, D3: 20,
  A4: 'Apple',  B4: 'South', C4: 150, D4: 15,
  A5: 'Cherry', B5: 'North', C5: 300, D5: 30,
  A6: 'Banana', B6: 'North', C6: 250, D6: 25,
  A7: 'Apple',  B7: 'North', C7: 175, D7: 17,
  // Numeric test data
  E1: 10, E2: 20, E3: 30, E4: 40, E5: 50,
  F1: 5,  F2: 15, F3: 25, F4: 35, F5: 45,
  // Mixed data
  G1: '', G2: 'text', G3: 100, G4: '', G5: 200,
  H1: 1,  H2: 2,  H3: 3,   H4: 4,  H5: 5,
};

const resolveCell = async (ref: string) => cells[ref] ?? '';

async function evalFormula(formula: string) {
  return engine.evaluateFormula(formula, { resolveCell });
}

/** Helper for floating-point assertions. */
function expectApprox(actual: any, expected: number) {
  expect(typeof actual).toBe('number');
  expect(actual).toBeCloseTo(expected, 3);
}

describe('Conditional Aggregate Functions', () => {
  describe('SUMIF', () => {
    it.each([
      ['text match Apple',   '=SUMIF(A2:A7,"Apple",C2:C7)',     425],
      ['text match Banana',  '=SUMIF(A2:A7,"Banana",C2:C7)',    450],
      ['text match Cherry',  '=SUMIF(A2:A7,"Cherry",C2:C7)',    300],
      ['numeric equal',      '=SUMIF(C2:C7,100,D2:D7)',         10],
      ['numeric greater',    '=SUMIF(C2:C7,">200",D2:D7)',      55],
      ['numeric less',       '=SUMIF(C2:C7,"<150",D2:D7)',      10],
      ['numeric gte',        '=SUMIF(C2:C7,">=200",D2:D7)',     75],
      ['numeric lte',        '=SUMIF(C2:C7,"<=150",D2:D7)',     25],
      ['not equal',          '=SUMIF(C2:C7,"<>100",D2:D7)',     107],
      ['wildcard *',         '=SUMIF(A2:A7,"A*",C2:C7)',        425],
      ['wildcard ?',         '=SUMIF(A2:A7,"?pple",C2:C7)',     425],
      ['wildcard end',       '=SUMIF(A2:A7,"*rry",C2:C7)',      300],
      ['no sum_range',       '=SUMIF(E1:E5,">25")',             120],
      ['case insensitive',   '=SUMIF(A2:A7,"APPLE",C2:C7)',     425],
      ['case lower',         '=SUMIF(A2:A7,"apple",C2:C7)',     425],
      ['no match',           '=SUMIF(A2:A7,"Orange",C2:C7)',    0],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('SUMIFS', () => {
    it.each([
      ['two criteria',        '=SUMIFS(C2:C7,A2:A7,"Apple",B2:B7,"North")',          275],
      ['two criteria 2',      '=SUMIFS(C2:C7,A2:A7,"Banana",B2:B7,"North")',         250],
      ['two criteria 3',      '=SUMIFS(C2:C7,A2:A7,"Apple",B2:B7,"South")',           150],
      ['numeric criteria',    '=SUMIFS(D2:D7,C2:C7,">100",B2:B7,"North")',            72],
      ['mixed criteria',      '=SUMIFS(C2:C7,A2:A7,"Apple",C2:C7,">100")',            325],
      ['wildcard',            '=SUMIFS(C2:C7,A2:A7,"*an*",B2:B7,"North")',            250],
      ['three criteria',      '=SUMIFS(D2:D7,A2:A7,"Apple",B2:B7,"North",C2:C7,">150")', 17],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('AVERAGEIF', () => {
    it('text match Apple', async () => {
      expectApprox(await evalFormula('=AVERAGEIF(A2:A7,"Apple",C2:C7)'), 141.6667);
    });
    it('text Banana', async () => {
      expect(await evalFormula('=AVERAGEIF(A2:A7,"Banana",C2:C7)')).toBe(225);
    });
    it('numeric gt', async () => {
      expect(await evalFormula('=AVERAGEIF(C2:C7,">200",D2:D7)')).toBe(27.5);
    });
    it('numeric lte', async () => {
      expect(await evalFormula('=AVERAGEIF(C2:C7,"<=150",D2:D7)')).toBe(12.5);
    });
    it('no avg_range', async () => {
      expect(await evalFormula('=AVERAGEIF(E1:E5,">25")')).toBe(40);
    });
    it('wildcard', async () => {
      expect(await evalFormula('=AVERAGEIF(A2:A7,"B*",C2:C7)')).toBe(225);
    });
    it('no match → #DIV/0!', async () => {
      expect(await evalFormula('=AVERAGEIF(A2:A7,"Orange",C2:C7)')).toBe('#DIV/0!');
    });
  });

  describe('AVERAGEIFS', () => {
    it('two criteria', async () => {
      expect(await evalFormula('=AVERAGEIFS(C2:C7,A2:A7,"Apple",B2:B7,"North")')).toBe(137.5);
    });
    it('two criteria 2', async () => {
      expect(await evalFormula('=AVERAGEIFS(C2:C7,A2:A7,"Banana",B2:B7,"South")')).toBe(200);
    });
    it('numeric criteria', async () => {
      expect(await evalFormula('=AVERAGEIFS(D2:D7,C2:C7,">100",B2:B7,"North")')).toBe(24);
    });
    it('no match → #DIV/0!', async () => {
      expect(await evalFormula('=AVERAGEIFS(C2:C7,A2:A7,"Orange",B2:B7,"North")')).toBe('#DIV/0!');
    });
  });

  describe('Edge cases', () => {
    it.each([
      ['SUMIF blank criteria',     '=SUMIF(G1:G5,"",H1:H5)',  5],
      ['SUMIF text in numeric',    '=SUMIF(G1:G5,">50",H1:H5)', 8],
      ['IF with SUMIF',            '=IF(SUMIF(A2:A7,"Apple",C2:C7)>400,"High","Low")', 'High'],
      ['IFERROR with AVERAGEIF',   '=IFERROR(AVERAGEIF(A2:A7,"Orange",C2:C7),"None")', 'None'],
      ['SUMIF single cell',        '=SUMIF(A2:A2,"Apple",C2:C2)',      100],
      ['AVERAGEIF single cell',    '=AVERAGEIF(A2:A2,"Apple",C2:C2)',  100],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });
});
