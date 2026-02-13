import { describe, it, expect } from 'vitest';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';

const engine = createFormulaEngine({});

const cells: Record<string, any> = {
  A1: 10, A2: 20, A3: 30,
  B1: 5,  B2: 15, B3: 25,
  C1: 'Hello', C2: 'World',
  D1: '',
  G8: 'test',
};

const resolveCell = async (ref: string) => cells[ref] ?? '';

async function evalFormula(formula: string) {
  return engine.evaluateFormula(formula, { resolveCell });
}

describe('Formula Engine', () => {
  describe('Arithmetic', () => {
    it.each([
      ['1+2',        '=1+2',          3],
      ['10-3',       '=10-3',         7],
      ['4*5',        '=4*5',          20],
      ['20/4',       '=20/4',         5],
      ['1/0',        '=1/0',          '#DIV/0!'],
      ['2^3',        '=2^3',          8],
      ['(1+2)*3',    '=(1+2)*3',      9],
      ['-5',         '=-5',           -5],
      ['1+2*3-4/2',  '=1+2*3-4/2',   5],
    ])('%s → %s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('Cell references', () => {
    it('resolves a cell ref', async () => {
      expect(await evalFormula('=A1')).toBe(10);
    });
    it('adds two cells', async () => {
      expect(await evalFormula('=A1+A2')).toBe(30);
    });
    it('sums a range', async () => {
      expect(await evalFormula('=SUM(A1:A3)')).toBe(60);
    });
  });

  describe('Math functions', () => {
    it.each([
      ['SUM',    '=SUM(1,2,3)',          6],
      ['AVERAGE','=AVERAGE(10,20,30)',    20],
      ['MIN',    '=MIN(5,10,3)',          3],
      ['MAX',    '=MAX(5,10,3)',          10],
      ['COUNT',  '=COUNT(1,2,3)',         3],
      ['ABS',    '=ABS(-5)',             5],
      ['ROUND',  '=ROUND(3.567,2)',      3.57],
      ['SQRT',   '=SQRT(16)',            4],
      ['POWER',  '=POWER(2,3)',          8],
      ['MOD',    '=MOD(10,3)',           1],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('LN function', () => {
    it.each([
      ['LN(1)',                   '=LN(1)',                   0],
      ['LN(10) rounded',         '=ROUND(LN(10),5)',         2.30259],
      ['LN text numeric',        '=ROUND(LN("2"),5)',        0.69315],
      ['LN(0) → #NUM!',         '=LN(0)',                   '#NUM!'],
      ['LN(-1) → #NUM!',        '=LN(-1)',                  '#NUM!'],
      ['LN("abc") → #VALUE!',   '=LN("abc")',               '#VALUE!'],
      ['LN("abc") in ROUND',    '=ROUND(LN("abc"),0)',      '#VALUE!'],
      ['LN("abc") in expr',     '=LN("abc")+6',            '#VALUE!'],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('COUNTIFS', () => {
    it.each([
      ['basic',       '=COUNTIFS(A1:A3,">15",B1:B3,">10")', 2],
      ['wildcard',    '=COUNTIFS(C1:C2,"H*")',                1],
      ['blank',       '=COUNTIFS(D1:D1,"")',                  1],
      ['cell crit',   '=COUNTIFS(A1:A3,A1)',                  1],
      ['large range', '=COUNTIFS(A1:A100,10)',                1],
      ['no $ ref',    '=COUNTIFS(A1:A3,G8)',                  0],
      ['with $ ref',  '=COUNTIFS(A1:A3,$G8)',                 0],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('Logical', () => {
    it.each([
      ['IF true',   '=IF(1>0,"yes","no")',  'yes'],
      ['IF false',  '=IF(1<0,"yes","no")',  'no'],
      ['AND true',  '=AND(1,1)',            true],
      ['AND false', '=AND(1,0)',            false],
      ['OR true',   '=OR(0,1)',             true],
      ['OR false',  '=OR(0,0)',             false],
      ['NOT',       '=NOT(0)',              true],
      ['TRUE',      '=TRUE()',              true],
      ['FALSE',     '=FALSE()',             false],
      ['IFERROR',   '=IFERROR(1/0,"error")','error'],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('Text', () => {
    it.each([
      ['CONCAT',          '=CONCAT("a","b","c")',  'abc'],
      ['LEFT',            '=LEFT("Hello",2)',       'He'],
      ['RIGHT',           '=RIGHT("Hello",2)',      'lo'],
      ['MID',             '=MID("Hello",2,3)',      'ell'],
      ['LEN',             '=LEN("Hello")',          5],
      ['LOWER',           '=LOWER("HELLO")',        'hello'],
      ['UPPER',           '=UPPER("hello")',        'HELLO'],
      ['TRIM',            '=TRIM("  hi  ")',        'hi'],
      ['Concatenation &', '="a"&"b"',               'ab'],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('Comparison', () => {
    it.each([
      ['Greater than', '=5>3',   true],
      ['Less than',    '=3<5',   true],
      ['Equal',        '=5=5',   true],
      ['Not equal',    '=5<>3',  true],
      ['GTE',          '=5>=5',  true],
      ['LTE',          '=3<=5',  true],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('Error / type checking', () => {
    it.each([
      ['ISBLANK true',   '=ISBLANK(D1)',      true],
      ['ISBLANK false',  '=ISBLANK(A1)',       false],
      ['ISERROR',        '=ISERROR(1/0)',      true],
      ['ISNUMBER num',   '=ISNUMBER(123)',     true],
      ['ISNUMBER text',  '=ISNUMBER("123")',   false],
      ['ISTEXT text',    '=ISTEXT("abc")',     true],
      ['ISTEXT num',     '=ISTEXT(123)',       false],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('VALUE / SUBSTITUTE / FIND', () => {
    it.each([
      ['VALUE number',      '=VALUE("123.45")',                            123.45],
      ['VALUE invalid',     '=VALUE("abc")',                               '#VALUE!'],
      ['SUBSTITUTE all',    '=SUBSTITUTE("Hello World","o","0")',          'Hell0 W0rld'],
      ['SUBSTITUTE nth',    '=SUBSTITUTE("ababa","a","x",2)',              'abxba'],
      ['FIND found',        '=FIND("lo","Hello")',                         4],
      ['FIND not found',    '=FIND("z","Hello")',                          '#VALUE!'],
      ['FIND start',        '=FIND("l","Hello",4)',                        4],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });

  describe('Statistical', () => {
    it.each([
      ['MEDIAN odd',    '=MEDIAN(1,3,2)',     2],
      ['MEDIAN even',   '=MEDIAN(1,2,3,4)',   2.5],
      ['MEDIAN range',  '=MEDIAN(A1:A3)',     20],
      ['STDEV sample',  '=STDEV(1,2,3)',      1],
      ['STDEV single',  '=STDEV(1)',          '#DIV/0!'],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalFormula(formula)).toEqual(expected);
    });
  });
});
