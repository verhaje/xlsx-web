import { describe, it, expect } from 'vitest';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';

const engine = createFormulaEngine({});

const cells: Record<string, any> = {
  A1: 10, A2: 20, A3: 30, A4: 40, A5: 50,
  B1: 5, B2: 15, B3: 25, B4: 35, B5: 45,
  C1: 'Hello', C2: 'World', C3: '', C4: 'test', C5: 'HELLO',
  D1: true, D2: false, D3: 0, D4: '', D5: null,
};

const resolveCell = async (ref: string) => cells[ref] ?? '';

async function evalFormula(formula: string) {
  return engine.evaluateFormula(formula, { resolveCell });
}

// =====================================================================
// Math functions
// =====================================================================
describe('High-Priority Math Functions', () => {
  it.each([
    ['ROUNDUP(3.2,0)', '=ROUNDUP(3.2,0)', 4],
    ['ROUNDUP(3.14159,3)', '=ROUNDUP(3.14159,3)', 3.142],
    ['ROUNDUP(-3.2,0)', '=ROUNDUP(-3.2,0)', -4],
    ['ROUNDDOWN(3.9,0)', '=ROUNDDOWN(3.9,0)', 3],
    ['ROUNDDOWN(3.14159,3)', '=ROUNDDOWN(3.14159,3)', 3.141],
    ['ROUNDDOWN(-3.9,0)', '=ROUNDDOWN(-3.9,0)', -3],
    ['INT(5.7)', '=INT(5.7)', 5],
    ['INT(-5.7)', '=INT(-5.7)', -6],
    ['SIGN(10)', '=SIGN(10)', 1],
    ['SIGN(-5)', '=SIGN(-5)', -1],
    ['SIGN(0)', '=SIGN(0)', 0],
    ['PI()', '=PI()', Math.PI],
    ['LOG(100)', '=LOG(100)', 2],
    ['LOG(8,2)', '=LOG(8,2)', 3],
    ['LOG10(1000)', '=LOG10(1000)', 3],
    ['EXP(1)', '=ROUND(EXP(1),5)', 2.71828],
    ['TRUNC(4.9)', '=TRUNC(4.9)', 4],
    ['TRUNC(-4.9)', '=TRUNC(-4.9)', -4],
    ['TRUNC(3.14159,2)', '=TRUNC(3.14159,2)', 3.14],
    ['EVEN(3)', '=EVEN(3)', 4],
    ['EVEN(2)', '=EVEN(2)', 2],
    ['EVEN(-1)', '=EVEN(-1)', -2],
    ['ODD(2)', '=ODD(2)', 3],
    ['ODD(3)', '=ODD(3)', 3],
    ['ODD(0)', '=ODD(0)', 1],
    ['GCD(12,8)', '=GCD(12,8)', 4],
    ['GCD(5,10,15)', '=GCD(5,10,15)', 5],
    ['LCM(4,6)', '=LCM(4,6)', 12],
    ['MROUND(10,3)', '=MROUND(10,3)', 9],
    ['MROUND(7.5,5)', '=MROUND(7.5,5)', 10],
    ['QUOTIENT(10,3)', '=QUOTIENT(10,3)', 3],
    ['QUOTIENT(-10,3)', '=QUOTIENT(-10,3)', -3],
    ['RADIANS(180)', '=ROUND(RADIANS(180),5)', 3.14159],
    ['DEGREES(PI())', '=DEGREES(PI())', 180],
    ['PRODUCT(2,3,4)', '=PRODUCT(2,3,4)', 24],
    ['FACT(5)', '=FACT(5)', 120],
    ['FACT(0)', '=FACT(0)', 1],
    ['COMBIN(10,3)', '=COMBIN(10,3)', 120],
    ['SUMSQ(1,2,3)', '=SUMSQ(1,2,3)', 14],
  ])('%s', async (_name, formula, expected) => {
    expect(await evalFormula(formula)).toEqual(expected);
  });

  it('RAND returns between 0 and 1', async () => {
    const val = await evalFormula('=RAND()');
    expect(val).toBeGreaterThanOrEqual(0);
    expect(val).toBeLessThan(1);
  });

  it('RANDBETWEEN returns within range', async () => {
    const val = await evalFormula('=RANDBETWEEN(1,10)') as number;
    expect(val).toBeGreaterThanOrEqual(1);
    expect(val).toBeLessThanOrEqual(10);
    expect(Number.isInteger(val)).toBe(true);
  });

  it('CEILING(4.2,1) = 5', async () => {
    expect(await evalFormula('=CEILING(4.2,1)')).toBe(5);
  });

  it('CEILING(-4.2,-1) = -5', async () => {
    expect(await evalFormula('=CEILING(-4.2,-1)')).toBe(-5);
  });

  it('FLOOR(4.9,1) = 4', async () => {
    expect(await evalFormula('=FLOOR(4.9,1)')).toBe(4);
  });

  it('FLOOR(-4.9,-1) = -4', async () => {
    expect(await evalFormula('=FLOOR(-4.9,-1)')).toBe(-4);
  });

  it('SUBTOTAL 9 (SUM) of args', async () => {
    expect(await evalFormula('=SUBTOTAL(9,1,2,3)')).toBe(6);
  });

  it('SUBTOTAL 1 (AVERAGE) of args', async () => {
    expect(await evalFormula('=SUBTOTAL(1,10,20,30)')).toBe(20);
  });

  it('SUBTOTAL 2 (COUNT) of args', async () => {
    expect(await evalFormula('=SUBTOTAL(2,10,20,30)')).toBe(3);
  });
});

// =====================================================================
// Logical functions
// =====================================================================
describe('High-Priority Logical Functions', () => {
  it.each([
    ['IFNA on non-NA', '=IFNA(5,"fallback")', 5],
    ['IFNA on #N/A', '=IFNA(MATCH(999,{1,2,3},0),"Not found")', 'Not found'],
    ['IFS first true', '=IFS(FALSE,"a",TRUE,"b",TRUE,"c")', 'b'],
    ['IFS none true → #N/A', '=IFS(FALSE,"a",FALSE,"b")', '#N/A'],
    ['SWITCH match', '=SWITCH(2,1,"a",2,"b",3,"c")', 'b'],
    ['SWITCH default', '=SWITCH(99,1,"a",2,"b","default")', 'default'],
    ['SWITCH no match → #N/A', '=SWITCH(99,1,"a",2,"b")', '#N/A'],
    ['XOR(TRUE,FALSE)', '=XOR(TRUE,FALSE)', true],
    ['XOR(TRUE,TRUE)', '=XOR(TRUE,TRUE)', false],
    ['XOR(FALSE,FALSE)', '=XOR(FALSE,FALSE)', false],
    ['XOR(TRUE,TRUE,TRUE)', '=XOR(TRUE,TRUE,TRUE)', true],
  ])('%s', async (_name, formula, expected) => {
    expect(await evalFormula(formula)).toEqual(expected);
  });
});

// =====================================================================
// Lookup functions
// =====================================================================
describe('High-Priority Lookup Functions', () => {
  it.each([
    ['CHOOSE(2,"a","b","c")', '=CHOOSE(2,"a","b","c")', 'b'],
    ['CHOOSE(1,"x","y")', '=CHOOSE(1,"x","y")', 'x'],
  ])('%s', async (_name, formula, expected) => {
    expect(await evalFormula(formula)).toEqual(expected);
  });

  it('CHOOSE out of range → #VALUE!', async () => {
    expect(await evalFormula('=CHOOSE(5,"a","b")')).toBe('#VALUE!');
  });

  it('ADDRESS(1,1) → $A$1', async () => {
    expect(await evalFormula('=ADDRESS(1,1)')).toBe('$A$1');
  });

  it('ADDRESS(2,3,4) → C2 (relative)', async () => {
    expect(await evalFormula('=ADDRESS(2,3,4)')).toBe('C2');
  });

  it('ADDRESS(1,27) → $AA$1', async () => {
    expect(await evalFormula('=ADDRESS(1,27)')).toBe('$AA$1');
  });
});

// =====================================================================
// Statistical functions
// =====================================================================
describe('High-Priority Statistical Functions', () => {
  it.each([
    ['COUNTA(1,"a",TRUE)', '=COUNTA(1,"a",TRUE)', 3],
    ['COUNTBLANK on empty', '=COUNTBLANK("")', 1],
    ['MODE(1,2,2,3)', '=MODE(1,2,2,3)', 2],
    ['MODE.SNGL(1,1,2,3)', '=MODE.SNGL(1,1,2,3)', 1],
    ['MODE all unique → #N/A', '=MODE(1,2,3)', '#N/A'],
  ])('%s', async (_name, formula, expected) => {
    expect(await evalFormula(formula)).toEqual(expected);
  });

  it('STDEV.S sample standard deviation', async () => {
    const result = await evalFormula('=ROUND(STDEV.S(2,4,4,4,5,5,7,9),6)');
    expect(result).toBe(2.138090);
  });

  it('STDEV.P population standard deviation', async () => {
    const result = await evalFormula('=ROUND(STDEV.P(2,4,4,4,5,5,7,9),6)');
    expect(result).toBe(2);
  });

  it('VAR sample variance', async () => {
    const result = await evalFormula('=ROUND(VAR(2,4,4,4,5,5,7,9),4)');
    expect(result).toBeCloseTo(4.5714, 3);
  });

  it('VAR.P population variance', async () => {
    const result = await evalFormula('=VAR.P(2,4,4,4,5,5,7,9)');
    expect(result).toBe(4);
  });

  it('AVERAGEA includes text as 0 and TRUE as 1', async () => {
    const result = await evalFormula('=AVERAGEA(1,TRUE,"text")');
    // 1 + 1 + 0 = 2, count=3 → 2/3
    expect(result).toBeCloseTo(2 / 3, 10);
  });

  it('MAXA with boolean', async () => {
    expect(await evalFormula('=MAXA(0,FALSE,TRUE)')).toBe(1);
  });

  it('MINA with boolean', async () => {
    expect(await evalFormula('=MINA(1,TRUE,FALSE)')).toBe(0);
  });
});

// =====================================================================
// Date/Time functions  
// =====================================================================
describe('High-Priority Date/Time Functions', () => {
  it('NOW returns a number > TODAY', async () => {
    const today = await evalFormula('=TODAY()') as number;
    const now = await evalFormula('=NOW()') as number;
    expect(now).toBeGreaterThanOrEqual(today);
    expect(now).toBeLessThan(today + 1);
  });

  it('TIME(12,30,0)', async () => {
    expect(await evalFormula('=TIME(12,30,0)')).toBeCloseTo(0.520833, 4);
  });

  it('HOUR(0.75) = 18', async () => {
    expect(await evalFormula('=HOUR(0.75)')).toBe(18);
  });

  it('MINUTE(0.75) = 0', async () => {
    expect(await evalFormula('=MINUTE(0.75)')).toBe(0);
  });

  it('SECOND(0.7006944) ≈ 59', async () => {
    expect(await evalFormula('=SECOND(0.7006944)')).toBe(59);
  });

  it('DAYS(DATE(2025,6,15),DATE(2025,1,1))', async () => {
    const result = await evalFormula('=DAYS(DATE(2025,6,15),DATE(2025,1,1))');
    expect(result).toBe(165);
  });

  it('EDATE(DATE(2025,1,31),1) = end of Feb', async () => {
    const result = await evalFormula('=DAY(EDATE(DATE(2025,1,31),1))');
    expect(result).toBe(28);
  });

  it('EOMONTH(DATE(2025,1,15),0) = Jan 31 serial', async () => {
    const result = await evalFormula('=DAY(EOMONTH(DATE(2025,1,15),0))');
    expect(result).toBe(31);
  });

  it('EOMONTH(DATE(2025,1,15),1) = Feb 28 serial', async () => {
    const result = await evalFormula('=DAY(EOMONTH(DATE(2025,1,15),1))');
    expect(result).toBe(28);
  });

  it('WEEKDAY(DATE(2025,2,15)) with default type', async () => {
    // Feb 15, 2025 is a Saturday → type 1 returns 7
    const result = await evalFormula('=WEEKDAY(DATE(2025,2,15))');
    expect(result).toBe(7);
  });

  it('WEEKDAY(DATE(2025,2,15),2) Mon=1', async () => {
    // Saturday → type 2 returns 6
    const result = await evalFormula('=WEEKDAY(DATE(2025,2,15),2)');
    expect(result).toBe(6);
  });

  it('DATEDIF years', async () => {
    const result = await evalFormula('=DATEDIF(DATE(2020,1,1),DATE(2025,6,15),"Y")');
    expect(result).toBe(5);
  });

  it('DATEDIF months', async () => {
    const result = await evalFormula('=DATEDIF(DATE(2020,1,1),DATE(2025,6,15),"M")');
    expect(result).toBe(65);
  });

  it('DATEDIF days', async () => {
    const result = await evalFormula('=DATEDIF(DATE(2025,1,1),DATE(2025,1,31),"D")');
    expect(result).toBe(30);
  });

  it('YEARFRAC basis 0', async () => {
    const result = await evalFormula('=ROUND(YEARFRAC(DATE(2025,1,1),DATE(2025,7,1),0),4)');
    expect(result).toBeCloseTo(0.5, 1);
  });
});

// =====================================================================
// Text functions
// =====================================================================
describe('High-Priority Text Functions', () => {
  it.each([
    ['SEARCH case-insensitive', '=SEARCH("world","Hello World")', 7],
    ['REPLACE', '=REPLACE("Hello",1,5,"Hi")', 'Hi'],
    ['REPLACE mid', '=REPLACE("abcdef",3,2,"XY")', 'abXYef'],
    ['REPT', '=REPT("ab",3)', 'ababab'],
    ['REPT 0 times', '=REPT("ab",0)', ''],
    ['EXACT match', '=EXACT("Hello","Hello")', true],
    ['EXACT mismatch', '=EXACT("Hello","hello")', false],
    ['CHAR(65)', '=CHAR(65)', 'A'],
    ['CODE("A")', '=CODE("A")', 65],
    ['CLEAN', '=CLEAN("test")', 'test'],
    ['PROPER', '=PROPER("hello world")', 'Hello World'],
    ['PROPER mixed', '=PROPER("hELLO wORLD")', 'Hello World'],
    ['FIXED(1234.567,2)', '=FIXED(1234.567,2,TRUE)', '1234.57'],
    ['NUMBERVALUE("1,234.56")', '=NUMBERVALUE("1,234.56",".",",")', 1234.56],
    ['NUMBERVALUE with %', '=NUMBERVALUE("50%")', 0.5],
    ['TEXTBEFORE', '=TEXTBEFORE("Hello-World","-")', 'Hello'],
    ['TEXTAFTER', '=TEXTAFTER("Hello-World","-")', 'World'],
    ['TEXTBEFORE 2nd instance', '=TEXTBEFORE("a-b-c","-",2)', 'a-b'],
    ['TEXTAFTER 2nd instance', '=TEXTAFTER("a-b-c","-",2)', 'c'],
  ])('%s', async (_name, formula, expected) => {
    expect(await evalFormula(formula)).toEqual(expected);
  });

  it('TEXTJOIN with delimiter', async () => {
    expect(await evalFormula('=TEXTJOIN(",",TRUE,"a","b","c")')).toBe('a,b,c');
  });

  it('TEXTJOIN ignore empty', async () => {
    expect(await evalFormula('=TEXTJOIN(",",TRUE,"a","","c")')).toBe('a,c');
  });

  it('TEXTJOIN keep empty', async () => {
    expect(await evalFormula('=TEXTJOIN(",",FALSE,"a","","c")')).toBe('a,,c');
  });

  it('DOLLAR formatting', async () => {
    const result = await evalFormula('=DOLLAR(1234.567,2)');
    expect(result).toBe('$1,234.57');
  });

  it('DOLLAR negative', async () => {
    const result = await evalFormula('=DOLLAR(-1234.567,2)');
    expect(result).toBe('($1,234.57)');
  });
});

// =====================================================================
// Information functions
// =====================================================================
describe('High-Priority Information Functions', () => {
  it.each([
    ['TYPE number', '=TYPE(1)', 1],
    ['TYPE text', '=TYPE("hello")', 2],
    ['TYPE boolean', '=TYPE(TRUE)', 4],
    ['N(TRUE)', '=N(TRUE)', 1],
    ['N(FALSE)', '=N(FALSE)', 0],
    ['N(5)', '=N(5)', 5],
    ['N("text")', '=N("text")', 0],
    ['NA()', '=NA()', '#N/A'],
    ['ERROR.TYPE(#N/A)', '=ERROR.TYPE(NA())', 7],
    ['ISEVEN(4)', '=ISEVEN(4)', true],
    ['ISEVEN(3)', '=ISEVEN(3)', false],
    ['ISODD(3)', '=ISODD(3)', true],
    ['ISODD(4)', '=ISODD(4)', false],
    ['ISLOGICAL(TRUE)', '=ISLOGICAL(TRUE)', true],
    ['ISLOGICAL(1)', '=ISLOGICAL(1)', false],
    ['ISNA(NA())', '=ISNA(NA())', true],
    ['ISNA(5)', '=ISNA(5)', false],
    ['ISNONTEXT(5)', '=ISNONTEXT(5)', true],
    ['ISNONTEXT("hi")', '=ISNONTEXT("hi")', false],
    ['ISERR(1/0)', '=ISERR(1/0)', true],
  ])('%s', async (_name, formula, expected) => {
    expect(await evalFormula(formula)).toEqual(expected);
  });
});

// =====================================================================
// XLOOKUP and XMATCH
// =====================================================================
describe('XLOOKUP and XMATCH', () => {
  it('XLOOKUP exact match', async () => {
    expect(await evalFormula('=XLOOKUP(2,{1,2,3},{"a","b","c"})')).toBe('b');
  });

  it('XLOOKUP not found → default', async () => {
    expect(await evalFormula('=XLOOKUP(99,{1,2,3},{"a","b","c"},"missing")')).toBe('missing');
  });

  it('XLOOKUP not found → #N/A default', async () => {
    expect(await evalFormula('=XLOOKUP(99,{1,2,3},{"a","b","c"})')).toBe('#N/A');
  });

  it('XMATCH exact match', async () => {
    expect(await evalFormula('=XMATCH(2,{1,2,3})')).toBe(2);
  });

  it('XMATCH not found', async () => {
    expect(await evalFormula('=XMATCH(99,{1,2,3})')).toBe('#N/A');
  });
});

// =====================================================================
// SUMPRODUCT
// =====================================================================
describe('SUMPRODUCT', () => {
  it('basic SUMPRODUCT', async () => {
    expect(await evalFormula('=SUMPRODUCT({1,2,3},{4,5,6})')).toBe(32); // 1*4+2*5+3*6
  });
});

// =====================================================================
// COUNTIF
// =====================================================================
describe('COUNTIF', () => {
  it('COUNTIF counts matching values', async () => {
    expect(await evalFormula('=COUNTIF({1,2,2,3,2},2)')).toBe(3);
  });

  it('COUNTIF with criteria string', async () => {
    expect(await evalFormula('=COUNTIF({1,2,3,4,5},">3")')).toBe(2);
  });
});

// =====================================================================
// LARGE / SMALL
// =====================================================================
describe('LARGE and SMALL', () => {
  it('LARGE 2nd largest', async () => {
    expect(await evalFormula('=LARGE({3,1,4,1,5,9},2)')).toBe(5);
  });

  it('SMALL 2nd smallest', async () => {
    expect(await evalFormula('=SMALL({3,1,4,1,5,9},2)')).toBe(1);
  });

  it('LARGE k out of range → #NUM!', async () => {
    expect(await evalFormula('=LARGE({1,2,3},5)')).toBe('#NUM!');
  });
});

// =====================================================================
// RANK
// =====================================================================
describe('RANK', () => {
  it('RANK descending', async () => {
    expect(await evalFormula('=RANK(3,{1,2,3,4,5})')).toBe(3);
  });

  it('RANK ascending', async () => {
    expect(await evalFormula('=RANK(3,{1,2,3,4,5},1)')).toBe(3);
  });

  it('RANK.EQ', async () => {
    expect(await evalFormula('=RANK.EQ(5,{1,2,3,4,5})')).toBe(1);
  });
});

// =====================================================================
// MAXIFS / MINIFS
// =====================================================================
describe('MAXIFS and MINIFS', () => {
  it('MAXIFS with criteria', async () => {
    expect(await evalFormula('=MAXIFS({10,20,30,40},{1,2,1,2},1)')).toBe(30);
  });

  it('MINIFS with criteria', async () => {
    expect(await evalFormula('=MINIFS({10,20,30,40},{1,2,1,2},1)')).toBe(10);
  });
});

// =====================================================================
// PERCENTILE / QUARTILE
// =====================================================================
describe('PERCENTILE and QUARTILE', () => {
  it('PERCENTILE 0.5 is median', async () => {
    expect(await evalFormula('=PERCENTILE({1,2,3,4,5},0.5)')).toBe(3);
  });

  it('QUARTILE Q2 is median', async () => {
    expect(await evalFormula('=QUARTILE({1,2,3,4,5},2)')).toBe(3);
  });

  it('PERCENTILE out of range → #NUM!', async () => {
    expect(await evalFormula('=PERCENTILE({1,2,3},1.5)')).toBe('#NUM!');
  });
});
