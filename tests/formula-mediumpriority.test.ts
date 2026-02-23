// Medium-priority formula function tests
import { describe, it, expect } from 'vitest';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';

const engine = createFormulaEngine({});

const resolveCell = async (_ref: string) => '';

async function evalFormula(formula: string) {
  return engine.evaluateFormula(formula, { resolveCell });
}

// ============================================================
// Math / Trigonometry
// ============================================================

describe('Trigonometric functions', () => {
  it('SIN returns sine of angle', async () => {
    expect(await evalFormula('=SIN(0)')).toBeCloseTo(0, 10);
    expect(await evalFormula('=SIN(1)')).toBeCloseTo(Math.sin(1), 10);
    expect(await evalFormula('=SIN(PI()/2)')).toBeCloseTo(1, 10);
  });

  it('COS returns cosine of angle', async () => {
    expect(await evalFormula('=COS(0)')).toBeCloseTo(1, 10);
    expect(await evalFormula('=COS(PI())')).toBeCloseTo(-1, 10);
  });

  it('TAN returns tangent of angle', async () => {
    expect(await evalFormula('=TAN(0)')).toBeCloseTo(0, 10);
    expect(await evalFormula('=TAN(PI()/4)')).toBeCloseTo(1, 10);
  });

  it('ASIN returns arcsine', async () => {
    expect(await evalFormula('=ASIN(0)')).toBeCloseTo(0, 10);
    expect(await evalFormula('=ASIN(1)')).toBeCloseTo(Math.PI / 2, 10);
    expect(await evalFormula('=ASIN(2)')).toBe('#NUM!');
  });

  it('ACOS returns arccosine', async () => {
    expect(await evalFormula('=ACOS(1)')).toBeCloseTo(0, 10);
    expect(await evalFormula('=ACOS(0)')).toBeCloseTo(Math.PI / 2, 10);
    expect(await evalFormula('=ACOS(-2)')).toBe('#NUM!');
  });

  it('ATAN returns arctangent', async () => {
    expect(await evalFormula('=ATAN(0)')).toBeCloseTo(0, 10);
    expect(await evalFormula('=ATAN(1)')).toBeCloseTo(Math.PI / 4, 10);
  });

  it('ATAN2 returns arctangent from x,y', async () => {
    expect(await evalFormula('=ATAN2(1,1)')).toBeCloseTo(Math.PI / 4, 10);
    expect(await evalFormula('=ATAN2(0,0)')).toBe('#DIV/0!');
  });

  it('SINH/COSH/TANH hyperbolic functions', async () => {
    expect(await evalFormula('=SINH(0)')).toBeCloseTo(0, 10);
    expect(await evalFormula('=COSH(0)')).toBeCloseTo(1, 10);
    expect(await evalFormula('=TANH(0)')).toBeCloseTo(0, 10);
    expect(await evalFormula('=SINH(1)')).toBeCloseTo(Math.sinh(1), 10);
    expect(await evalFormula('=COSH(1)')).toBeCloseTo(Math.cosh(1), 10);
    expect(await evalFormula('=TANH(1)')).toBeCloseTo(Math.tanh(1), 10);
  });

  it('ACOT returns arccotangent', async () => {
    expect(await evalFormula('=ACOT(1)')).toBeCloseTo(Math.PI / 4, 10);
  });

  it('COT returns cotangent', async () => {
    expect(await evalFormula('=COT(0)')).toBe('#DIV/0!');
    expect(await evalFormula('=COT(1)')).toBeCloseTo(1 / Math.tan(1), 10);
  });

  it('CSC returns cosecant', async () => {
    expect(await evalFormula('=CSC(0)')).toBe('#DIV/0!');
    expect(await evalFormula('=CSC(1)')).toBeCloseTo(1 / Math.sin(1), 10);
  });

  it('SEC returns secant', async () => {
    expect(await evalFormula('=SEC(0)')).toBeCloseTo(1, 10);
    expect(await evalFormula('=SEC(1)')).toBeCloseTo(1 / Math.cos(1), 10);
  });

  it('SQRTPI returns sqrt(n*pi)', async () => {
    expect(await evalFormula('=SQRTPI(1)')).toBeCloseTo(Math.sqrt(Math.PI), 10);
    expect(await evalFormula('=SQRTPI(0)')).toBeCloseTo(0, 10);
    expect(await evalFormula('=SQRTPI(-1)')).toBe('#NUM!');
  });
});

describe('Combinatorics & special math', () => {
  it('PERMUT returns permutations', async () => {
    expect(await evalFormula('=PERMUT(5,3)')).toBe(60);
    expect(await evalFormula('=PERMUT(10,2)')).toBe(90);
    expect(await evalFormula('=PERMUT(3,5)')).toBe('#NUM!');
  });

  it('COMBINA returns combinations with repetition', async () => {
    expect(await evalFormula('=COMBINA(4,3)')).toBe(20);
    expect(await evalFormula('=COMBINA(10,3)')).toBe(220);
  });

  it('MULTINOMIAL returns multinomial', async () => {
    expect(await evalFormula('=MULTINOMIAL(2,3,4)')).toBe(1260);
  });

  it('FACTDOUBLE returns double factorial', async () => {
    expect(await evalFormula('=FACTDOUBLE(5)')).toBe(15); // 5*3*1
    expect(await evalFormula('=FACTDOUBLE(6)')).toBe(48); // 6*4*2
    expect(await evalFormula('=FACTDOUBLE(0)')).toBe(1);
    expect(await evalFormula('=FACTDOUBLE(-1)')).toBe(1);
  });

  it('SERIESSUM returns power series sum', async () => {
    // 1*x^1 + 2*x^2 + 3*x^3 with x=2, n=1, m=1
    // = 1*2 + 2*4 + 3*8 = 2 + 8 + 24 = 34
    expect(await evalFormula('=SERIESSUM(2,1,1,{1,2,3})')).toBeCloseTo(34, 10);
  });
});

describe('Number base & Roman numeral conversions', () => {
  it('BASE converts to specified base', async () => {
    expect(await evalFormula('=BASE(255,16)')).toBe('FF');
    expect(await evalFormula('=BASE(7,2)')).toBe('111');
    expect(await evalFormula('=BASE(7,2,6)')).toBe('000111');
  });

  it('DECIMAL converts from specified base', async () => {
    expect(await evalFormula('=DECIMAL("FF",16)')).toBe(255);
    expect(await evalFormula('=DECIMAL("111",2)')).toBe(7);
  });

  it('ROMAN converts number to Roman numeral', async () => {
    expect(await evalFormula('=ROMAN(499)')).toBe('CDXCIX');
    expect(await evalFormula('=ROMAN(2024)')).toBe('MMXXIV');
    expect(await evalFormula('=ROMAN(0)')).toBe('');
  });

  it('ARABIC converts Roman numeral to number', async () => {
    expect(await evalFormula('=ARABIC("MCMXC")')).toBe(1990);
    expect(await evalFormula('=ARABIC("MMXXIV")')).toBe(2024);
    expect(await evalFormula('=ARABIC("IV")')).toBe(4);
  });
});

describe('AGGREGATE function', () => {
  it('AGGREGATE with AVERAGE (fn=1)', async () => {
    expect(await evalFormula('=AGGREGATE(1,0,10,20,30)')).toBe(20);
  });
  it('AGGREGATE with SUM (fn=9)', async () => {
    expect(await evalFormula('=AGGREGATE(9,0,10,20,30)')).toBe(60);
  });
  it('AGGREGATE with MIN (fn=5)', async () => {
    expect(await evalFormula('=AGGREGATE(5,0,10,20,30)')).toBe(10);
  });
  it('AGGREGATE with MAX (fn=4)', async () => {
    expect(await evalFormula('=AGGREGATE(4,0,10,20,30)')).toBe(30);
  });
});

describe('Precise ceiling/floor', () => {
  it('CEILING.PRECISE rounds up', async () => {
    expect(await evalFormula('=CEILING.PRECISE(4.3,1)')).toBe(5);
    expect(await evalFormula('=CEILING.PRECISE(-4.3,1)')).toBe(-4);
  });
  it('FLOOR.PRECISE rounds down', async () => {
    expect(await evalFormula('=FLOOR.PRECISE(4.7,1)')).toBe(4);
    expect(await evalFormula('=FLOOR.PRECISE(-4.7,1)')).toBe(-5);
  });
});

describe('Sum array product functions', () => {
  it('SUMX2MY2 returns sum of x²-y²', async () => {
    expect(await evalFormula('=SUMX2MY2({1,2,3},{4,5,6})')).toBe(
      (1 - 16) + (4 - 25) + (9 - 36) // -15 + -21 + -27 = -63
    );
  });
  it('SUMX2PY2 returns sum of x²+y²', async () => {
    expect(await evalFormula('=SUMX2PY2({1,2,3},{4,5,6})')).toBe(
      (1 + 16) + (4 + 25) + (9 + 36) // 17 + 29 + 45 = 91
    );
  });
  it('SUMXMY2 returns sum of (x-y)²', async () => {
    expect(await evalFormula('=SUMXMY2({1,2,3},{4,5,6})')).toBe(
      9 + 9 + 9 // 27
    );
  });
});

// ============================================================
// Statistical functions
// ============================================================

describe('Regression / Correlation', () => {
  it('CORREL returns correlation coefficient', async () => {
    const r = await evalFormula('=CORREL({3,2,4,5,6},{9,7,12,15,17})');
    expect(r).toBeCloseTo(0.9970, 3);
  });

  it('COVARIANCE.P returns population covariance', async () => {
    const cv = await evalFormula('=COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})');
    expect(cv).toBeCloseTo(5.2, 1);
  });

  it('COVARIANCE.S returns sample covariance', async () => {
    const cv = await evalFormula('=COVARIANCE.S({3,2,4,5,6},{9,7,12,15,17})');
    expect(cv).toBeCloseTo(6.5, 1);
  });

  it('SLOPE returns slope of linear regression', async () => {
    const slope = await evalFormula('=SLOPE({9,7,12,15,17},{3,2,4,5,6})');
    expect(slope).toBeCloseTo(2.6, 1);
  });

  it('INTERCEPT returns y-intercept', async () => {
    const intercept = await evalFormula('=INTERCEPT({9,7,12,15,17},{3,2,4,5,6})');
    expect(intercept).toBeCloseTo(1.6, 1);
  });

  it('RSQ returns R-squared', async () => {
    const rsq = await evalFormula('=RSQ({9,7,12,15,17},{3,2,4,5,6})');
    expect(rsq).toBeCloseTo(0.9940, 3);
  });

  it('FORECAST predicts a value', async () => {
    const f = await evalFormula('=FORECAST(7,{9,7,12,15,17},{3,2,4,5,6})');
    expect(f).toBeCloseTo(19.8, 1);
  });

  it('FORECAST.LINEAR is same as FORECAST', async () => {
    const f = await evalFormula('=FORECAST.LINEAR(7,{9,7,12,15,17},{3,2,4,5,6})');
    expect(f).toBeCloseTo(19.8, 1);
  });
});

describe('Fisher transformation', () => {
  it('FISHER returns Fisher transform', async () => {
    expect(await evalFormula('=FISHER(0.75)')).toBeCloseTo(0.9730, 3);
    expect(await evalFormula('=FISHER(0)')).toBeCloseTo(0, 10);
  });

  it('FISHERINV returns inverse Fisher', async () => {
    expect(await evalFormula('=FISHERINV(0.9730)')).toBeCloseTo(0.75, 2);
    expect(await evalFormula('=FISHERINV(0)')).toBeCloseTo(0, 10);
  });

  it('FISHER returns error for out-of-range', async () => {
    expect(await evalFormula('=FISHER(1)')).toBe('#NUM!');
    expect(await evalFormula('=FISHER(-1)')).toBe('#NUM!');
  });
});

describe('Special means', () => {
  it('GEOMEAN returns geometric mean', async () => {
    expect(await evalFormula('=GEOMEAN(4,9)')).toBeCloseTo(6, 10);
    expect(await evalFormula('=GEOMEAN(1,2,3,4,5)')).toBeCloseTo(2.6052, 3);
  });

  it('HARMEAN returns harmonic mean', async () => {
    expect(await evalFormula('=HARMEAN(1,2,4)')).toBeCloseTo(12 / 7, 3);
  });

  it('TRIMMEAN returns trimmed mean', async () => {
    const r = await evalFormula('=TRIMMEAN({1,2,3,4,5,6,7,8,9,10},0.2)');
    // Trim 1 from each end: mean of 2-9 = 5.5
    expect(r).toBeCloseTo(5.5, 5);
  });
});

describe('Deviation measures', () => {
  it('DEVSQ returns sum of squared deviations', async () => {
    expect(await evalFormula('=DEVSQ(4,5,8,7,11,4,3)')).toBeCloseTo(48, 0);
  });

  it('AVEDEV returns average absolute deviation', async () => {
    expect(await evalFormula('=AVEDEV(4,5,8,7,11,4,3)')).toBeCloseTo(2.2857, 3);
  });
});

describe('Distribution shape (SKEW, KURT)', () => {
  it('SKEW returns skewness', async () => {
    const r = await evalFormula('=SKEW(3,4,5,2,3,4,5,6,4,7)');
    expect(typeof r).toBe('number');
    expect(r).toBeCloseTo(0.3595, 3);
  });

  it('KURT returns kurtosis', async () => {
    const r = await evalFormula('=KURT(3,4,5,2,3,4,5,6,4,7)');
    expect(typeof r).toBe('number');
  });

  it('SKEW with < 3 values returns error', async () => {
    expect(await evalFormula('=SKEW(1,2)')).toBe('#DIV/0!');
  });

  it('KURT with < 4 values returns error', async () => {
    expect(await evalFormula('=KURT(1,2,3)')).toBe('#DIV/0!');
  });
});

describe('STANDARDIZE', () => {
  it('returns normalized z-score', async () => {
    expect(await evalFormula('=STANDARDIZE(42,40,1.5)')).toBeCloseTo(4 / 3, 3);
  });
  it('errors on zero stddev', async () => {
    expect(await evalFormula('=STANDARDIZE(42,40,0)')).toBe('#NUM!');
  });
});

describe('Normal distribution', () => {
  it('NORM.S.DIST CDF at z=0 is 0.5', async () => {
    const r = await evalFormula('=NORM.S.DIST(0,TRUE)');
    expect(r).toBeCloseTo(0.5, 3);
  });

  it('NORM.S.DIST CDF at z=1 is ~0.8413', async () => {
    const r = await evalFormula('=NORM.S.DIST(1,TRUE)');
    expect(r).toBeCloseTo(0.8413, 3);
  });

  it('NORM.S.DIST PDF at z=0', async () => {
    const r = await evalFormula('=NORM.S.DIST(0,FALSE)');
    expect(r).toBeCloseTo(1 / Math.sqrt(2 * Math.PI), 5);
  });

  it('NORM.DIST with mean=0, sd=1 matches NORM.S.DIST', async () => {
    const r = await evalFormula('=NORM.DIST(1,0,1,TRUE)');
    expect(r).toBeCloseTo(0.8413, 3);
  });

  it('NORM.S.INV of 0.5 is 0', async () => {
    const r = await evalFormula('=NORM.S.INV(0.5)');
    expect(r).toBeCloseTo(0, 3);
  });

  it('NORM.S.INV of 0.975 is ~1.96', async () => {
    const r = await evalFormula('=NORM.S.INV(0.975)');
    expect(r).toBeCloseTo(1.96, 1);
  });

  it('NORM.INV transforms correctly', async () => {
    const r = await evalFormula('=NORM.INV(0.5,10,2)');
    expect(r).toBeCloseTo(10, 3);
  });

  it('NORM.S.INV errors for out-of-range', async () => {
    expect(await evalFormula('=NORM.S.INV(0)')).toBe('#NUM!');
    expect(await evalFormula('=NORM.S.INV(1)')).toBe('#NUM!');
  });
});

describe('Discrete distributions', () => {
  it('POISSON.DIST PMF', async () => {
    // P(X=2) when mean=5: e^(-5) * 5^2 / 2! = 0.0842
    const r = await evalFormula('=POISSON.DIST(2,5,FALSE)');
    expect(r).toBeCloseTo(0.0842, 3);
  });

  it('POISSON.DIST CDF', async () => {
    const r = await evalFormula('=POISSON.DIST(2,5,TRUE)');
    expect(r).toBeCloseTo(0.1247, 3);
  });

  it('BINOM.DIST PMF', async () => {
    // P(X=6) for n=10, p=0.5 = C(10,6)*0.5^10 = 0.2051
    const r = await evalFormula('=BINOM.DIST(6,10,0.5,FALSE)');
    expect(r).toBeCloseTo(0.2051, 3);
  });

  it('BINOM.DIST CDF', async () => {
    const r = await evalFormula('=BINOM.DIST(6,10,0.5,TRUE)');
    expect(r).toBeCloseTo(0.8281, 3);
  });
});

describe('Continuous distributions', () => {
  it('WEIBULL.DIST CDF', async () => {
    const r = await evalFormula('=WEIBULL.DIST(105,20,100,TRUE)');
    expect(r).toBeCloseTo(0.9295, 2);
  });

  it('EXPON.DIST CDF', async () => {
    // F(1) = 1 - e^(-0.5) = 0.3935
    const r = await evalFormula('=EXPON.DIST(1,0.5,TRUE)');
    expect(r).toBeCloseTo(0.3935, 3);
  });

  it('EXPON.DIST PDF', async () => {
    const r = await evalFormula('=EXPON.DIST(1,0.5,FALSE)');
    expect(r).toBeCloseTo(0.5 * Math.exp(-0.5), 5);
  });

  it('LOGNORM.DIST CDF', async () => {
    const r = await evalFormula('=LOGNORM.DIST(4,3.5,1.2,TRUE)');
    expect(typeof r).toBe('number');
    expect(r).toBeGreaterThan(0);
    expect(r).toBeLessThan(1);
  });
});

// ============================================================
// Dynamic array / Lookup
// ============================================================

describe('Dynamic array functions', () => {
  it('SORT sorts ascending by default', async () => {
    const r = await evalFormula('=SORT({3,1,4,1,5})');
    expect(r).toEqual([1, 1, 3, 4, 5]);
  });

  it('SORT sorts descending with order=-1', async () => {
    const r = await evalFormula('=SORT({3,1,4,1,5},1,-1)');
    expect(r).toEqual([5, 4, 3, 1, 1]);
  });

  it('UNIQUE returns unique values', async () => {
    const r = await evalFormula('=UNIQUE({1,2,2,3,3,3})');
    expect(r).toEqual([1, 2, 3]);
  });

  it('FILTER filters by boolean array', async () => {
    const r = await evalFormula('=FILTER({10,20,30},{1,0,1})');
    expect(r).toEqual([10, 30]);
  });

  it('FILTER returns if_empty when no match', async () => {
    const r = await evalFormula('=FILTER({10,20,30},{0,0,0},"None")');
    expect(r).toBe('None');
  });

  it('SORTBY sorts by another array', async () => {
    const r = await evalFormula('=SORTBY({"a","b","c"},{3,1,2})');
    expect(r).toEqual(['b', 'c', 'a']);
  });

  it('AREAS returns argument count', async () => {
    expect(await evalFormula('=AREAS(1)')).toBe(1);
  });
});

// ============================================================
// Text
// ============================================================

describe('Text medium-priority', () => {
  it('T returns text if text, empty otherwise', async () => {
    expect(await evalFormula('=T("hello")')).toBe('hello');
    expect(await evalFormula('=T(123)')).toBe('');
  });

  it('UNICHAR returns character from code', async () => {
    expect(await evalFormula('=UNICHAR(65)')).toBe('A');
    expect(await evalFormula('=UNICHAR(9731)')).toBe('☃');
  });

  it('UNICODE returns code from character', async () => {
    expect(await evalFormula('=UNICODE("A")')).toBe(65);
    expect(await evalFormula('=UNICODE("☃")')).toBe(9731);
  });

  it('TEXTSPLIT splits text by delimiter', async () => {
    const r = await evalFormula('=TEXTSPLIT("a,b,c",",")');
    expect(r).toEqual(['a', 'b', 'c']);
  });

  it('VALUETOTEXT converts value to text', async () => {
    expect(await evalFormula('=VALUETOTEXT(123)')).toBe('123');
    expect(await evalFormula('=VALUETOTEXT("hello")')).toBe('hello');
  });

  it('ARRAYTOTEXT joins array elements', async () => {
    const r = await evalFormula('=ARRAYTOTEXT({1,2,3})');
    expect(r).toBe('1, 2, 3');
  });

  it('ENCODEURL encodes URL', async () => {
    expect(await evalFormula('=ENCODEURL("hello world")')).toBe('hello%20world');
    expect(await evalFormula('=ENCODEURL("a&b=c")')).toBe('a%26b%3Dc');
  });
});

// ============================================================
// Financial
// ============================================================

describe('Time Value of Money', () => {
  it('FV calculates future value', async () => {
    // 5% annual, 10 years, $100/month
    const r = await evalFormula('=FV(0.05/12,120,-100)');
    expect(r).toBeCloseTo(15528.23, 0);
  });

  it('PV calculates present value', async () => {
    // 5% annual, 10 years, $100/month
    const r = await evalFormula('=PV(0.05/12,120,-100)');
    expect(r).toBeCloseTo(9428.13, 0);
  });

  it('PMT calculates payment', async () => {
    // $200000 loan, 30yr, 6% annual
    const r = await evalFormula('=PMT(0.06/12,360,200000)');
    expect(r).toBeCloseTo(-1199.10, 0);
  });

  it('NPER calculates number of periods', async () => {
    // At 1% per period, $-100 payment, PV=0, FV=10000
    const r = await evalFormula('=NPER(0.01,-100,0,10000)');
    expect(typeof r).toBe('number');
    expect(r).toBeGreaterThan(0);
  });

  it('RATE finds interest rate', async () => {
    // 360 months, -$1199.10/month, $200000 PV
    const r = await evalFormula('=RATE(360,-1199.10,200000)');
    expect(r).toBeCloseTo(0.005, 2); // ~0.5% per month = 6% annually
  });

  it('NPV calculates net present value', async () => {
    const r = await evalFormula('=NPV(0.1,-10000,3000,4200,6800)');
    expect(r).toBeCloseTo(1188.44, 0);
  });

  it('IRR finds internal rate of return', async () => {
    const r = await evalFormula('=IRR({-70000,12000,15000,18000,21000,26000})');
    expect(r).toBeCloseTo(0.0866, 2);
  });

  it('MIRR calculates modified IRR', async () => {
    const r = await evalFormula('=MIRR({-120000,39000,30000,21000,37000,46000},0.10,0.12)');
    expect(typeof r).toBe('number');
  });
});

describe('Depreciation', () => {
  it('SLN calculates straight-line depreciation', async () => {
    expect(await evalFormula('=SLN(30000,7500,10)')).toBe(2250);
  });

  it('SYD calculates sum-of-years depreciation', async () => {
    // Cost=30000, Salvage=7500, Life=10, Period=1
    const r = await evalFormula('=SYD(30000,7500,10,1)');
    expect(r).toBeCloseTo(4090.91, 0);
  });

  it('DDB calculates double-declining balance', async () => {
    const r = await evalFormula('=DDB(2400,300,10,1)');
    expect(r).toBeCloseTo(480, 0);
  });
});

describe('Payment breakdown', () => {
  it('IPMT returns interest portion', async () => {
    // 5% annual, period 1, 12 months, $10000 pv
    const r = await evalFormula('=IPMT(0.05/12,1,12,10000)');
    expect(typeof r).toBe('number');
    expect(r).toBeCloseTo(41.67, 0);
  });

  it('PPMT returns principal portion', async () => {
    const r = await evalFormula('=PPMT(0.05/12,1,12,10000)');
    expect(typeof r).toBe('number');
  });

  it('IPMT + PPMT = PMT', async () => {
    const ipmt = await evalFormula('=IPMT(0.05/12,1,12,10000)');
    const ppmt = await evalFormula('=PPMT(0.05/12,1,12,10000)');
    const pmt = await evalFormula('=PMT(0.05/12,12,10000)');
    expect(Number(ipmt) + Number(ppmt)).toBeCloseTo(Number(pmt), 2);
  });
});

describe('Interest rate conversion', () => {
  it('EFFECT converts nominal to effective', async () => {
    // 5.25% nominal, compounded quarterly
    const r = await evalFormula('=EFFECT(0.0525,4)');
    expect(r).toBeCloseTo(0.0535, 3);
  });

  it('NOMINAL converts effective to nominal', async () => {
    const r = await evalFormula('=NOMINAL(0.0535,4)');
    expect(r).toBeCloseTo(0.0525, 3);
  });
});

describe('Dollar conversion', () => {
  it('DOLLARDE converts fractional to decimal', async () => {
    // 1.02 in base-16 fraction = 1 + 0.02*10/16 = 1.0125
    const r = await evalFormula('=DOLLARDE(1.02,16)');
    expect(r).toBeCloseTo(1.0125, 4);
  });

  it('DOLLARFR converts decimal to fractional', async () => {
    const r = await evalFormula('=DOLLARFR(1.125,16)');
    expect(r).toBeCloseTo(1.2, 4);
  });
});

// ============================================================
// Date/Time
// ============================================================

describe('International work day functions', () => {
  it('NETWORKDAYS.INTL counts work days', async () => {
    // DATE(2024,1,1) to DATE(2024,1,31)
    const r = await evalFormula('=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31))');
    expect(typeof r).toBe('number');
    expect(r).toBeGreaterThan(20); // January 2024 has 23 work days
    expect(r).toBeLessThan(25);
  });

  it('WORKDAY.INTL calculates work date', async () => {
    // 10 work days from DATE(2024,1,1) — should skip weekends
    const r = await evalFormula('=WORKDAY.INTL(DATE(2024,1,1),10)');
    expect(typeof r).toBe('number');
    expect(r).toBeGreaterThan(0);
  });
});

// ============================================================
// Information
// ============================================================

describe('Information functions', () => {
  it('INFO returns a string', async () => {
    const r = await evalFormula('=INFO("system")');
    expect(typeof r).toBe('string');
  });

  it('CELL returns empty string stub', async () => {
    const r = await evalFormula('=CELL()');
    expect(r).toBe('');
  });
});

// ============================================================
// CONFIDENCE.NORM
// ============================================================

describe('CONFIDENCE.NORM', () => {
  it('returns confidence interval width', async () => {
    // 95% confidence, stddev=2.5, n=50
    const r = await evalFormula('=CONFIDENCE.NORM(0.05,2.5,50)');
    expect(typeof r).toBe('number');
    expect(r).toBeCloseTo(0.6929, 1);
  });
});
