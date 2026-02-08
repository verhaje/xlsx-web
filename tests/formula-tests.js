// Formula engine tests
// Run in browser console or with: node --experimental-vm-modules tests/formula-tests.js

import { createFormulaEngine } from '../src/js/formula/engine/index.js';

const engine = createFormulaEngine({});

// Mock cell values
const cells = {
  A1: 10,
  A2: 20,
  A3: 30,
  B1: 5,
  B2: 15,
  B3: 25,
  C1: 'Hello',
  C2: 'World',
  D1: '',
  G8: 'test',
};

const resolveCell = async (ref) => cells[ref] ?? '';

async function test(name, formula, expected) {
  const result = await engine.evaluateFormula(formula, { resolveCell });
  const pass = JSON.stringify(result) === JSON.stringify(expected);
  console.log(pass ? '✓' : '✗', name, '|', formula, '=', result, pass ? '' : `(expected ${expected})`);
  return pass;
}

async function runTests() {
  console.log('--- Formula Engine Tests ---\n');
  let passed = 0;
  let total = 0;

  // Arithmetic
  total++; if (await test('Addition', '=1+2', 3)) passed++;
  total++; if (await test('Subtraction', '=10-3', 7)) passed++;
  total++; if (await test('Multiplication', '=4*5', 20)) passed++;
  total++; if (await test('Division', '=20/4', 5)) passed++;
  total++; if (await test('Division by zero', '=1/0', '#DIV/0!')) passed++;
  total++; if (await test('Power', '=2^3', 8)) passed++;
  total++; if (await test('Parentheses', '=(1+2)*3', 9)) passed++;
  total++; if (await test('Unary minus', '=-5', -5)) passed++;
  total++; if (await test('Complex expr', '=1+2*3-4/2', 5)) passed++;

  // Cell references
  total++; if (await test('Cell ref', '=A1', 10)) passed++;
  total++; if (await test('Cell math', '=A1+A2', 30)) passed++;
  total++; if (await test('Cell range sum', '=SUM(A1:A3)', 60)) passed++;

  // Functions
  total++; if (await test('SUM', '=SUM(1,2,3)', 6)) passed++;
  total++; if (await test('AVERAGE', '=AVERAGE(10,20,30)', 20)) passed++;
  total++; if (await test('MIN', '=MIN(5,10,3)', 3)) passed++;
  total++; if (await test('MAX', '=MAX(5,10,3)', 10)) passed++;
  total++; if (await test('COUNT', '=COUNT(1,2,3)', 3)) passed++;
  total++; if (await test('ABS', '=ABS(-5)', 5)) passed++;
  total++; if (await test('ROUND', '=ROUND(3.567,2)', 3.57)) passed++;
  total++; if (await test('SQRT', '=SQRT(16)', 4)) passed++;
  total++; if (await test('LN 1', '=LN(1)', 0)) passed++;
  total++; if (await test('LN rounded', '=ROUND(LN(10),5)', 2.30259)) passed++;
  total++; if (await test('LN text numeric', '=ROUND(LN("2"),5)', 0.69315)) passed++;
  total++; if (await test('LN zero', '=LN(0)', '#NUM!')) passed++;
  total++; if (await test('LN negative', '=LN(-1)', '#NUM!')) passed++;
  total++; if (await test('LN invalid', '=LN("abc")', '#VALUE!')) passed++;
  total++; if (await test('LN invalid in ROUND', '=ROUND(LN("abc"),0)', '#VALUE!')) passed++;
  total++; if (await test('LN invalid in expression', '=LN("abc")+6', '#VALUE!')) passed++;
  total++; if (await test('POWER', '=POWER(2,3)', 8)) passed++;
  total++; if (await test('MOD', '=MOD(10,3)', 1)) passed++;

  // COUNTIFS tests
  total++; if (await test('COUNTIFS basic', '=COUNTIFS(A1:A3,">15",B1:B3,">10")', 2)) passed++;
  total++; if (await test('COUNTIFS wildcard', '=COUNTIFS(C1:C2,"H*")', 1)) passed++;
  total++; if (await test('COUNTIFS blank', '=COUNTIFS(D1:D1,"")', 1)) passed++;
  total++; if (await test('COUNTIFS cell criteria', '=COUNTIFS(A1:A3,A1)', 1)) passed++;
  total++; if (await test('COUNTIFS large range', '=COUNTIFS(A1:A100,10)', 1)) passed++;
  total++; if (await test('COUNTIFS no $ ref', '=COUNTIFS(A1:A3,G8)', 0)) passed++;
  total++; if (await test('COUNTIFS with $ ref', '=COUNTIFS(A1:A3,$G8)', 0)) passed++;

  // Logical
  total++; if (await test('IF true', '=IF(1>0,"yes","no")', 'yes')) passed++;
  total++; if (await test('IF false', '=IF(1<0,"yes","no")', 'no')) passed++;
  total++; if (await test('AND true', '=AND(1,1)', true)) passed++;
  total++; if (await test('AND false', '=AND(1,0)', false)) passed++;
  total++; if (await test('OR true', '=OR(0,1)', true)) passed++;
  total++; if (await test('OR false', '=OR(0,0)', false)) passed++;
  total++; if (await test('NOT', '=NOT(0)', true)) passed++;
  total++; if (await test('TRUE', '=TRUE()', true)) passed++;
  total++; if (await test('FALSE', '=FALSE()', false)) passed++;
  total++; if (await test('IFERROR', '=IFERROR(1/0,"error")', 'error')) passed++;

  // Text
  total++; if (await test('CONCAT', '=CONCAT("a","b","c")', 'abc')) passed++;
  total++; if (await test('LEFT', '=LEFT("Hello",2)', 'He')) passed++;
  total++; if (await test('RIGHT', '=RIGHT("Hello",2)', 'lo')) passed++;
  total++; if (await test('MID', '=MID("Hello",2,3)', 'ell')) passed++;
  total++; if (await test('LEN', '=LEN("Hello")', 5)) passed++;
  total++; if (await test('LOWER', '=LOWER("HELLO")', 'hello')) passed++;
  total++; if (await test('UPPER', '=UPPER("hello")', 'HELLO')) passed++;
  total++; if (await test('TRIM', '=TRIM("  hi  ")', 'hi')) passed++;
  total++; if (await test('Concatenation &', '="a"&"b"', 'ab')) passed++;

  // Comparison
  total++; if (await test('Greater than', '=5>3', true)) passed++;
  total++; if (await test('Less than', '=3<5', true)) passed++;
  total++; if (await test('Equal', '=5=5', true)) passed++;
  total++; if (await test('Not equal', '=5<>3', true)) passed++;
  total++; if (await test('GTE', '=5>=5', true)) passed++;
  total++; if (await test('LTE', '=3<=5', true)) passed++;

  // Error handling
  total++; if (await test('ISBLANK true', '=ISBLANK(D1)', true)) passed++;
  total++; if (await test('ISBLANK false', '=ISBLANK(A1)', false)) passed++;
  total++; if (await test('ISERROR', '=ISERROR(1/0)', true)) passed++;
  total++; if (await test('ISNUMBER number', '=ISNUMBER(123)', true)) passed++;
  total++; if (await test('ISNUMBER text', '=ISNUMBER("123")', false)) passed++;
  total++; if (await test('ISTEXT text', '=ISTEXT("abc")', true)) passed++;
  total++; if (await test('ISTEXT number', '=ISTEXT(123)', false)) passed++;

  // VALUE, SUBSTITUTE, FIND
  total++; if (await test('VALUE number', '=VALUE("123.45")', 123.45)) passed++;
  total++; if (await test('VALUE invalid', '=VALUE("abc")', '#VALUE!')) passed++;
  total++; if (await test('SUBSTITUTE all', '=SUBSTITUTE("Hello World","o","0")', 'Hell0 W0rld')) passed++;
  total++; if (await test('SUBSTITUTE nth', '=SUBSTITUTE("ababa","a","x",2)', 'abxba')) passed++;
  total++; if (await test('FIND found', '=FIND("lo","Hello")', 4)) passed++;
  total++; if (await test('FIND not found', '=FIND("z","Hello")', '#VALUE!')) passed++;
  total++; if (await test('FIND start', '=FIND("l","Hello",4)', 4)) passed++;

  // Statistical
  total++; if (await test('MEDIAN odd', '=MEDIAN(1,3,2)', 2)) passed++;
  total++; if (await test('MEDIAN even', '=MEDIAN(1,2,3,4)', 2.5)) passed++;
  total++; if (await test('MEDIAN range', '=MEDIAN(A1:A3)', 20)) passed++;
  total++; if (await test('STDEV sample', '=STDEV(1,2,3)', 1)) passed++;
  total++; if (await test('STDEV single', '=STDEV(1)', '#DIV/0!')) passed++;

  console.log(`\n--- Results: ${passed}/${total} passed ---`);
  return passed === total;
}

runTests().then((ok) => {
  if (typeof process !== 'undefined') process.exit(ok ? 0 : 1);
});
