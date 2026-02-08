// Date/time function tests (DATE, YEAR, MONTH, DAY, TODAY)
// Run with: node --experimental-vm-modules tests/formula-datetime-tests.js

import { createFormulaEngine } from '../src/js/formula/engine/index.js';

const engine = createFormulaEngine({});

async function test(name, formula, expected) {
  const result = await engine.evaluateFormula(formula, { resolveCell: async () => 0 });
  const pass = JSON.stringify(result) === JSON.stringify(expected);
  console.log(pass ? '✓' : '✗', name, '|', formula, '=', result, pass ? '' : `(expected ${expected})`);
  return pass;
}

async function runTests() {
  console.log('--- Date/Time Function Tests ---\n');
  let passed = 0;
  let total = 0;

  // Excel 1900 date system baseline checks
  total++; if (await test('DATE 1900-01-01', '=DATE(1900,1,1)', 1)) passed++;
  total++; if (await test('DATE 1900-02-28', '=DATE(1900,2,28)', 59)) passed++;
  total++; if (await test('DATE 1900-02-29 (bug)', '=DATE(1900,2,29)', 60)) passed++;
  total++; if (await test('DATE 1900-03-01', '=DATE(1900,3,1)', 61)) passed++;

  // Extractors
  total++; if (await test('YEAR from serial', '=YEAR(61)', 1900)) passed++;
  total++; if (await test('MONTH from serial', '=MONTH(61)', 3)) passed++;
  total++; if (await test('DAY from serial', '=DAY(61)', 1)) passed++;
  total++; if (await test('YEAR from 1900-02-29', '=YEAR(60)', 1900)) passed++;
  total++; if (await test('MONTH from 1900-02-29', '=MONTH(60)', 2)) passed++;
  total++; if (await test('DAY from 1900-02-29', '=DAY(60)', 29)) passed++;

  // Nested DATE extraction
  total++; if (await test('YEAR(DATE)', '=YEAR(DATE(2024,1,15))', 2024)) passed++;
  total++; if (await test('MONTH(DATE)', '=MONTH(DATE(2024,12,5))', 12)) passed++;
  total++; if (await test('DAY(DATE)', '=DAY(DATE(2024,12,5))', 5)) passed++;

  // TODAY based on local date
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth() + 1;
  const d = now.getDate();
  total++; if (await test('YEAR(TODAY)', '=YEAR(TODAY())', y)) passed++;
  total++; if (await test('MONTH(TODAY)', '=MONTH(TODAY())', m)) passed++;
  total++; if (await test('DAY(TODAY)', '=DAY(TODAY())', d)) passed++;

  console.log(`\n--- Results: ${passed}/${total} passed ---`);
  if (passed < total) process.exitCode = 1;
}

runTests().catch(err => {
  console.error('Test error:', err);
  process.exitCode = 1;
});
