// Cross-sheet reference tests
// Run with: node --experimental-vm-modules tests/formula-crosssheet-tests.js

import { createFormulaEngine } from '../src/js/formula/engine/index.js';

async function test(name, formula, expected, cellData = {}) {
  const engine = createFormulaEngine();
  const resolveCell = async (ref) => {
    if (cellData[ref] !== undefined) return cellData[ref];
    return 0;
  };
  const result = await engine.evaluateFormula(formula, { resolveCell });
  const pass = JSON.stringify(result) === JSON.stringify(expected);
  console.log(pass ? '✓' : '✗', name, '|', formula, '=', result, pass ? '' : `(expected ${expected})`);
  return pass;
}

async function runTests() {
  console.log('--- Cross-Sheet Reference Tests ---\n');
  let passed = 0;
  let total = 0;

  // Simple cross-sheet cell reference
  total++; if (await test('Sheet ref', '=Sheet2!A1', 100, { 'Sheet2!A1': 100 })) passed++;
  
  // Quoted sheet name with space
  total++; if (await test('Quoted sheet', "='My Sheet'!B2", 200, { 'My Sheet!B2': 200 })) passed++;
  
  // Cross-sheet math
  total++; if (await test('Sheet math', '=Sheet2!A1 + Sheet2!B1', 150, { 'Sheet2!A1': 100, 'Sheet2!B1': 50 })) passed++;
  
  // Mixed same-sheet and cross-sheet
  total++; if (await test('Mixed refs', '=A1 + Sheet2!A1', 130, { 'A1': 30, 'Sheet2!A1': 100 })) passed++;
  
  // Cross-sheet range in SUM
  total++; if (await test('Sheet range SUM', '=SUM(Sheet2!A1:A3)', 60, { 
    'Sheet2!A1': 10, 
    'Sheet2!A2': 20, 
    'Sheet2!A3': 30 
  })) passed++;
  
  // Quoted sheet with range
  total++; if (await test('Quoted sheet range', "=SUM('Data Sheet'!B1:B2)", 75, { 
    'Data Sheet!B1': 25, 
    'Data Sheet!B2': 50 
  })) passed++;
  
  // Multiple sheet references
  total++; if (await test('Multiple sheets', '=Sheet1!A1 + Sheet2!A1 + Sheet3!A1', 60, { 
    'Sheet1!A1': 10, 
    'Sheet2!A1': 20, 
    'Sheet3!A1': 30 
  })) passed++;
  
  // Cross-sheet with function
  total++; if (await test('Sheet IF', '=IF(Sheet2!A1>50,"High","Low")', 'High', { 'Sheet2!A1': 100 })) passed++;
  
  // Sheet name that looks like a cell
  total++; if (await test('Sheet like cell', '=A1B2!C3', 999, { 'A1B2!C3': 999 })) passed++;
  
  // Cross-sheet string concat
  total++; if (await test('Sheet concat', '=Sheet1!A1&" "&Sheet2!A1', 'Hello World', { 
    'Sheet1!A1': 'Hello', 
    'Sheet2!A1': 'World' 
  })) passed++;
  
  // Complex quoted sheet name
  total++; if (await test('Complex sheet name', "='Sales 2024 (Q1)'!A1", 5000, { 'Sales 2024 (Q1)!A1': 5000 })) passed++;
  
  // Cross-sheet AVERAGE
  total++; if (await test('Sheet AVERAGE', '=AVERAGE(Data!A1:A4)', 25, { 
    'Data!A1': 10, 
    'Data!A2': 20, 
    'Data!A3': 30, 
    'Data!A4': 40 
  })) passed++;
  
  // Cross-sheet comparison
  total++; if (await test('Sheet comparison', '=Sheet1!A1 > Sheet2!A1', true, { 
    'Sheet1!A1': 100, 
    'Sheet2!A1': 50 
  })) passed++;
  
  // Nested function with cross-sheet
  total++; if (await test('Nested cross-sheet', '=IF(SUM(Data!A1:A2)>100,"High","Low")', 'High', { 
    'Data!A1': 60, 
    'Data!A2': 50 
  })) passed++;

  // Absolute references with $
  total++; if (await test('Sheet $col', '=Sheet1!$A1', 42, { 'Sheet1!A1': 42 })) passed++;
  total++; if (await test('Sheet $col$row', '=Sheet1!$A$1', 77, { 'Sheet1!A1': 77 })) passed++;
  total++; if (await test('Sheet $row', '=Sheet1!A$1', 88, { 'Sheet1!A1': 88 })) passed++;
  total++; if (await test('Quoted $ ref', "='test test'!$B$2", 123, { 'test test!B2': 123 })) passed++;
  total++; if (await test('Mixed $ math', '=Sheet1!$A$1 + Sheet2!B$2', 150, { 'Sheet1!A1': 100, 'Sheet2!B2': 50 })) passed++;
  total++; if (await test('SUM with $ range', '=SUM(Sheet1!$A$1:$A$3)', 60, { 
    'Sheet1!A1': 10, 
    'Sheet1!A2': 20, 
    'Sheet1!A3': 30 
  })) passed++;

  console.log(`\n--- Results: ${passed}/${total} passed ---`);
  return passed === total;
}

runTests().then((ok) => {
  if (typeof process !== 'undefined') process.exit(ok ? 0 : 1);
});
