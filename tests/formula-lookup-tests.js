// Lookup function tests (VLOOKUP, HLOOKUP, INDEX, MATCH)
// Run with: node --experimental-vm-modules tests/formula-lookup-tests.js

import { createFormulaEngine } from '../src/js/formula/engine/index.js';

const engine = createFormulaEngine({});

// Mock cell values - vertical lookup table (A1:C5)
// A      B        C
// 1      Apple    10
// 2      Banana   20
// 3      Cherry   30
// 4      Date     40
// 5      Elderberry 50

// Horizontal lookup table (E1:I3)
// E1:I1 = 1, 2, 3, 4, 5 (headers)
// E2:I2 = Apple, Banana, Cherry, Date, Elderberry
// E3:I3 = 10, 20, 30, 40, 50

const cells = {
  // Vertical table A1:C5
  A1: 1, B1: 'Apple', C1: 10,
  A2: 2, B2: 'Banana', C2: 20,
  A3: 3, B3: 'Cherry', C3: 30,
  A4: 4, B4: 'Date', C4: 40,
  A5: 5, B5: 'Elderberry', C5: 50,
  
  // Horizontal table E1:I3
  E1: 1, F1: 2, G1: 3, H1: 4, I1: 5,
  E2: 'Apple', F2: 'Banana', G2: 'Cherry', H2: 'Date', I2: 'Elderberry',
  E3: 10, F3: 20, G3: 30, H3: 40, I3: 50,
  
  // Simple 1D array for MATCH (J1:J5)
  J1: 10, J2: 20, J3: 30, J4: 40, J5: 50,
  
  // Horizontal 1D array for MATCH (K1:O1)
  K1: 'alpha', L1: 'beta', M1: 'gamma', N1: 'delta', O1: 'epsilon',
  
  // Unsorted array for exact match (P1:P5)
  P1: 30, P2: 10, P3: 50, P4: 20, P5: 40,
  
  // Lookup values
  Z1: 3,      // Numeric lookup value
  Z2: 'Cherry', // String lookup value
  Z3: 99,     // Non-existent value
};

const resolveCell = async (ref) => cells[ref] ?? '';

async function test(name, formula, expected) {
  const result = await engine.evaluateFormula(formula, { resolveCell });
  const pass = JSON.stringify(result) === JSON.stringify(expected);
  console.log(pass ? '✓' : '✗', name, '|', formula, '=', result, pass ? '' : `(expected ${JSON.stringify(expected)})`);
  return pass;
}

async function runTests() {
  console.log('--- Lookup Function Tests ---\n');
  let passed = 0;
  let total = 0;

  // ========== MATCH Tests ==========
  console.log('\n=== MATCH Tests ===');
  
  // Exact match (match_type = 0)
  total++; if (await test('MATCH exact number', '=MATCH(30,J1:J5,0)', 3)) passed++;
  total++; if (await test('MATCH exact first', '=MATCH(10,J1:J5,0)', 1)) passed++;
  total++; if (await test('MATCH exact last', '=MATCH(50,J1:J5,0)', 5)) passed++;
  total++; if (await test('MATCH exact not found', '=MATCH(99,J1:J5,0)', '#N/A')) passed++;
  total++; if (await test('MATCH exact string', '=MATCH("gamma",K1:O1,0)', 3)) passed++;
  total++; if (await test('MATCH exact string case insensitive', '=MATCH("BETA",K1:O1,0)', 2)) passed++;
  total++; if (await test('MATCH exact string not found', '=MATCH("omega",K1:O1,0)', '#N/A')) passed++;
  
  // Approximate match (match_type = 1, default - ascending)
  total++; if (await test('MATCH approx ascending', '=MATCH(25,J1:J5,1)', 2)) passed++;
  total++; if (await test('MATCH approx exact value', '=MATCH(30,J1:J5,1)', 3)) passed++;
  total++; if (await test('MATCH approx below min', '=MATCH(5,J1:J5,1)', '#N/A')) passed++;
  total++; if (await test('MATCH approx above max', '=MATCH(100,J1:J5,1)', 5)) passed++;
  total++; if (await test('MATCH default match_type', '=MATCH(35,J1:J5)', 3)) passed++;
  
  // Approximate match (match_type = -1, descending)
  // For descending, we need a descending array - using P1:P5 which is unsorted
  // Actually -1 looks for smallest value >= lookup, assuming descending order
  total++; if (await test('MATCH approx descending', '=MATCH(25,J1:J5,-1)', 3)) passed++;
  
  // ========== INDEX Tests ==========
  console.log('\n=== INDEX Tests ===');
  
  // 2D array INDEX
  total++; if (await test('INDEX 2D row 1 col 1', '=INDEX(A1:C5,1,1)', 1)) passed++;
  total++; if (await test('INDEX 2D row 1 col 2', '=INDEX(A1:C5,1,2)', 'Apple')) passed++;
  total++; if (await test('INDEX 2D row 3 col 3', '=INDEX(A1:C5,3,3)', 30)) passed++;
  total++; if (await test('INDEX 2D row 5 col 2', '=INDEX(A1:C5,5,2)', 'Elderberry')) passed++;
  total++; if (await test('INDEX 2D last cell', '=INDEX(A1:C5,5,3)', 50)) passed++;
  
  // 1D array INDEX
  total++; if (await test('INDEX 1D row 3', '=INDEX(J1:J5,3,1)', 30)) passed++;
  total++; if (await test('INDEX 1D row 1', '=INDEX(J1:J5,1,1)', 10)) passed++;
  
  // Error cases
  total++; if (await test('INDEX row out of bounds', '=INDEX(A1:C5,10,1)', '#REF!')) passed++;
  total++; if (await test('INDEX col out of bounds', '=INDEX(A1:C5,1,10)', '#REF!')) passed++;
  total++; if (await test('INDEX negative row', '=INDEX(A1:C5,-1,1)', '#VALUE!')) passed++;
  
  // ========== VLOOKUP Tests ==========
  console.log('\n=== VLOOKUP Tests ===');
  
  // Exact match (range_lookup = FALSE)
  total++; if (await test('VLOOKUP exact number col 2', '=VLOOKUP(2,A1:C5,2,FALSE)', 'Banana')) passed++;
  total++; if (await test('VLOOKUP exact number col 3', '=VLOOKUP(3,A1:C5,3,FALSE)', 30)) passed++;
  total++; if (await test('VLOOKUP exact first row', '=VLOOKUP(1,A1:C5,2,FALSE)', 'Apple')) passed++;
  total++; if (await test('VLOOKUP exact last row', '=VLOOKUP(5,A1:C5,2,FALSE)', 'Elderberry')) passed++;
  total++; if (await test('VLOOKUP exact not found', '=VLOOKUP(99,A1:C5,2,FALSE)', '#N/A')) passed++;
  total++; if (await test('VLOOKUP exact using 0', '=VLOOKUP(4,A1:C5,3,0)', 40)) passed++;
  
  // Approximate match (range_lookup = TRUE or omitted)
  total++; if (await test('VLOOKUP approx default', '=VLOOKUP(2.5,A1:C5,2)', 'Banana')) passed++;
  total++; if (await test('VLOOKUP approx TRUE', '=VLOOKUP(3.9,A1:C5,2,TRUE)', 'Cherry')) passed++;
  total++; if (await test('VLOOKUP approx exact value', '=VLOOKUP(4,A1:C5,3,TRUE)', 40)) passed++;
  total++; if (await test('VLOOKUP approx above max', '=VLOOKUP(100,A1:C5,2,TRUE)', 'Elderberry')) passed++;
  total++; if (await test('VLOOKUP approx below min', '=VLOOKUP(0,A1:C5,2,TRUE)', '#N/A')) passed++;
  
  // Error cases
  total++; if (await test('VLOOKUP col index 0', '=VLOOKUP(2,A1:C5,0,FALSE)', '#VALUE!')) passed++;
  total++; if (await test('VLOOKUP col index too large', '=VLOOKUP(2,A1:C5,10,FALSE)', '#REF!')) passed++;
  
  // Using cell reference for lookup value
  total++; if (await test('VLOOKUP with cell ref', '=VLOOKUP(Z1,A1:C5,2,FALSE)', 'Cherry')) passed++;
  
  // ========== HLOOKUP Tests ==========
  console.log('\n=== HLOOKUP Tests ===');
  
  // Exact match
  total++; if (await test('HLOOKUP exact number row 2', '=HLOOKUP(2,E1:I3,2,FALSE)', 'Banana')) passed++;
  total++; if (await test('HLOOKUP exact number row 3', '=HLOOKUP(3,E1:I3,3,FALSE)', 30)) passed++;
  total++; if (await test('HLOOKUP exact first col', '=HLOOKUP(1,E1:I3,2,FALSE)', 'Apple')) passed++;
  total++; if (await test('HLOOKUP exact last col', '=HLOOKUP(5,E1:I3,2,FALSE)', 'Elderberry')) passed++;
  total++; if (await test('HLOOKUP exact not found', '=HLOOKUP(99,E1:I3,2,FALSE)', '#N/A')) passed++;
  
  // Approximate match
  total++; if (await test('HLOOKUP approx default', '=HLOOKUP(2.5,E1:I3,2)', 'Banana')) passed++;
  total++; if (await test('HLOOKUP approx above max', '=HLOOKUP(100,E1:I3,3,TRUE)', 50)) passed++;
  
  // Error cases
  total++; if (await test('HLOOKUP row index 0', '=HLOOKUP(2,E1:I3,0,FALSE)', '#VALUE!')) passed++;
  total++; if (await test('HLOOKUP row index too large', '=HLOOKUP(2,E1:I3,10,FALSE)', '#REF!')) passed++;
  
  // ========== Combined/Nested Tests ==========
  console.log('\n=== Combined/Nested Tests ===');
  
  // INDEX-MATCH pattern (two-way lookup)
  total++; if (await test('INDEX-MATCH combo', '=INDEX(C1:C5,MATCH(3,A1:A5,0),1)', 30)) passed++;
  total++; if (await test('INDEX-MATCH horizontal', '=INDEX(E3:I3,1,MATCH(3,E1:I1,0))', 30)) passed++;
  
  // VLOOKUP nested in IF
  total++; if (await test('IF with VLOOKUP', '=IF(VLOOKUP(2,A1:C5,3,FALSE)>15,"big","small")', 'big')) passed++;
  
  // IFERROR with lookup
  total++; if (await test('IFERROR with VLOOKUP miss', '=IFERROR(VLOOKUP(99,A1:C5,2,FALSE),"Not found")', 'Not found')) passed++;
  total++; if (await test('IFERROR with VLOOKUP hit', '=IFERROR(VLOOKUP(2,A1:C5,2,FALSE),"Not found")', 'Banana')) passed++;
  
  // ========== Edge Cases ==========
  console.log('\n=== Edge Cases ===');
  
  // Empty/blank handling
  total++; if (await test('VLOOKUP with empty result', '=VLOOKUP(1,A1:C5,2,FALSE)', 'Apple')) passed++;
  
  // String lookups in VLOOKUP (first column is text)
  // Need a text-keyed table... let's test with what we have in B column
  
  // Single cell as range
  total++; if (await test('INDEX single cell', '=INDEX(A1:A1,1,1)', 1)) passed++;
  total++; if (await test('MATCH in single cell', '=MATCH(1,A1:A1,0)', 1)) passed++;
  
  console.log(`\n--- Results: ${passed}/${total} tests passed ---`);
  if (passed < total) process.exitCode = 1;
}

runTests().catch(err => {
  console.error('Test error:', err);
  process.exitCode = 1;
});
