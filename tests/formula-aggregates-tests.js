// Conditional aggregate function tests (SUMIF, SUMIFS, AVERAGEIF, AVERAGEIFS)
// Run with: node --experimental-vm-modules tests/formula-aggregates-tests.js

import { createFormulaEngine } from '../src/js/formula/engine/index.js';

const engine = createFormulaEngine({});

// Mock cell values - Sales data table
// A        B        C        D
// Product  Region   Sales    Quantity
// Apple    North    100      10
// Banana   South    200      20
// Apple    South    150      15
// Cherry   North    300      30
// Banana   North    250      25
// Apple    North    175      17

const cells = {
  // Headers
  A1: 'Product', B1: 'Region', C1: 'Sales', D1: 'Quantity',
  // Data rows
  A2: 'Apple',  B2: 'North', C2: 100, D2: 10,
  A3: 'Banana', B3: 'South', C3: 200, D3: 20,
  A4: 'Apple',  B4: 'South', C4: 150, D4: 15,
  A5: 'Cherry', B5: 'North', C5: 300, D5: 30,
  A6: 'Banana', B6: 'North', C6: 250, D6: 25,
  A7: 'Apple',  B7: 'North', C7: 175, D7: 17,
  
  // Numeric test data (E1:E5)
  E1: 10, E2: 20, E3: 30, E4: 40, E5: 50,
  
  // Criteria values (F1:F5)
  F1: 5, F2: 15, F3: 25, F4: 35, F5: 45,
  
  // Mixed data for blank/text tests (G1:G5)
  G1: '', G2: 'text', G3: 100, G4: '', G5: 200,
  
  // Corresponding values for sum (H1:H5)
  H1: 1, H2: 2, H3: 3, H4: 4, H5: 5,
};

const resolveCell = async (ref) => cells[ref] ?? '';

async function test(name, formula, expected) {
  const result = await engine.evaluateFormula(formula, { resolveCell });
  // Handle floating point comparison
  let pass;
  if (typeof expected === 'number' && typeof result === 'number') {
    pass = Math.abs(result - expected) < 0.0001;
  } else {
    pass = JSON.stringify(result) === JSON.stringify(expected);
  }
  console.log(pass ? '✓' : '✗', name, '|', formula, '=', result, pass ? '' : `(expected ${JSON.stringify(expected)})`);
  return pass;
}

async function runTests() {
  console.log('--- Conditional Aggregate Function Tests ---\n');
  let passed = 0;
  let total = 0;

  // ========== SUMIF Tests ==========
  console.log('\n=== SUMIF Tests ===');
  
  // Basic SUMIF with text criteria
  total++; if (await test('SUMIF text match', '=SUMIF(A2:A7,"Apple",C2:C7)', 425)) passed++; // 100+150+175
  total++; if (await test('SUMIF text match Banana', '=SUMIF(A2:A7,"Banana",C2:C7)', 450)) passed++; // 200+250
  total++; if (await test('SUMIF text match Cherry', '=SUMIF(A2:A7,"Cherry",C2:C7)', 300)) passed++;
  
  // SUMIF with numeric criteria
  total++; if (await test('SUMIF numeric equal', '=SUMIF(C2:C7,100,D2:D7)', 10)) passed++;
  total++; if (await test('SUMIF numeric greater', '=SUMIF(C2:C7,">200",D2:D7)', 55)) passed++; // 30+25
  total++; if (await test('SUMIF numeric less', '=SUMIF(C2:C7,"<150",D2:D7)', 10)) passed++; // only 100
  total++; if (await test('SUMIF numeric gte', '=SUMIF(C2:C7,">=200",D2:D7)', 75)) passed++; // 20+30+25
  total++; if (await test('SUMIF numeric lte', '=SUMIF(C2:C7,"<=150",D2:D7)', 25)) passed++; // 10+15
  total++; if (await test('SUMIF not equal', '=SUMIF(C2:C7,"<>100",D2:D7)', 107)) passed++; // all except 10
  
  // SUMIF with wildcards
  total++; if (await test('SUMIF wildcard *', '=SUMIF(A2:A7,"A*",C2:C7)', 425)) passed++; // Apple matches
  total++; if (await test('SUMIF wildcard ?', '=SUMIF(A2:A7,"?pple",C2:C7)', 425)) passed++; // Apple matches
  total++; if (await test('SUMIF wildcard end', '=SUMIF(A2:A7,"*rry",C2:C7)', 300)) passed++; // Cherry
  
  // SUMIF without sum_range (sum criteria_range itself)
  total++; if (await test('SUMIF no sum_range', '=SUMIF(E1:E5,">25")', 120)) passed++; // 30+40+50
  
  // SUMIF case insensitive
  total++; if (await test('SUMIF case insensitive', '=SUMIF(A2:A7,"APPLE",C2:C7)', 425)) passed++;
  total++; if (await test('SUMIF case insensitive lower', '=SUMIF(A2:A7,"apple",C2:C7)', 425)) passed++;
  
  // SUMIF no matches
  total++; if (await test('SUMIF no match', '=SUMIF(A2:A7,"Orange",C2:C7)', 0)) passed++;
  
  // ========== SUMIFS Tests ==========
  console.log('\n=== SUMIFS Tests ===');
  
  // SUMIFS with multiple criteria
  total++; if (await test('SUMIFS two criteria', '=SUMIFS(C2:C7,A2:A7,"Apple",B2:B7,"North")', 275)) passed++; // 100+175
  total++; if (await test('SUMIFS two criteria 2', '=SUMIFS(C2:C7,A2:A7,"Banana",B2:B7,"North")', 250)) passed++;
  total++; if (await test('SUMIFS two criteria 3', '=SUMIFS(C2:C7,A2:A7,"Apple",B2:B7,"South")', 150)) passed++;
  
  // SUMIFS with numeric criteria
  total++; if (await test('SUMIFS numeric criteria', '=SUMIFS(D2:D7,C2:C7,">100",B2:B7,"North")', 72)) passed++; // 30+25+17
  total++; if (await test('SUMIFS mixed criteria', '=SUMIFS(C2:C7,A2:A7,"Apple",C2:C7,">100")', 325)) passed++; // 150+175
  
  // SUMIFS with wildcards
  total++; if (await test('SUMIFS wildcard', '=SUMIFS(C2:C7,A2:A7,"*an*",B2:B7,"North")', 250)) passed++; // Banana North
  
  // SUMIFS all criteria
  total++; if (await test('SUMIFS three criteria', '=SUMIFS(D2:D7,A2:A7,"Apple",B2:B7,"North",C2:C7,">150")', 17)) passed++;
  
  // ========== AVERAGEIF Tests ==========
  console.log('\n=== AVERAGEIF Tests ===');
  
  // Basic AVERAGEIF with text criteria
  total++; if (await test('AVERAGEIF text match', '=AVERAGEIF(A2:A7,"Apple",C2:C7)', 141.6667)) passed++; // (100+150+175)/3
  total++; if (await test('AVERAGEIF text Banana', '=AVERAGEIF(A2:A7,"Banana",C2:C7)', 225)) passed++; // (200+250)/2
  
  // AVERAGEIF with numeric criteria
  total++; if (await test('AVERAGEIF numeric gt', '=AVERAGEIF(C2:C7,">200",D2:D7)', 27.5)) passed++; // (30+25)/2
  total++; if (await test('AVERAGEIF numeric lte', '=AVERAGEIF(C2:C7,"<=150",D2:D7)', 12.5)) passed++; // (10+15)/2
  
  // AVERAGEIF without average_range
  total++; if (await test('AVERAGEIF no avg_range', '=AVERAGEIF(E1:E5,">25")', 40)) passed++; // (30+40+50)/3
  
  // AVERAGEIF with wildcards
  total++; if (await test('AVERAGEIF wildcard', '=AVERAGEIF(A2:A7,"B*",C2:C7)', 225)) passed++; // Banana avg
  
  // AVERAGEIF no matches (should return #DIV/0!)
  total++; if (await test('AVERAGEIF no match', '=AVERAGEIF(A2:A7,"Orange",C2:C7)', '#DIV/0!')) passed++;
  
  // ========== AVERAGEIFS Tests ==========
  console.log('\n=== AVERAGEIFS Tests ===');
  
  // AVERAGEIFS with multiple criteria
  total++; if (await test('AVERAGEIFS two criteria', '=AVERAGEIFS(C2:C7,A2:A7,"Apple",B2:B7,"North")', 137.5)) passed++; // (100+175)/2
  total++; if (await test('AVERAGEIFS two criteria 2', '=AVERAGEIFS(C2:C7,A2:A7,"Banana",B2:B7,"South")', 200)) passed++;
  
  // AVERAGEIFS with numeric criteria
  total++; if (await test('AVERAGEIFS numeric', '=AVERAGEIFS(D2:D7,C2:C7,">100",B2:B7,"North")', 24)) passed++; // (30+25+17)/3
  
  // AVERAGEIFS no matches
  total++; if (await test('AVERAGEIFS no match', '=AVERAGEIFS(C2:C7,A2:A7,"Orange",B2:B7,"North")', '#DIV/0!')) passed++;
  
  // ========== Edge Cases ==========
  console.log('\n=== Edge Cases ===');
  
  // Empty/blank handling
  total++; if (await test('SUMIF blank criteria', '=SUMIF(G1:G5,"",H1:H5)', 5)) passed++; // H1+H4 = 1+4
  total++; if (await test('SUMIF text in numeric check', '=SUMIF(G1:G5,">50",H1:H5)', 8)) passed++; // H3+H5 = 3+5
  
  // Combined with other functions
  total++; if (await test('IF with SUMIF', '=IF(SUMIF(A2:A7,"Apple",C2:C7)>400,"High","Low")', 'High')) passed++;
  total++; if (await test('IFERROR with AVERAGEIF', '=IFERROR(AVERAGEIF(A2:A7,"Orange",C2:C7),"None")', 'None')) passed++;
  
  // Single cell ranges
  total++; if (await test('SUMIF single cell', '=SUMIF(A2:A2,"Apple",C2:C2)', 100)) passed++;
  total++; if (await test('AVERAGEIF single cell', '=AVERAGEIF(A2:A2,"Apple",C2:C2)', 100)) passed++;

  console.log(`\n--- Results: ${passed}/${total} tests passed ---`);
  if (passed < total) process.exitCode = 1;
}

runTests().catch(err => {
  console.error('Test error:', err);
  process.exitCode = 1;
});
