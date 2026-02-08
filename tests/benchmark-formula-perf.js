// Formula engine performance benchmark
// Run with: node --experimental-vm-modules tests/benchmark-formula-perf.js

import { createFormulaEngine } from '../src/js/formula/engine/index.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function hrMs(start) {
  const [s, ns] = process.hrtime(start);
  return s * 1e3 + ns / 1e6;
}

function fmt(ms) {
  return ms < 1 ? `${(ms * 1000).toFixed(0)} µs` : `${ms.toFixed(2)} ms`;
}

// ---------------------------------------------------------------------------
// Data setup — 1 000 cells (A1:J100)
// ---------------------------------------------------------------------------
const ROWS = 100;
const COLS = 10;
const cells = {};
for (let r = 1; r <= ROWS; r++) {
  for (let c = 1; c <= COLS; c++) {
    let col = '';
    let cc = c;
    while (cc > 0) { const rem = (cc - 1) % 26; col = String.fromCharCode(65 + rem) + col; cc = Math.floor((cc - 1) / 26); }
    cells[`${col}${r}`] = r * c; // deterministic value
  }
}

let resolveCalls = 0;
const resolveCell = async (ref) => { resolveCalls++; return cells[ref] ?? ''; };

// Batch resolver — returns all values in one go
const resolveCellsBatch = async (refs) => {
  resolveCalls += refs.length;
  return refs.map(ref => cells[ref] ?? '');
};

// ---------------------------------------------------------------------------
// Benchmark runner
// ---------------------------------------------------------------------------
async function bench(label, iterations, fn) {
  // warm up
  for (let i = 0; i < Math.min(10, iterations); i++) await fn();
  resolveCalls = 0;
  const start = process.hrtime();
  for (let i = 0; i < iterations; i++) await fn();
  const elapsed = hrMs(start);
  const perIter = elapsed / iterations;
  console.log(`  ${label.padEnd(42)} ${fmt(elapsed).padStart(10)}  (${fmt(perIter)}/iter, resolves: ${resolveCalls})`);
  return { elapsed, perIter, resolveCalls };
}

// ---------------------------------------------------------------------------
// Benchmarks
// ---------------------------------------------------------------------------
async function run() {
  console.log('=== Formula Engine Performance Benchmark ===\n');

  const engine = createFormulaEngine();
  const ctx = { resolveCell, resolveCellsBatch };
  const ctxNoBatch = { resolveCell };
  const N = 500;

  console.log(`--- Simple arithmetic (${N} iterations) ---`);
  await bench('constant expression  =1+2*3', N, () => engine.evaluateFormula('=1+2*3', ctx));
  await bench('cell ref math        =A1+B1', N, () => engine.evaluateFormula('=A1+B1', ctx));
  await bench('nested IF            =IF(A1>50,"y","n")', N, () => engine.evaluateFormula('=IF(A1>50,"y","n")', ctx));

  console.log(`\n--- Aggregates over 100-cell range (${N} iterations) ---`);
  await bench('SUM(A1:A100)         batch resolve', N, () => engine.evaluateFormula('=SUM(A1:A100)', ctx));
  await bench('SUM(A1:A100)         no batch', N, () => engine.evaluateFormula('=SUM(A1:A100)', ctxNoBatch));
  await bench('AVERAGE(A1:A100)', N, () => engine.evaluateFormula('=AVERAGE(A1:A100)', ctx));
  await bench('MIN(A1:A100)', N, () => engine.evaluateFormula('=MIN(A1:A100)', ctx));
  await bench('MAX(A1:A100)', N, () => engine.evaluateFormula('=MAX(A1:A100)', ctx));
  await bench('COUNT(A1:A100)', N, () => engine.evaluateFormula('=COUNT(A1:A100)', ctx));
  await bench('MEDIAN(A1:A100)', N, () => engine.evaluateFormula('=MEDIAN(A1:A100)', ctx));

  console.log(`\n--- Large range 1000 cells (${N} iterations) ---`);
  await bench('SUM(A1:J100)         batch resolve', N, () => engine.evaluateFormula('=SUM(A1:J100)', ctx));
  await bench('SUM(A1:J100)         no batch', N, () => engine.evaluateFormula('=SUM(A1:J100)', ctxNoBatch));

  console.log(`\n--- Tokenization cache effect (${N} iterations) ---`);
  // First call tokenizes, subsequent calls hit cache
  const formula = '=IF(SUM(A1:A50)>1000,AVERAGE(B1:B50),0)';
  engine.clearCache();
  await bench('complex formula cold+cached mix', N, () => engine.evaluateFormula(formula, ctx));

  console.log(`\n--- Text / lookup functions (${N} iterations) ---`);
  await bench('CONCAT("a","b","c")', N, () => engine.evaluateFormula('=CONCAT("a","b","c")', ctx));
  await bench('LEN("Hello World")', N, () => engine.evaluateFormula('=LEN("Hello World")', ctx));

  console.log(`\n--- Date functions (${N} iterations) ---`);
  await bench('DATE(2026,2,6)', N, () => engine.evaluateFormula('=DATE(2026,2,6)', ctx));
  await bench('YEAR(TODAY())', N, () => engine.evaluateFormula('=YEAR(TODAY())', ctx));

  console.log(`\n--- Resolve cache dedup test ---`);
  // Formula referencing same cell many times — resolve cache should dedup
  resolveCalls = 0;
  const dedupFormula = '=A1+A1+A1+A1+A1';
  await bench('A1+A1+A1+A1+A1 (dedup)', N, () => engine.evaluateFormula(dedupFormula, ctx));

  console.log('\n=== Benchmark complete ===');
}

run().catch(err => {
  console.error('Benchmark error:', err);
  process.exitCode = 1;
});
