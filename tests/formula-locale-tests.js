// Localized formula tests (French)
// Run with: node --experimental-vm-modules tests/formula-locale-tests.js

import { createFormulaEngine } from '../src/js/formula/engine/index.js';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const frMap = JSON.parse(readFileSync(join(__dirname, '../src/formula/locales/functions.fr.json'), 'utf-8'));
const deMap = JSON.parse(readFileSync(join(__dirname, '../src/formula/locales/functions.de.json'), 'utf-8'));

async function testLocale(name, localeMap, formula, expected) {
  const engine = createFormulaEngine({ localeMap });
  const result = await engine.evaluateFormula(formula, { resolveCell: async () => 0 });
  const pass = JSON.stringify(result) === JSON.stringify(expected);
  console.log(pass ? '✓' : '✗', name, '|', formula, '=', result, pass ? '' : `(expected ${expected})`);
  return pass;
}

async function runTests() {
  console.log('--- Localized Formula Tests ---\n');
  let passed = 0;
  let total = 0;

  // French
  total++; if (await testLocale('FR SOMME', frMap, '=SOMME(1,2,3)', 6)) passed++;
  total++; if (await testLocale('FR MOYENNE', frMap, '=MOYENNE(10,20,30)', 20)) passed++;
  total++; if (await testLocale('FR SI', frMap, '=SI(1>0,"oui","non")', 'oui')) passed++;
  total++; if (await testLocale('FR ET', frMap, '=ET(1,1)', true)) passed++;
  total++; if (await testLocale('FR OU', frMap, '=OU(0,1)', true)) passed++;
  total++; if (await testLocale('FR NON', frMap, '=NON(0)', true)) passed++;
  total++; if (await testLocale('FR CONCATENER', frMap, '=CONCATENER("a","b")', 'ab')) passed++;
  total++; if (await testLocale('FR GAUCHE', frMap, '=GAUCHE("Bonjour",3)', 'Bon')) passed++;
  total++; if (await testLocale('FR ARRONDI', frMap, '=ARRONDI(3.567,2)', 3.57)) passed++;

  // German
  total++; if (await testLocale('DE SUMME', deMap, '=SUMME(1,2,3)', 6)) passed++;
  total++; if (await testLocale('DE MITTELWERT', deMap, '=MITTELWERT(10,20,30)', 20)) passed++;
  total++; if (await testLocale('DE WENN', deMap, '=WENN(1>0,"ja","nein")', 'ja')) passed++;
  total++; if (await testLocale('DE UND', deMap, '=UND(1,1)', true)) passed++;
  total++; if (await testLocale('DE ODER', deMap, '=ODER(0,1)', true)) passed++;
  total++; if (await testLocale('DE RUNDEN', deMap, '=RUNDEN(3.567,2)', 3.57)) passed++;

  console.log(`\n--- Results: ${passed}/${total} passed ---`);
  return passed === total;
}

runTests().then((ok) => {
  if (typeof process !== 'undefined') process.exit(ok ? 0 : 1);
});
