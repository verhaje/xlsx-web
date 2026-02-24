import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';

const __dirname = dirname(fileURLToPath(import.meta.url));
const frMap: Record<string, string> = JSON.parse(
  readFileSync(join(__dirname, '../src/formula/locales/functions.fr.json'), 'utf-8')
);
const deMap: Record<string, string> = JSON.parse(
  readFileSync(join(__dirname, '../src/formula/locales/functions.de.json'), 'utf-8')
);

async function evalLocale(localeMap: Record<string, string>, formula: string) {
  const engine = createFormulaEngine({ localeMap });
  return engine.evaluateFormula(formula, { resolveCell: async () => 0 });
}

describe('Localized Formulas', () => {
  describe('French', () => {
    it.each([
      ['SOMME',       '=SOMME(1,2,3)',           6],
      ['MOYENNE',     '=MOYENNE(10,20,30)',       20],
      ['SI',          '=SI(1>0,"oui","non")',     'oui'],
      ['ET',          '=ET(1,1)',                 true],
      ['OU',          '=OU(0,1)',                 true],
      ['NON',         '=NON(0)',                  true],
      ['CONCATENER',  '=CONCATENER("a","b")',     'ab'],
      ['GAUCHE',      '=GAUCHE("Bonjour",3)',     'Bon'],
      ['ARRONDI',     '=ARRONDI(3.567,2)',        3.57],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalLocale(frMap, formula)).toEqual(expected);
    });
  });

  describe('German', () => {
    it.each([
      ['SUMME',       '=SUMME(1,2,3)',            6],
      ['MITTELWERT',  '=MITTELWERT(10,20,30)',    20],
      ['WENN',        '=WENN(1>0,"ja","nein")',   'ja'],
      ['UND',         '=UND(1,1)',                true],
      ['ODER',        '=ODER(0,1)',               true],
      ['RUNDEN',      '=RUNDEN(3.567,2)',         3.57],
    ])('%s', async (_name, formula, expected) => {
      expect(await evalLocale(deMap, formula)).toEqual(expected);
    });
  });
});
