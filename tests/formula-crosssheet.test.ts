import { describe, it, expect } from 'vitest';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';

async function evalWithCells(formula: string, cellData: Record<string, any> = {}) {
  const engine = createFormulaEngine();
  const resolveCell = async (ref: string) => cellData[ref] ?? 0;
  return engine.evaluateFormula(formula, { resolveCell });
}

describe('Cross-Sheet References', () => {
  it.each([
    ['Sheet ref',            '=Sheet2!A1',                        100,   { 'Sheet2!A1': 100 }],
    ['Quoted sheet',         "='My Sheet'!B2",                    200,   { 'My Sheet!B2': 200 }],
    ['Sheet math',           '=Sheet2!A1 + Sheet2!B1',            150,   { 'Sheet2!A1': 100, 'Sheet2!B1': 50 }],
    ['Mixed refs',           '=A1 + Sheet2!A1',                   130,   { A1: 30, 'Sheet2!A1': 100 }],
    ['Sheet range SUM',      '=SUM(Sheet2!A1:A3)',                60,    { 'Sheet2!A1': 10, 'Sheet2!A2': 20, 'Sheet2!A3': 30 }],
    ['Quoted sheet range',   "=SUM('Data Sheet'!B1:B2)",          75,    { 'Data Sheet!B1': 25, 'Data Sheet!B2': 50 }],
    ['Multiple sheets',      '=Sheet1!A1 + Sheet2!A1 + Sheet3!A1', 60,  { 'Sheet1!A1': 10, 'Sheet2!A1': 20, 'Sheet3!A1': 30 }],
    ['Sheet IF',             '=IF(Sheet2!A1>50,"High","Low")',    'High',{ 'Sheet2!A1': 100 }],
    ['Sheet like cell',      '=A1B2!C3',                          999,  { 'A1B2!C3': 999 }],
    ['Sheet concat',         '=Sheet1!A1&" "&Sheet2!A1',          'Hello World', { 'Sheet1!A1': 'Hello', 'Sheet2!A1': 'World' }],
    ['Complex sheet name',   "='Sales 2024 (Q1)'!A1",            5000,  { 'Sales 2024 (Q1)!A1': 5000 }],
    ['Sheet AVERAGE',        '=AVERAGE(Data!A1:A4)',              25,    { 'Data!A1': 10, 'Data!A2': 20, 'Data!A3': 30, 'Data!A4': 40 }],
    ['Sheet comparison',     '=Sheet1!A1 > Sheet2!A1',            true,  { 'Sheet1!A1': 100, 'Sheet2!A1': 50 }],
    ['Nested cross-sheet',   '=IF(SUM(Data!A1:A2)>100,"High","Low")', 'High', { 'Data!A1': 60, 'Data!A2': 50 }],
  ])('%s', async (_name, formula, expected, cellData) => {
    expect(await evalWithCells(formula, cellData)).toEqual(expected);
  });

  describe('Absolute $ references', () => {
    it.each([
      ['Sheet $col',       '=Sheet1!$A1',                42,  { 'Sheet1!A1': 42 }],
      ['Sheet $col$row',   '=Sheet1!$A$1',               77,  { 'Sheet1!A1': 77 }],
      ['Sheet $row',       '=Sheet1!A$1',                88,  { 'Sheet1!A1': 88 }],
      ['Quoted $ ref',     "='test test'!$B$2",          123, { 'test test!B2': 123 }],
      ['Mixed $ math',     '=Sheet1!$A$1 + Sheet2!B$2',  150, { 'Sheet1!A1': 100, 'Sheet2!B2': 50 }],
      ['SUM with $ range', '=SUM(Sheet1!$A$1:$A$3)',      60, { 'Sheet1!A1': 10, 'Sheet1!A2': 20, 'Sheet1!A3': 30 }],
    ])('%s', async (_name, formula, expected, cellData) => {
      expect(await evalWithCells(formula, cellData)).toEqual(expected);
    });
  });
});
