import { describe, it, expect } from 'vitest';
import { createFormulaEngine } from '../src/ts/formula/FormulaEngine';

describe('Formula Engine Performance', () => {
  const ROWS = 100;
  const COLS = 10;
  const cells: Record<string, number> = {};
  for (let r = 1; r <= ROWS; r++) {
    for (let c = 1; c <= COLS; c++) {
      let col = '';
      let cc = c;
      while (cc > 0) {
        const rem = (cc - 1) % 26;
        col = String.fromCharCode(65 + rem) + col;
        cc = Math.floor((cc - 1) / 26);
      }
      cells[`${col}${r}`] = r * c;
    }
  }

  const resolveCell = async (ref: string) => cells[ref] ?? '';
  const resolveCellsBatch = async (refs: string[]) => refs.map((ref) => cells[ref] ?? '');

  function createCtx(batch = false) {
    const engine = createFormulaEngine({});
    const ctx: any = { resolveCell };
    if (batch) ctx.resolveCellsBatch = resolveCellsBatch;
    return { engine, ctx };
  }

  it('evaluates simple arithmetic quickly', async () => {
    const { engine, ctx } = createCtx();
    const start = performance.now();
    for (let i = 0; i < 1000; i++) {
      await engine.evaluateFormula('=1+2*3-4/2', ctx);
    }
    const elapsed = performance.now() - start;
    // Should complete 1000 iterations in under 5 seconds even on slow machines
    expect(elapsed).toBeLessThan(5000);
  });

  it('evaluates SUM over range', async () => {
    const { engine, ctx } = createCtx();
    const result = await engine.evaluateFormula('=SUM(A1:J100)', ctx);
    expect(typeof result).toBe('number');
    expect(result).toBeGreaterThan(0);
  });

  it('evaluates COUNTIFS', async () => {
    const { engine, ctx } = createCtx();
    const result = await engine.evaluateFormula('=COUNTIFS(A1:A100,">50")', ctx);
    expect(typeof result).toBe('number');
  });

  it('handles batch resolver', async () => {
    const { engine, ctx } = createCtx(true);
    const result = await engine.evaluateFormula('=SUM(A1:J10)', ctx);
    expect(typeof result).toBe('number');
    expect(result).toBeGreaterThan(0);
  });
});
