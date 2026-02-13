// FormulaEngine.ts - Main formula engine class

import type { Token, FormulaContext, FormulaEngineOptions, BuiltinFunction, IFormulaEngine } from '../types';
import { ERRORS } from './Errors';
import { Tokenizer } from './Tokenizer';
import { ExpressionParser } from './ExpressionParser';
import { BuiltinFunctions } from './BuiltinFunctions';

/**
 * Formula engine with function registry, locale support, error handling,
 * and cross-sheet reference support.
 *
 * Usage:
 *   const engine = new FormulaEngine({ localeMap });
 *   const value = await engine.evaluateFormula('=SUM(A1:A3)', { resolveCell });
 */
export class FormulaEngine implements IFormulaEngine {
  readonly ERRORS = ERRORS;

  private localeMap: Record<string, string>;
  private builtins: Record<string, BuiltinFunction>;
  private customFunctions: Record<string, BuiltinFunction>;
  private tokenCache: Map<string, Token[]>;

  private static readonly TOKEN_CACHE_MAX = 2048;

  constructor(options: FormulaEngineOptions = {}) {
    this.localeMap = options.localeMap || {};
    this.builtins = BuiltinFunctions.create();
    this.customFunctions = {};
    this.tokenCache = new Map();
  }

  /**
   * Register a custom function.
   */
  registerFunction(name: string, fn: BuiltinFunction): void {
    this.customFunctions[name.toUpperCase()] = fn;
  }

  /**
   * Evaluate a formula string and return the computed value.
   */
  async evaluateFormula(text: string, context: FormulaContext = { resolveCell: async () => 0 }): Promise<any> {
    if (!text || text.trim() === '') return '';

    const tokens = this.getCachedTokens(text);
    const evaluating = new Set<string>();
    const resolveCache = new Map<string, any>();

    const resolveCell = async (ref: string): Promise<any> => {
      if (!context.resolveCell) return 0;
      if (evaluating.has(ref)) return ERRORS.CYCLE;
      if (resolveCache.has(ref)) return resolveCache.get(ref);
      evaluating.add(ref);
      try {
        const v = await context.resolveCell(ref);
        const result = v ?? '';
        resolveCache.set(ref, result);
        return result;
      } catch {
        resolveCache.set(ref, ERRORS.REF);
        return ERRORS.REF;
      } finally {
        evaluating.delete(ref);
      }
    };

    const resolveCellsBatch = async (refs: string[]): Promise<any[]> => {
      const results = new Array(refs.length);
      const toFetch: string[] = [];
      const toFetchIdx: number[] = [];
      for (let i = 0; i < refs.length; i++) {
        if (resolveCache.has(refs[i])) {
          results[i] = resolveCache.get(refs[i]);
        } else {
          toFetch.push(refs[i]);
          toFetchIdx.push(i);
        }
      }
      if (toFetch.length > 0) {
        if (context.resolveCellsBatch) {
          const batchVals = await context.resolveCellsBatch(toFetch);
          for (let j = 0; j < toFetch.length; j++) {
            const val = batchVals[j] ?? '';
            resolveCache.set(toFetch[j], val);
            results[toFetchIdx[j]] = val;
          }
        } else {
          const promises = toFetch.map((ref) => resolveCell(ref));
          const vals = await Promise.all(promises);
          for (let j = 0; j < toFetch.length; j++) {
            results[toFetchIdx[j]] = vals[j];
          }
        }
      }
      return results;
    };

    const parser = new ExpressionParser(tokens, resolveCell, {
      resolveFunction: (name) => this.resolveFunction(name),
      customFunctions: this.customFunctions,
      builtins: this.builtins,
      resolveCellsBatch,
    });

    try {
      return await parser.parseExpression();
    } catch {
      return ERRORS.VALUE;
    }
  }

  /**
   * Clear the tokenization cache.
   */
  clearCache(): void {
    this.tokenCache.clear();
  }

  // ---- Private helpers ----

  private resolveFunction(name: string): string {
    const upper = name.toUpperCase();
    if (this.localeMap[upper]) return this.localeMap[upper].toUpperCase();
    if (this.localeMap[name]) return this.localeMap[name].toUpperCase();
    return upper;
  }

  private getCachedTokens(text: string): Token[] {
    if (this.tokenCache.has(text)) return this.tokenCache.get(text)!;
    const tokens = Tokenizer.tokenize(text);
    if (this.tokenCache.size >= FormulaEngine.TOKEN_CACHE_MAX) {
      const first = this.tokenCache.keys().next().value;
      if (first !== undefined) this.tokenCache.delete(first);
    }
    this.tokenCache.set(text, tokens);
    return tokens;
  }
}

/**
 * Factory function for backwards compatibility.
 */
export function createFormulaEngine(options: FormulaEngineOptions = {}): FormulaEngine {
  return new FormulaEngine(options);
}
