import { createBuiltins } from './builtins.js';
import { ERRORS } from './errors.js';
import { makeParser } from './parser.js';
import { tokenize } from './tokenizer.js';

// Formula engine with function registry, locale support, error handling, and cross-sheet references
// evaluateFormula(formulaText, { resolveCell }) -> Promise resolving to value
// resolveCell receives cell references in format "A1" or "SheetName!A1" for cross-sheet refs
export function createFormulaEngine(options = {}) {
  const localeMap = options.localeMap || {};
  const builtins = createBuiltins();
  const customFunctions = {};

  // --- Performance: tokenization cache (keyed by formula text) ---
  const tokenCache = new Map();
  const TOKEN_CACHE_MAX = 2048;

  function getCachedTokens(text) {
    if (tokenCache.has(text)) return tokenCache.get(text);
    const tokens = tokenize(text);
    if (tokenCache.size >= TOKEN_CACHE_MAX) {
      // evict oldest entry (first key)
      const first = tokenCache.keys().next().value;
      tokenCache.delete(first);
    }
    tokenCache.set(text, tokens);
    return tokens;
  }

  // Register a custom function
  function registerFunction(name, fn) {
    customFunctions[name.toUpperCase()] = fn;
  }

  // Resolve localized function name to canonical
  function resolveFunction(name) {
    const upper = name.toUpperCase();
    // check locale map first
    if (localeMap[upper]) return localeMap[upper].toUpperCase();
    if (localeMap[name]) return localeMap[name].toUpperCase();
    return upper;
  }

  async function evaluateFormula(text, context = {}) {
    if (!text || text.trim() === '') return '';
    // Use cached tokens (deep-clone positions are reset by makeParser via index)
    const tokens = getCachedTokens(text);
    const evaluating = new Set(); // cycle detection

    // --- Performance: per-evaluation resolve cache ---
    // Avoids duplicate async fetches for the same cell ref within one formula eval
    const resolveCache = new Map();

    const resolveCell = async (ref) => {
      if (!context || !context.resolveCell) return 0;
      if (evaluating.has(ref)) return ERRORS.CYCLE;
      if (resolveCache.has(ref)) return resolveCache.get(ref);
      evaluating.add(ref);
      try {
        const v = await context.resolveCell(ref);
        const result = v ?? '';
        resolveCache.set(ref, result);
        return result;
      } catch (err) {
        resolveCache.set(ref, ERRORS.REF);
        return ERRORS.REF;
      } finally {
        evaluating.delete(ref);
      }
    };

    // Batch resolve helper — resolves multiple refs in one go, backed by per-eval cache
    const resolveCellsBatch = async (refs) => {
      // Check which refs need fetching
      const results = new Array(refs.length);
      const toFetch = [];
      const toFetchIdx = [];
      for (let i = 0; i < refs.length; i++) {
        if (resolveCache.has(refs[i])) {
          results[i] = resolveCache.get(refs[i]);
        } else {
          toFetch.push(refs[i]);
          toFetchIdx.push(i);
        }
      }
      if (toFetch.length > 0) {
        // If context provides a batch resolver, use it; otherwise resolve individually
        if (context.resolveCellsBatch) {
          const batchVals = await context.resolveCellsBatch(toFetch);
          for (let j = 0; j < toFetch.length; j++) {
            const val = batchVals[j] ?? '';
            resolveCache.set(toFetch[j], val);
            results[toFetchIdx[j]] = val;
          }
        } else {
          // Parallel individual resolves for uncached refs
          const promises = toFetch.map(ref => resolveCell(ref));
          const vals = await Promise.all(promises);
          for (let j = 0; j < toFetch.length; j++) {
            results[toFetchIdx[j]] = vals[j];
          }
        }
      }
      return results;
    };

    const parser = makeParser(tokens, resolveCell, { resolveFunction, customFunctions, builtins, resolveCellsBatch });
    try {
      const val = await parser.parseExpression();
      return val;
    } catch (err) {
      return ERRORS.VALUE;
    }
  }

  /** Clear tokenization cache (call after locale change or when memory is a concern) */
  function clearCache() {
    tokenCache.clear();
  }

  return { evaluateFormula, registerFunction, clearCache, ERRORS };
}
