// ExpressionParser.ts - Recursive descent parser for Excel formula expressions

import type { Token, BuiltinFunction } from '../types';
import { ERRORS, isError } from './Errors';
import { RangeUtils } from './RangeUtils';

interface ExpressionParserOptions {
  resolveFunction: (name: string) => string;
  customFunctions: Record<string, BuiltinFunction>;
  builtins: Record<string, BuiltinFunction>;
  resolveCellsBatch?: (refs: string[]) => Promise<any[]>;
  /** Resolve a contiguous range by numeric coordinates (avoids A1 string expansion). */
  resolveRange?: (sheetName: string | undefined, startRow: number, startCol: number, endRow: number, endCol: number) => Promise<any[]>;
  /** Shared cache for resolved range arrays (keyed by canonical range string). */
  rangeCache?: Map<string, any[]>;
  /** Return the maximum populated row for a given sheet name (used for whole-column ranges). */
  getSheetMaxRow?: (sheetName?: string) => number;
}

interface ParsedWithMeta {
  value: any;
  meta: { rows: number; cols: number } | null;
}

/**
 * Recursive descent expression parser for Excel formulas.
 * Handles operator precedence, function calls, cell/range references,
 * and cross-sheet refs.
 */
export class ExpressionParser {
  private tokens: Token[];
  private pos: number;
  private resolveCell: (ref: string) => Promise<any>;
  private resolveFunction: (name: string) => string;
  private customFunctions: Record<string, BuiltinFunction>;
  private builtins: Record<string, BuiltinFunction>;
  private resolveCellsBatch?: (refs: string[]) => Promise<any[]>;
  private resolveRange?: (sheetName: string | undefined, startRow: number, startCol: number, endRow: number, endCol: number) => Promise<any[]>;
  private rangeCache?: Map<string, any[]>;
  private getSheetMaxRow?: (sheetName?: string) => number;

  constructor(
    tokens: Token[],
    resolveCell: (ref: string) => Promise<any>,
    options: ExpressionParserOptions
  ) {
    this.tokens = tokens;
    this.pos = 0;
    this.resolveCell = resolveCell;
    this.resolveFunction = options.resolveFunction;
    this.customFunctions = options.customFunctions;
    this.builtins = options.builtins;
    this.resolveCellsBatch = options.resolveCellsBatch;
    this.resolveRange = options.resolveRange;
    this.rangeCache = options.rangeCache;
    this.getSheetMaxRow = options.getSheetMaxRow;
  }

  /**
   * Parse the top-level expression and return its value.
   */
  async parseExpression(): Promise<any> {
    return this.parseComparison();
  }

  // ---- Grammar rules ----

  private peek(): Token {
    return this.tokens[this.pos];
  }

  private consume(): Token {
    return this.tokens[this.pos++];
  }

  private async parseArgListWithMeta(): Promise<{ args: any[]; meta: (any | null)[] }> {
    const args: any[] = [];
    const meta: (any | null)[] = [];
    if (this.peek().type === 'OP' && this.peek().value === ')') return { args, meta };
    const first = await this.parseExpressionWithMeta();
    args.push(first.value);
    meta.push(first.meta);
    while (this.peek().type === 'OP' && (this.peek().value === ',' || this.peek().value === ';')) {
      this.consume();
      const next = await this.parseExpressionWithMeta();
      args.push(next.value);
      meta.push(next.meta);
    }
    return { args, meta };
  }

  private async parsePrimaryWithMeta(): Promise<ParsedWithMeta> {
    const t = this.peek();

    if (t.type === 'NUMBER') {
      this.consume();
      return { value: t.value, meta: null };
    }
    if (t.type === 'STRING') {
      this.consume();
      return { value: t.value, meta: null };
    }
    if (t.type === 'BOOL') {
      this.consume();
      return { value: t.value, meta: null };
    }
    if (t.type === 'CELL') {
      this.consume();
      const cellRef = t.sheet ? `${t.sheet}!${t.value}` : String(t.value);
      const v = await this.resolveCell(cellRef);
      return { value: v, meta: null };
    }
    if (t.type === 'RANGE') {
      this.consume();
      const rangeStr = String(t.value);
      const rangeKey = t.sheet ? `${t.sheet}!${rangeStr}` : rangeStr;

      // For whole-column ranges (e.g. A:F), resolve maxRow from the target sheet
      let maxRow: number | undefined;
      if (RangeUtils.isColumnRange(rangeStr) && this.getSheetMaxRow) {
        maxRow = this.getSheetMaxRow(t.sheet);
      }

      let dims = RangeUtils.getRangeDimensions(rangeStr, maxRow);

      // Check range cache first (avoids re-resolving identical ranges across formulas).
      // Only resolved arrays are stored here (never Promises) to prevent circular
      // range dependencies from deadlocking through cached in-flight promises.
      // CellEvaluator.resolveRange handles coalescing of concurrent calls safely
      // via its pendingRanges Map with activeRangeSheets cycle detection.
      if (this.rangeCache?.has(rangeKey)) {
        const cached = this.rangeCache.get(rangeKey)!;
        // For column ranges, derive dims.rows from the actual resolved array
        // to guarantee the metadata matches the data.  getSheetMaxRow may
        // return different values before and after the sheet is loaded, so
        // getRangeDimensions could compute a different row count than what
        // was used when the array was first resolved.
        if (RangeUtils.isColumnRange(rangeStr) && dims.cols > 0) {
          dims = { rows: Math.ceil(cached.length / dims.cols), cols: dims.cols };
        }
        return { value: cached, meta: { rows: dims.rows, cols: dims.cols } };
      }

      let vals: any[];

      // Prefer resolveRange (numeric path) to avoid expanding range to A1 strings
      if (this.resolveRange && !RangeUtils.isColumnRange(rangeStr)) {
        const parts = rangeStr.split(':');
        const s = RangeUtils.parseRef(parts[0]);
        const e = parts.length > 1 ? RangeUtils.parseRef(parts[1]) : s;
        if (s && e) {
          vals = await this.resolveRange(t.sheet ?? undefined, s.row, s.col, e.row, e.col);
        } else {
          // Fallback: single ref or parse failure
          const refs = RangeUtils.expandRange(rangeStr, maxRow);
          if (this.resolveCellsBatch) {
            const cellRefs = t.sheet ? refs.map((ref) => `${t.sheet}!${ref}`) : refs;
            vals = await this.resolveCellsBatch(cellRefs);
          } else {
            vals = [];
            for (const ref of refs) {
              const cellRef = t.sheet ? `${t.sheet}!${ref}` : ref;
              vals.push(await this.resolveCell(cellRef));
            }
          }
        }
      } else if (this.resolveRange && RangeUtils.isColumnRange(rangeStr)) {
        // Whole-column range: resolve via numeric path with row bounds.
        // Cap at the actual sheet extent to avoid creating massive arrays
        // for ranges like A:AAB (704 columns × thousands of rows).
        const parts = rangeStr.split(':');
        const sc = RangeUtils.colLetterToNum(parts[0]);
        const ec = RangeUtils.colLetterToNum(parts[1]);
        const colCount = ec - sc + 1;
        let mr = maxRow && maxRow > 0 ? maxRow : 100;
        // Cap total cells to prevent multi-hundred-KB array allocations for
        // wide column ranges (e.g. A:AAB = 704 cols).  CellEvaluator.resolveRange
        // already clamps to the sheet's actual data extent, so this cap only
        // affects the outer array allocation, not the actual cell iteration.
        const MAX_RANGE_CELLS = 50_000;
        if (colCount > 1) {
          mr = Math.min(mr, Math.max(50, Math.floor(MAX_RANGE_CELLS / colCount)));
        }
        vals = await this.resolveRange(t.sheet ?? undefined, 1, sc, mr, ec);
        // Update dims to match actual resolved rows (fixes metadata mismatch
        // that caused VLOOKUP to iterate beyond the resolved array bounds)
        dims = { rows: mr, cols: ec - sc + 1 };
      } else {
        const refs = RangeUtils.expandRange(rangeStr, maxRow);
        if (this.resolveCellsBatch) {
          const cellRefs = t.sheet ? refs.map((ref) => `${t.sheet}!${ref}`) : refs;
          vals = await this.resolveCellsBatch(cellRefs);
        } else {
          vals = [];
          for (const ref of refs) {
            const cellRef = t.sheet ? `${t.sheet}!${ref}` : ref;
            vals.push(await this.resolveCell(cellRef));
          }
        }
      }

      // Store resolved array in range cache (only resolved arrays, never promises)
      this.rangeCache?.set(rangeKey, vals);

      return { value: vals, meta: { rows: dims.rows, cols: dims.cols } };
    }
    // Array literal: {1,2,3} or {1,2;3,4} (semicolons = row separators)
    if (t.type === 'OP' && t.value === '{') {
      this.consume(); // consume '{'
      const elements: any[] = [];
      if (!(this.peek().type === 'OP' && this.peek().value === '}')) {
        // Parse first element
        elements.push(await this.parseExpression());
        while (this.peek().type === 'OP' && (this.peek().value === ',' || this.peek().value === ';')) {
          this.consume(); // consume ',' or ';'
          elements.push(await this.parseExpression());
        }
      }
      const close = this.consume();
      if (!(close.type === 'OP' && close.value === '}')) throw new Error('Expected } after array literal');
      return { value: elements, meta: { rows: 1, cols: elements.length } };
    }
    if (t.type === 'FUNC') {
      const fnName = this.resolveFunction(this.consume().value as string);
      const next = this.consume();
      if (!(next.type === 'OP' && next.value === '(')) throw new Error(`Expected ( after ${fnName}`);
      const { args, meta } = await this.parseArgListWithMeta();
      const close = this.consume();
      if (!(close.type === 'OP' && close.value === ')')) throw new Error(`Expected ) after ${fnName} args`);

      // Flatten array args for most functions
      const flatArgs: any[] = [];
      for (const a of args) {
        if (Array.isArray(a)) flatArgs.push(...a);
        else flatArgs.push(a);
      }

      const fn = this.customFunctions[fnName] || this.builtins[fnName];
      if (!fn) return { value: ERRORS.NAME, meta: null };

      // Conditional / lookup functions need original (non-flattened) args
      const KEEP_ARGS_FUNCS = [
        'COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS',
        'COUNTIF', 'SUMPRODUCT',
        'LARGE', 'SMALL',
        'RANK', 'RANK.EQ', 'RANK.AVG',
        'MAXIFS', 'MINIFS',
        'PERCENTILE', 'PERCENTILE.INC', 'PERCENTILE.EXC',
        'QUARTILE', 'QUARTILE.INC',
        'FREQUENCY', 'TRANSPOSE',
        // Medium-priority: statistical regression, array functions, financial
        'CORREL', 'COVARIANCE.P', 'COVARIANCE.S',
        'SLOPE', 'INTERCEPT', 'RSQ',
        'FORECAST', 'FORECAST.LINEAR',
        'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2',
        'TRIMMEAN',
        'SORT', 'UNIQUE', 'FILTER', 'SORTBY',
        'IRR', 'MIRR', 'SERIESSUM',
        'TEXTSPLIT', 'ARRAYTOTEXT',
      ];
      const LOOKUP_FUNCS = ['VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP', 'XMATCH', 'LOOKUP', 'FORMULATEXT'];

      if (KEEP_ARGS_FUNCS.includes(fnName)) {
        return { value: await fn(args), meta: null };
      }
      if (LOOKUP_FUNCS.includes(fnName)) {
        return { value: await fn(args, meta), meta: null };
      }

      return { value: await fn(flatArgs), meta: null };
    }
    if (t.type === 'OP' && t.value === '(') {
      this.consume();
      const result = await this.parseExpressionWithMeta();
      const next = this.consume();
      if (!(next.type === 'OP' && next.value === ')')) throw new Error('Expected )');
      return result;
    }
    if (t.type === 'OP' && (t.value === '+' || t.value === '-')) {
      const op = this.consume().value;
      const prim = await this.parsePrimaryWithMeta();
      if (isError(prim.value)) return { value: prim.value, meta: null };
      return { value: op === '-' ? -Number(prim.value) : Number(prim.value), meta: null };
    }
    if (t.type === 'EOF') return { value: 0, meta: null };
    throw new Error(`Unexpected token: ${t.type} ${t.value}`);
  }

  private async parsePrimary(): Promise<any> {
    const result = await this.parsePrimaryWithMeta();
    return result.value;
  }

  private async parsePower(): Promise<any> {
    let left = await this.parsePrimary();
    while (this.peek().type === 'OP' && this.peek().value === '^') {
      this.consume();
      const right = await this.parsePrimary();
      if (isError(left)) return left;
      if (isError(right)) return right;
      left = Math.pow(Number(left), Number(right));
    }
    return left;
  }

  private async parseTerm(): Promise<any> {
    let left = await this.parsePower();
    while (true) {
      const t = this.peek();
      if (t.type === 'OP' && (t.value === '*' || t.value === '/')) {
        const op = this.consume().value;
        const right = await this.parsePower();
        if (isError(left)) return left;
        if (isError(right)) return right;
        if (op === '*') left = Number(left) * Number(right);
        else {
          if (Number(right) === 0) return ERRORS.DIV0;
          left = Number(left) / Number(right);
        }
      } else break;
    }
    return left;
  }

  private async parseAddSub(): Promise<any> {
    let left = await this.parseTerm();
    while (true) {
      const t = this.peek();
      if (t.type === 'OP' && (t.value === '+' || t.value === '-')) {
        const op = this.consume().value;
        const right = await this.parseTerm();
        if (isError(left)) return left;
        if (isError(right)) return right;
        left = op === '+' ? Number(left) + Number(right) : Number(left) - Number(right);
      } else if (t.type === 'OP' && t.value === '&') {
        this.consume();
        const right = await this.parseTerm();
        if (isError(left)) return left;
        if (isError(right)) return right;
        left = String(left) + String(right);
      } else break;
    }
    return left;
  }

  private async parseComparison(): Promise<any> {
    let left = await this.parseAddSub();
    const t = this.peek();
    if (t.type === 'OP' && ['<', '>', '<=', '>=', '=', '==', '<>'].includes(t.value as string)) {
      const op = this.consume().value as string;
      const right = await this.parseAddSub();
      if (isError(left)) return left;
      if (isError(right)) return right;
      const l = Number(left);
      const r = Number(right);
      const lNum = !Number.isNaN(l);
      const rNum = !Number.isNaN(r);
      if (lNum && rNum) {
        switch (op) {
          case '<': return l < r;
          case '>': return l > r;
          case '<=': return l <= r;
          case '>=': return l >= r;
          case '=': case '==': return l === r;
          case '<>': return l !== r;
        }
      } else {
        const ls = String(left);
        const rs = String(right);
        switch (op) {
          case '=': case '==': return ls === rs;
          case '<>': return ls !== rs;
          default: return ls.localeCompare(rs) < 0 ? (op === '<' || op === '<=') : (op === '>' || op === '>=');
        }
      }
    }
    return left;
  }

  private async parseExpressionWithMeta(): Promise<ParsedWithMeta> {
    const t = this.peek();
    if (t.type === 'RANGE') {
      return await this.parsePrimaryWithMeta();
    }
    const value = await this.parseExpression();
    return { value, meta: null };
  }
}
