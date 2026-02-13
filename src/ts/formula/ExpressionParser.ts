// ExpressionParser.ts - Recursive descent parser for Excel formula expressions

import type { Token, BuiltinFunction } from '../types';
import { ERRORS, isError } from './Errors';
import { RangeUtils } from './RangeUtils';

interface ExpressionParserOptions {
  resolveFunction: (name: string) => string;
  customFunctions: Record<string, BuiltinFunction>;
  builtins: Record<string, BuiltinFunction>;
  resolveCellsBatch?: (refs: string[]) => Promise<any[]>;
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
      const refs = RangeUtils.expandRange(String(t.value));
      const dims = RangeUtils.getRangeDimensions(String(t.value));
      let vals: any[];
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
      return { value: vals, meta: { rows: dims.rows, cols: dims.cols } };
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
      const KEEP_ARGS_FUNCS = ['COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS'];
      const LOOKUP_FUNCS = ['VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH'];

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
