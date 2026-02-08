import { ERRORS, isError } from './errors.js';
import { expandRange, getRangeDimensions } from './range-utils.js';

export function makeParser(tokens, resolveCell, options) {
  const { resolveFunction, customFunctions, builtins, resolveCellsBatch } = options;
  let pos = 0;
  function peek() { return tokens[pos]; }
  function consume() { return tokens[pos++]; }

  // Parse argument list, returning both values and metadata
  async function parseArgListWithMeta() {
    const args = [];
    const meta = [];
    if (peek().type === 'OP' && peek().value === ')') return { args, meta };
    const first = await parseExpressionWithMeta();
    args.push(first.value);
    meta.push(first.meta);
    while (peek().type === 'OP' && (peek().value === ',' || peek().value === ';')) {
      consume();
      const next = await parseExpressionWithMeta();
      args.push(next.value);
      meta.push(next.meta);
    }
    return { args, meta };
  }

  async function parseArgList() {
    const result = await parseArgListWithMeta();
    return result.args;
  }

  async function parsePrimaryWithMeta() {
    const t = peek();
    if (t.type === 'NUMBER') { consume(); return { value: t.value, meta: null }; }
    if (t.type === 'STRING') { consume(); return { value: t.value, meta: null }; }
    if (t.type === 'BOOL') { consume(); return { value: t.value, meta: null }; }
    if (t.type === 'CELL') {
      consume();
      const cellRef = t.sheet ? `${t.sheet}!${t.value}` : t.value;
      const v = await resolveCell(cellRef);
      return { value: v, meta: null };
    }
    if (t.type === 'RANGE') {
      consume();
      const refs = expandRange(t.value);
      const dims = getRangeDimensions(t.value);
      // --- Performance: batch-resolve all refs in a range at once ---
      let vals;
      if (resolveCellsBatch) {
        const cellRefs = t.sheet ? refs.map(ref => `${t.sheet}!${ref}`) : refs;
        vals = await resolveCellsBatch(cellRefs);
      } else {
        vals = [];
        for (const ref of refs) {
          const cellRef = t.sheet ? `${t.sheet}!${ref}` : ref;
          vals.push(await resolveCell(cellRef));
        }
      }
      return { value: vals, meta: { rows: dims.rows, cols: dims.cols } };
    }
    if (t.type === 'FUNC') {
      const fnName = resolveFunction(consume().value);
      const next = consume();
      if (!(next.type === 'OP' && next.value === '(')) throw new Error(`Expected ( after ${fnName}`);
      const { args, meta } = await parseArgListWithMeta();
      const close = consume();
      if (!(close.type === 'OP' && close.value === ')')) throw new Error(`Expected ) after ${fnName} args`);
      // flatten array args for functions
      const flatArgs = [];
      for (const a of args) { if (Array.isArray(a)) flatArgs.push(...a); else flatArgs.push(a); }
      const fn = customFunctions[fnName] || builtins[fnName];
      if (!fn) return { value: ERRORS.NAME, meta: null };
      // Conditional aggregate functions need the original (non-flattened) args to see range/criteria pairs
      if (fnName === 'COUNTIFS' || fnName === 'SUMIF' || fnName === 'SUMIFS' || fnName === 'AVERAGEIF' || fnName === 'AVERAGEIFS') {
        return { value: await fn(args), meta: null };
      }
      // Lookup functions need original args with metadata for row/col info
      if (fnName === 'VLOOKUP' || fnName === 'HLOOKUP' || fnName === 'INDEX' || fnName === 'MATCH') {
        return { value: await fn(args, meta), meta: null };
      }
      return { value: await fn(flatArgs), meta: null };
    }
    if (t.type === 'OP' && t.value === '(') {
      consume();
      const result = await parseExpressionWithMeta();
      const next = consume();
      if (!(next.type === 'OP' && next.value === ')')) throw new Error('Expected )');
      return result;
    }
    if (t.type === 'OP' && (t.value === '+' || t.value === '-')) {
      const op = consume().value;
      const prim = await parsePrimaryWithMeta();
      if (isError(prim.value)) return { value: prim.value, meta: null };
      return { value: op === '-' ? -Number(prim.value) : Number(prim.value), meta: null };
    }
    if (t.type === 'EOF') return { value: 0, meta: null };
    throw new Error(`Unexpected token: ${t.type} ${t.value}`);
  }

  async function parsePrimary() {
    const result = await parsePrimaryWithMeta();
    return result.value;
  }

  async function parsePower() {
    let left = await parsePrimary();
    while (peek().type === 'OP' && peek().value === '^') {
      consume();
      const right = await parsePrimary();
      if (isError(left)) return left;
      if (isError(right)) return right;
      left = Math.pow(Number(left), Number(right));
    }
    return left;
  }

  async function parseTerm() {
    let left = await parsePower();
    while (true) {
      const t = peek();
      if (t.type === 'OP' && (t.value === '*' || t.value === '/')) {
        const op = consume().value;
        const right = await parsePower();
        if (isError(left)) return left;
        if (isError(right)) return right;
        if (op === '*') left = Number(left) * Number(right);
        else { if (Number(right) === 0) return ERRORS.DIV0; left = Number(left) / Number(right); }
      } else break;
    }
    return left;
  }

  async function parseAddSub() {
    let left = await parseTerm();
    while (true) {
      const t = peek();
      if (t.type === 'OP' && (t.value === '+' || t.value === '-')) {
        const op = consume().value;
        const right = await parseTerm();
        if (isError(left)) return left;
        if (isError(right)) return right;
        left = op === '+' ? Number(left) + Number(right) : Number(left) - Number(right);
      } else if (t.type === 'OP' && t.value === '&') {
        consume();
        const right = await parseTerm();
        if (isError(left)) return left;
        if (isError(right)) return right;
        left = String(left) + String(right);
      } else break;
    }
    return left;
  }

  async function parseComparison() {
    let left = await parseAddSub();
    const t = peek();
    if (t.type === 'OP' && ['<', '>', '<=', '>=', '=', '==', '<>'].includes(t.value)) {
      const op = consume().value;
      const right = await parseAddSub();
      if (isError(left)) return left;
      if (isError(right)) return right;
      const l = Number(left), r = Number(right);
      const lNum = !Number.isNaN(l), rNum = !Number.isNaN(r);
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
        const ls = String(left), rs = String(right);
        switch (op) {
          case '=': case '==': return ls === rs;
          case '<>': return ls !== rs;
          default: return ls.localeCompare(rs) < 0 ? (op === '<' || op === '<=') : (op === '>' || op === '>=');
        }
      }
    }
    return left;
  }

  async function parseExpression() {
    return parseComparison();
  }

  // For lookup functions, we need to track metadata through simple arg parsing
  // but for most argument positions (direct range refs), metadata comes from parsePrimaryWithMeta
  async function parseExpressionWithMeta() {
    // For simple direct values that preserve metadata, use parsePrimaryWithMeta
    const t = peek();
    if (t.type === 'RANGE') {
      return await parsePrimaryWithMeta();
    }
    // For everything else (including FUNCs that may be in expressions), evaluate normally
    const value = await parseExpression();
    return { value, meta: null };
  }

  return { parseExpression };
}
