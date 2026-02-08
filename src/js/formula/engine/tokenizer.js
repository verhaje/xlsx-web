// Tokenizer: supports numbers, strings, booleans, cell refs, ranges, function names, operators, cross-sheet refs
export function tokenize(input) {
  const src = input.trim().replace(/^=/, '');
  const tokens = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    // skip whitespace
    if (/\s/.test(ch)) { i++; continue; }
    // string literal
    if (ch === '"') {
      let str = ''; i++;
      while (i < src.length && src[i] !== '"') { str += src[i++]; }
      i++; // skip closing quote
      tokens.push({ type: 'STRING', value: str });
      continue;
    }
    // number
    if (/\d/.test(ch) || (ch === '.' && /\d/.test(src[i + 1]))) {
      let num = '';
      while (i < src.length && /[\d.]/.test(src[i])) { num += src[i++]; }
      tokens.push({ type: 'NUMBER', value: Number(num) });
      continue;
    }
    // Quoted sheet name: support embedded single-quote by doubling ('')
    // Examples: 'Sheet Name'!A1 or 'O''Brian'!A1
    if (ch === "'") {
      let sheet = '';
      i++; // skip opening '
      while (i < src.length) {
        // doubled single-quote represents a literal '
        if (src[i] === "'" && src[i + 1] === "'") { sheet += "'"; i += 2; continue; }
        if (src[i] === "'") break;
        sheet += src[i++];
      }
      if (i < src.length && src[i] === "'") i++; // skip closing quote if present
      // expect ! for sheet reference
      if (src[i] === '!') {
        i++; // consume !
        // parse cell ref or range
        let cellOrRange = '';
        while (i < src.length && /[A-Za-z0-9_$:]/.test(src[i])) { cellOrRange += src[i++]; }
        const cleaned = cellOrRange.replace(/\$/g, '');
        if (cleaned.includes(':')) {
          const [start, end] = cleaned.split(':');
          tokens.push({ type: 'RANGE', value: `${start.toUpperCase()}:${end.toUpperCase()}`, sheet });
        } else if (/^[A-Za-z]+\d+$/i.test(cleaned)) {
          tokens.push({ type: 'CELL', value: cleaned.toUpperCase(), sheet });
        } else {
          tokens.push({ type: 'IDENT', value: cleaned.toUpperCase() });
        }
        // tolerate an extra stray quote immediately after the ref (some files may include it)
        if (src[i] === "'") i++;
      } else {
        // treat as simple string literal when not followed by '!'
        tokens.push({ type: 'STRING', value: sheet });
      }
      continue;
    }
    // identifier (function name or cell ref or range or sheet!ref)
    if (/[A-Za-z_$]/.test(ch)) {
      let id = '';
      while (i < src.length && /[A-Za-z0-9_$]/.test(src[i])) { id += src[i++]; }
      // check for sheet reference (Sheet1!A1)
      if (src[i] === '!') {
        const sheet = id;
        i++; // consume !
        // parse cell ref or range
        let cellOrRange = '';
        while (i < src.length && /[A-Za-z0-9_$:]/.test(src[i])) { cellOrRange += src[i++]; }
        const cleaned = cellOrRange.replace(/\$/g, '');
        if (cleaned.includes(':')) {
          const [start, end] = cleaned.split(':');
          tokens.push({ type: 'RANGE', value: `${start.toUpperCase()}:${end.toUpperCase()}`, sheet });
        } else if (/^[A-Za-z]+\d+$/i.test(cleaned)) {
          tokens.push({ type: 'CELL', value: cleaned.toUpperCase(), sheet });
        } else {
          tokens.push({ type: 'IDENT', value: cleaned.toUpperCase() });
        }
        continue;
      }
      // check for range (A1:B2)
      if (src[i] === ':') {
        i++; // consume :
        let id2 = '';
        while (i < src.length && /[A-Za-z0-9_$]/.test(src[i])) { id2 += src[i++]; }
        const clean1 = id.replace(/\$/g, '');
        const clean2 = id2.replace(/\$/g, '');
        tokens.push({ type: 'RANGE', value: `${clean1.toUpperCase()}:${clean2.toUpperCase()}` });
        continue;
      }
      // check if function (followed by '(')
      let j = i; while (j < src.length && /\s/.test(src[j])) j++;
      if (src[j] === '(') {
        tokens.push({ type: 'FUNC', value: id.toUpperCase() });
      } else {
        const cleanId = id.replace(/\$/g, '');
        if (/^[A-Za-z]+\d+$/i.test(cleanId)) {
          tokens.push({ type: 'CELL', value: cleanId.toUpperCase() });
        } else if (id.toUpperCase() === 'TRUE') {
          tokens.push({ type: 'BOOL', value: true });
        } else if (id.toUpperCase() === 'FALSE') {
          tokens.push({ type: 'BOOL', value: false });
        } else {
          tokens.push({ type: 'IDENT', value: id.toUpperCase() });
        }
      }
      continue;
    }
    // operators and punctuation
    if ('+-*/^(),:;&<>=!'.includes(ch)) {
      // handle multi-char operators
      if (ch === '<' && src[i + 1] === '=') { tokens.push({ type: 'OP', value: '<=' }); i += 2; continue; }
      if (ch === '>' && src[i + 1] === '=') { tokens.push({ type: 'OP', value: '>=' }); i += 2; continue; }
      if (ch === '<' && src[i + 1] === '>') { tokens.push({ type: 'OP', value: '<>' }); i += 2; continue; }
      if (ch === '=' && src[i + 1] === '=') { tokens.push({ type: 'OP', value: '==' }); i += 2; continue; }
      tokens.push({ type: 'OP', value: ch });
      i++;
      continue;
    }
    // skip unknown
    i++;
  }
  tokens.push({ type: 'EOF' });
  return tokens;
}
