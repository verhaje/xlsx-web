// Tokenizer.ts - Tokenizes Excel formula strings

import type { Token, TokenType } from '../types';

/**
 * Tokenizes an Excel formula string into an array of tokens.
 * Supports numbers, strings, booleans, cell references, ranges,
 * function names, operators, and cross-sheet references.
 */
export class Tokenizer {
  /**
   * Tokenize a formula string.
   */
  static tokenize(input: string): Token[] {
    const src = input.trim().replace(/^=/, '');
    const tokens: Token[] = [];
    let i = 0;

    while (i < src.length) {
      const ch = src[i];

      // Skip whitespace
      if (/\s/.test(ch)) {
        i++;
        continue;
      }

      // String literal
      if (ch === '"') {
        let str = '';
        i++;
        while (i < src.length && src[i] !== '"') {
          str += src[i++];
        }
        i++; // skip closing quote
        tokens.push({ type: 'STRING', value: str });
        continue;
      }

      // Number
      if (/\d/.test(ch) || (ch === '.' && /\d/.test(src[i + 1]))) {
        let num = '';
        while (i < src.length && /[\d.]/.test(src[i])) {
          num += src[i++];
        }
        tokens.push({ type: 'NUMBER', value: Number(num) });
        continue;
      }

      // Quoted sheet name: 'Sheet Name'!A1
      if (ch === "'") {
        let sheet = '';
        i++; // skip opening '
        while (i < src.length) {
          if (src[i] === "'" && src[i + 1] === "'") {
            sheet += "'";
            i += 2;
            continue;
          }
          if (src[i] === "'") break;
          sheet += src[i++];
        }
        if (i < src.length && src[i] === "'") i++; // skip closing quote
        if (src[i] === '!') {
          i++; // consume !
          let cellOrRange = '';
          while (i < src.length && /[A-Za-z0-9_$:]/.test(src[i])) {
            cellOrRange += src[i++];
          }
          const cleaned = cellOrRange.replace(/\$/g, '');
          if (cleaned.includes(':')) {
            const [start, end] = cleaned.split(':');
            tokens.push({ type: 'RANGE', value: `${start.toUpperCase()}:${end.toUpperCase()}`, sheet });
          } else if (/^[A-Za-z]+\d+$/i.test(cleaned)) {
            tokens.push({ type: 'CELL', value: cleaned.toUpperCase(), sheet });
          } else {
            tokens.push({ type: 'IDENT', value: cleaned.toUpperCase() });
          }
          if (src[i] === "'") i++; // tolerate stray quote
        } else {
          tokens.push({ type: 'STRING', value: sheet });
        }
        continue;
      }

      // Identifier (function name, cell ref, range, or sheet!ref)
      if (/[A-Za-z_$]/.test(ch)) {
        let id = '';
        while (i < src.length && /[A-Za-z0-9_$]/.test(src[i])) {
          id += src[i++];
        }
        // Support dot-notation function names like MODE.SNGL, STDEV.S, ERROR.TYPE
        while (i < src.length && src[i] === '.' && i + 1 < src.length && /[A-Za-z]/.test(src[i + 1])) {
          id += src[i++]; // consume '.'
          while (i < src.length && /[A-Za-z0-9_$]/.test(src[i])) {
            id += src[i++];
          }
        }

        // Sheet reference (Sheet1!A1)
        if (src[i] === '!') {
          const sheet = id;
          i++; // consume !
          let cellOrRange = '';
          while (i < src.length && /[A-Za-z0-9_$:]/.test(src[i])) {
            cellOrRange += src[i++];
          }
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

        // Range (A1:B2)
        if (src[i] === ':') {
          i++; // consume :
          let id2 = '';
          while (i < src.length && /[A-Za-z0-9_$]/.test(src[i])) {
            id2 += src[i++];
          }
          const clean1 = id.replace(/\$/g, '');
          const clean2 = id2.replace(/\$/g, '');
          tokens.push({ type: 'RANGE', value: `${clean1.toUpperCase()}:${clean2.toUpperCase()}` });
          continue;
        }

        // Function (followed by '(')
        let j = i;
        while (j < src.length && /\s/.test(src[j])) j++;
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

      // Operators and punctuation
      if ('+-*/^(),:;&<>=!{}'.includes(ch)) {
        if (ch === '<' && src[i + 1] === '=') {
          tokens.push({ type: 'OP', value: '<=' });
          i += 2;
          continue;
        }
        if (ch === '>' && src[i + 1] === '=') {
          tokens.push({ type: 'OP', value: '>=' });
          i += 2;
          continue;
        }
        if (ch === '<' && src[i + 1] === '>') {
          tokens.push({ type: 'OP', value: '<>' });
          i += 2;
          continue;
        }
        if (ch === '=' && src[i + 1] === '=') {
          tokens.push({ type: 'OP', value: '==' });
          i += 2;
          continue;
        }
        tokens.push({ type: 'OP', value: ch });
        i++;
        continue;
      }

      // Skip unknown characters
      i++;
    }

    tokens.push({ type: 'EOF', value: '' });
    return tokens;
  }
}
