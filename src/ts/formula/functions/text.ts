// text.ts - Text and string manipulation functions

import { ERRORS } from '../Errors';
import type { BuiltinFunction } from '../../types';

export function getTextFunctions(): Record<string, BuiltinFunction> {
  return {
    CONCAT: async (args) => args.map((a: any) => String(a ?? '')).join(''),
    CONCATENATE: async (args) => args.map((a: any) => String(a ?? '')).join(''),
    LEFT: async (args) => String(args[0] ?? '').slice(0, Number(args[1]) || 1),
    RIGHT: async (args) => {
      const s = String(args[0] ?? '');
      const n = Number(args[1]) || 1;
      return s.slice(-n);
    },
    MID: async (args) => {
      const s = String(args[0] ?? '');
      const start = (Number(args[1]) || 1) - 1;
      const len = Number(args[2]) || 1;
      return s.slice(start, start + len);
    },
    LEN: async (args) => String(args[0] ?? '').length,
    LOWER: async (args) => String(args[0] ?? '').toLowerCase(),
    UPPER: async (args) => String(args[0] ?? '').toUpperCase(),
    TRIM: async (args) => String(args[0] ?? '').trim(),
    TEXT: async (args) => String(args[0] ?? ''),
    VALUE: async (args) => {
      const s = String(args[0] ?? '').trim();
      const n = Number(s);
      return Number.isNaN(n) ? ERRORS.VALUE : n;
    },
    SUBSTITUTE: async (args) => {
      const text = String(args[0] ?? '');
      const oldText = String(args[1] ?? '');
      const newText = String(args[2] ?? '');
      const instance = args[3] !== undefined ? Number(args[3]) : 0;
      if (oldText === '') return text;
      if (instance <= 0) return text.split(oldText).join(newText);
      let count = 0;
      let result = '';
      let pos = 0;
      while (pos < text.length) {
        const idx = text.indexOf(oldText, pos);
        if (idx === -1) { result += text.slice(pos); break; }
        count++;
        if (count === instance) {
          result += text.slice(pos, idx) + newText;
          result += text.slice(idx + oldText.length);
          break;
        }
        result += text.slice(pos, idx + oldText.length);
        pos = idx + oldText.length;
      }
      return result;
    },
    FIND: async (args) => {
      const findText = String(args[0] ?? '');
      const withinText = String(args[1] ?? '');
      const startNum = args[2] !== undefined ? Number(args[2]) : 1;
      if (startNum < 1) return ERRORS.VALUE;
      const idx = withinText.indexOf(findText, startNum - 1);
      return idx === -1 ? ERRORS.VALUE : idx + 1;
    },
    SEARCH: async (args) => {
      const findText = String(args[0] ?? '').toLowerCase();
      const withinText = String(args[1] ?? '').toLowerCase();
      const startNum = args[2] !== undefined ? Number(args[2]) : 1;
      if (startNum < 1) return ERRORS.VALUE;
      if (/[*?]/.test(findText)) {
        const esc = findText.replace(/[-\\^$+.()|[\]{}]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
        try {
          const re = new RegExp(esc, 'i');
          const sub = withinText.slice(startNum - 1);
          const m = re.exec(sub);
          return m ? m.index + startNum : ERRORS.VALUE;
        } catch { return ERRORS.VALUE; }
      }
      const idx = withinText.indexOf(findText, startNum - 1);
      return idx === -1 ? ERRORS.VALUE : idx + 1;
    },
    REPLACE: async (args) => {
      const oldText = String(args[0] ?? '');
      const startNum = (Number(args[1]) || 1) - 1;
      const numChars = Number(args[2]) || 0;
      const newText = String(args[3] ?? '');
      return oldText.slice(0, startNum) + newText + oldText.slice(startNum + numChars);
    },
    REPT: async (args) => {
      const text = String(args[0] ?? '');
      const times = Math.trunc(Number(args[1]) || 0);
      if (times < 0) return ERRORS.VALUE;
      if (times === 0) return '';
      return text.repeat(times);
    },
    EXACT: async (args) => String(args[0] ?? '') === String(args[1] ?? ''),
    CHAR: async (args) => {
      const n = Math.trunc(Number(args[0]));
      if (n < 1 || n > 255) return ERRORS.VALUE;
      return String.fromCharCode(n);
    },
    CODE: async (args) => {
      const s = String(args[0] ?? '');
      if (s.length === 0) return ERRORS.VALUE;
      return s.charCodeAt(0);
    },
    CLEAN: async (args) => {
      const s = String(args[0] ?? '');
      return s.replace(/[\x00-\x1F]/g, '');
    },
    PROPER: async (args) => {
      const s = String(args[0] ?? '');
      return s.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
    },
    DOLLAR: async (args) => {
      const num = Number(args[0]) || 0;
      const decimals = args[1] !== undefined ? Number(args[1]) : 2;
      const rounded = Math.round(num * Math.pow(10, decimals)) / Math.pow(10, decimals);
      const formatted = Math.abs(rounded).toLocaleString('en-US', { minimumFractionDigits: Math.max(0, decimals), maximumFractionDigits: Math.max(0, decimals) });
      return rounded < 0 ? `($${formatted})` : `$${formatted}`;
    },
    FIXED: async (args) => {
      const num = Number(args[0]) || 0;
      const decimals = args[1] !== undefined ? Number(args[1]) : 2;
      const noCommas = args[2] ? true : false;
      const rounded = Math.round(num * Math.pow(10, decimals)) / Math.pow(10, decimals);
      if (noCommas) return rounded.toFixed(Math.max(0, decimals));
      return rounded.toLocaleString('en-US', { minimumFractionDigits: Math.max(0, decimals), maximumFractionDigits: Math.max(0, decimals) });
    },
    TEXTJOIN: async (args) => {
      const delimiter = String(args[0] ?? '');
      const ignoreEmpty = args[1] ? true : false;
      const texts: string[] = [];
      for (let i = 2; i < args.length; i++) {
        const v = args[i];
        if (Array.isArray(v)) {
          for (const item of v) {
            if (ignoreEmpty && (item === '' || item === null || item === undefined)) continue;
            texts.push(String(item ?? ''));
          }
        } else {
          if (ignoreEmpty && (v === '' || v === null || v === undefined)) continue;
          texts.push(String(v ?? ''));
        }
      }
      return texts.join(delimiter);
    },
    NUMBERVALUE: async (args) => {
      let text = String(args[0] ?? '').trim();
      const decSep = args[1] !== undefined ? String(args[1]) : '.';
      const grpSep = args[2] !== undefined ? String(args[2]) : ',';
      if (grpSep) text = text.split(grpSep).join('');
      if (decSep !== '.') text = text.replace(decSep, '.');
      let pctCount = 0;
      while (text.endsWith('%')) { text = text.slice(0, -1); pctCount++; }
      const n = Number(text);
      if (Number.isNaN(n)) return ERRORS.VALUE;
      return n / Math.pow(100, pctCount);
    },
    TEXTBEFORE: async (args) => {
      const text = String(args[0] ?? '');
      const delimiter = String(args[1] ?? '');
      const instanceNum = args[2] !== undefined ? Number(args[2]) : 1;
      if (delimiter === '') return '';
      if (instanceNum === 0) return ERRORS.VALUE;
      if (instanceNum > 0) {
        let pos = -1;
        for (let i = 0; i < instanceNum; i++) {
          pos = text.indexOf(delimiter, pos + 1);
          if (pos === -1) return ERRORS.NA;
        }
        return text.slice(0, pos);
      } else {
        let pos = text.length;
        for (let i = 0; i < Math.abs(instanceNum); i++) {
          pos = text.lastIndexOf(delimiter, pos - 1);
          if (pos === -1) return ERRORS.NA;
        }
        return text.slice(0, pos);
      }
    },
    TEXTAFTER: async (args) => {
      const text = String(args[0] ?? '');
      const delimiter = String(args[1] ?? '');
      const instanceNum = args[2] !== undefined ? Number(args[2]) : 1;
      if (delimiter === '') return text;
      if (instanceNum === 0) return ERRORS.VALUE;
      if (instanceNum > 0) {
        let pos = -1;
        for (let i = 0; i < instanceNum; i++) {
          pos = text.indexOf(delimiter, pos + 1);
          if (pos === -1) return ERRORS.NA;
        }
        return text.slice(pos + delimiter.length);
      } else {
        let pos = text.length;
        for (let i = 0; i < Math.abs(instanceNum); i++) {
          pos = text.lastIndexOf(delimiter, pos - 1);
          if (pos === -1) return ERRORS.NA;
        }
        return text.slice(pos + delimiter.length);
      }
    },
    T: async (args) => {
      return typeof args[0] === 'string' ? args[0] : '';
    },
    UNICHAR: async (args) => {
      const n = Math.trunc(Number(args[0]));
      if (!Number.isFinite(n) || n < 1 || n > 1114111) return ERRORS.VALUE;
      try { return String.fromCodePoint(n); } catch { return ERRORS.VALUE; }
    },
    UNICODE: async (args) => {
      const s = String(args[0]);
      if (s.length === 0) return ERRORS.VALUE;
      return s.codePointAt(0) || ERRORS.VALUE;
    },
    TEXTSPLIT: async (args) => {
      const text = String(args[0]);
      const colDelimiter = args[1] !== undefined ? String(args[1]) : undefined;
      const rowDelimiter = args[2] !== undefined ? String(args[2]) : undefined;
      if (!colDelimiter && !rowDelimiter) return ERRORS.VALUE;
      if (colDelimiter) return text.split(colDelimiter);
      if (rowDelimiter) return text.split(rowDelimiter);
      return text.split(colDelimiter!);
    },
    VALUETOTEXT: async (args) => {
      const v = args[0];
      if (v === null || v === undefined) return '';
      return String(v);
    },
    ARRAYTOTEXT: async (args) => {
      const arr = args[0];
      if (Array.isArray(arr)) return arr.join(', ');
      return String(arr ?? '');
    },
    ENCODEURL: async (args) => {
      return encodeURIComponent(String(args[0] ?? ''));
    },
  };
}
