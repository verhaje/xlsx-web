// math.ts - Math, trigonometry, and combinatorics functions

import { ERRORS, isError } from '../Errors';
import type { BuiltinFunction } from '../../types';

export function getMathFunctions(): Record<string, BuiltinFunction> {
  return {
    // ---- Core Math ----
    SUM: async (args) => {
      let total = 0;
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) total += n; }
      return total;
    },
    AVERAGE: async (args) => {
      let total = 0; let count = 0;
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) { total += n; count++; } }
      return count === 0 ? ERRORS.DIV0 : total / count;
    },
    MIN: async (args) => {
      let min = Infinity; let found = false;
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) { if (n < min) min = n; found = true; } }
      return found ? min : 0;
    },
    MAX: async (args) => {
      let max = -Infinity; let found = false;
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) { if (n > max) max = n; found = true; } }
      return found ? max : 0;
    },
    COUNT: async (args) => {
      let count = 0;
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) count++; }
      return count;
    },
    ABS: async (args) => Math.abs(Number(args[0]) || 0),
    ROUND: async (args) => {
      if (isError(args[0])) return args[0];
      const num = Number(args[0]);
      if (Number.isNaN(num)) return ERRORS.VALUE;
      const digits = Number(args[1]) || 0;
      const factor = Math.pow(10, digits);
      return Math.round(num * factor) / factor;
    },
    SQRT: async (args) => {
      const n = Number(args[0]);
      return n < 0 ? ERRORS.NUM : Math.sqrt(n);
    },
    LN: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      if (n <= 0) return ERRORS.NUM;
      return Math.log(n);
    },
    POWER: async (args) => Math.pow(Number(args[0]) || 0, Number(args[1]) || 0),
    MOD: async (args) => {
      const divisor = Number(args[1]);
      return divisor === 0 ? ERRORS.DIV0 : (Number(args[0]) || 0) % divisor;
    },

    // ---- Rounding ----
    ROUNDUP: async (args) => {
      const num = Number(args[0]) || 0;
      const digits = Number(args[1]) || 0;
      const factor = Math.pow(10, digits);
      return num >= 0
        ? Math.ceil(num * factor) / factor
        : -Math.ceil(Math.abs(num) * factor) / factor;
    },
    ROUNDDOWN: async (args) => {
      const num = Number(args[0]) || 0;
      const digits = Number(args[1]) || 0;
      const factor = Math.pow(10, digits);
      return Math.trunc(num * factor) / factor;
    },
    INT: async (args) => Math.floor(Number(args[0]) || 0),
    CEILING: async (args) => {
      const num = Number(args[0]);
      const sig = args[1] !== undefined ? Number(args[1]) : 1;
      if (!Number.isFinite(num)) return ERRORS.VALUE;
      if (sig === 0) return 0;
      if ((num > 0 && sig < 0)) return ERRORS.NUM;
      return Math.ceil(num / sig) * sig;
    },
    'CEILING.MATH': async (args) => {
      const num = Number(args[0]);
      const sig = args[1] !== undefined ? Number(args[1]) : 1;
      const mode = args[2] !== undefined ? Number(args[2]) : 0;
      if (!Number.isFinite(num)) return ERRORS.VALUE;
      if (sig === 0) return 0;
      const s = Math.abs(sig);
      if (num >= 0 || mode === 0) return Math.ceil(num / s) * s;
      return -Math.ceil(Math.abs(num) / s) * s;
    },
    'CEILING.PRECISE': async (args) => {
      const num = Number(args[0]);
      const sig = args[1] !== undefined ? Math.abs(Number(args[1])) : 1;
      if (!Number.isFinite(num)) return ERRORS.VALUE;
      if (sig === 0) return 0;
      return Math.ceil(num / sig) * sig;
    },
    FLOOR: async (args) => {
      const num = Number(args[0]);
      const sig = args[1] !== undefined ? Number(args[1]) : 1;
      if (!Number.isFinite(num)) return ERRORS.VALUE;
      if (sig === 0) return ERRORS.DIV0;
      if ((num > 0 && sig < 0) || (num < 0 && sig > 0)) return ERRORS.NUM;
      return Math.floor(num / sig) * sig;
    },
    'FLOOR.MATH': async (args) => {
      const num = Number(args[0]);
      const sig = args[1] !== undefined ? Number(args[1]) : 1;
      const mode = args[2] !== undefined ? Number(args[2]) : 0;
      if (!Number.isFinite(num)) return ERRORS.VALUE;
      if (sig === 0) return 0;
      const s = Math.abs(sig);
      if (num >= 0 || mode === 0) return Math.floor(num / s) * s;
      return -Math.floor(Math.abs(num) / s) * s;
    },
    'FLOOR.PRECISE': async (args) => {
      const num = Number(args[0]);
      const sig = args[1] !== undefined ? Math.abs(Number(args[1])) : 1;
      if (!Number.isFinite(num)) return ERRORS.VALUE;
      if (sig === 0) return 0;
      return Math.floor(num / sig) * sig;
    },
    TRUNC: async (args) => {
      const num = Number(args[0]) || 0;
      const digits = args[1] !== undefined ? Number(args[1]) : 0;
      const factor = Math.pow(10, digits);
      return Math.trunc(num * factor) / factor;
    },
    EVEN: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      const rounded = n >= 0 ? Math.ceil(n) : Math.floor(n);
      if (rounded % 2 === 0) return rounded;
      return n >= 0 ? rounded + 1 : rounded - 1;
    },
    ODD: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      if (n === 0) return 1;
      const rounded = n >= 0 ? Math.ceil(n) : Math.floor(n);
      if (Math.abs(rounded) % 2 === 1) return rounded;
      return n >= 0 ? rounded + 1 : rounded - 1;
    },
    MROUND: async (args) => {
      const num = Number(args[0]);
      const multiple = Number(args[1]);
      if (!Number.isFinite(num) || !Number.isFinite(multiple)) return ERRORS.VALUE;
      if (multiple === 0) return 0;
      if ((num > 0 && multiple < 0) || (num < 0 && multiple > 0)) return ERRORS.NUM;
      return Math.round(num / multiple) * multiple;
    },

    // ---- Products / Sums ----
    PRODUCT: async (args) => {
      let result = 1; let found = false;
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) { result *= n; found = true; } }
      return found ? result : 0;
    },
    SUMPRODUCT: async (args, meta) => {
      if (!Array.isArray(args) || args.length === 0) return ERRORS.VALUE;
      const arrays: number[][] = [];
      for (const a of args) {
        if (Array.isArray(a)) {
          arrays.push(a.map((v: any) => { const n = Number(v); return Number.isNaN(n) ? 0 : n; }));
        } else {
          return ERRORS.VALUE;
        }
      }
      const len = arrays[0].length;
      for (const arr of arrays) { if (arr.length !== len) return ERRORS.VALUE; }
      let total = 0;
      for (let i = 0; i < len; i++) {
        let prod = 1;
        for (const arr of arrays) prod *= arr[i];
        total += prod;
      }
      return total;
    },
    SUMSQ: async (args) => {
      let total = 0;
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) total += n * n; }
      return total;
    },
    'SUMX2MY2': async (args) => {
      const arrX = args[0];
      const arrY = args[1];
      if (!Array.isArray(arrX) || !Array.isArray(arrY)) return ERRORS.VALUE;
      const len = Math.min(arrX.length, arrY.length);
      if (len === 0) return ERRORS.VALUE;
      let total = 0;
      for (let i = 0; i < len; i++) {
        const x = Number(arrX[i]);
        const y = Number(arrY[i]);
        if (Number.isNaN(x) || Number.isNaN(y)) continue;
        total += x * x - y * y;
      }
      return total;
    },
    'SUMX2PY2': async (args) => {
      const arrX = args[0];
      const arrY = args[1];
      if (!Array.isArray(arrX) || !Array.isArray(arrY)) return ERRORS.VALUE;
      const len = Math.min(arrX.length, arrY.length);
      if (len === 0) return ERRORS.VALUE;
      let total = 0;
      for (let i = 0; i < len; i++) {
        const x = Number(arrX[i]);
        const y = Number(arrY[i]);
        if (Number.isNaN(x) || Number.isNaN(y)) continue;
        total += x * x + y * y;
      }
      return total;
    },
    'SUMXMY2': async (args) => {
      const arrX = args[0];
      const arrY = args[1];
      if (!Array.isArray(arrX) || !Array.isArray(arrY)) return ERRORS.VALUE;
      const len = Math.min(arrX.length, arrY.length);
      if (len === 0) return ERRORS.VALUE;
      let total = 0;
      for (let i = 0; i < len; i++) {
        const x = Number(arrX[i]);
        const y = Number(arrY[i]);
        if (Number.isNaN(x) || Number.isNaN(y)) continue;
        total += (x - y) ** 2;
      }
      return total;
    },

    // ---- Random / Sign / Constants ----
    RAND: async () => Math.random(),
    RANDBETWEEN: async (args) => {
      const bottom = Math.ceil(Number(args[0]));
      const top = Math.floor(Number(args[1]));
      if (!Number.isFinite(bottom) || !Number.isFinite(top)) return ERRORS.VALUE;
      if (bottom > top) return ERRORS.NUM;
      return Math.floor(Math.random() * (top - bottom + 1)) + bottom;
    },
    SIGN: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      return n > 0 ? 1 : n < 0 ? -1 : 0;
    },
    PI: async () => Math.PI,

    // ---- Logarithms / Exponentials ----
    LOG: async (args) => {
      const n = Number(args[0]);
      const base = args[1] !== undefined ? Number(args[1]) : 10;
      if (!Number.isFinite(n) || n <= 0) return ERRORS.NUM;
      if (!Number.isFinite(base) || base <= 0 || base === 1) return ERRORS.NUM;
      return Math.log(n) / Math.log(base);
    },
    LOG10: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n) || n <= 0) return ERRORS.NUM;
      return Math.log10(n);
    },
    EXP: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      return Math.exp(n);
    },

    // ---- Division / Quotient ----
    QUOTIENT: async (args) => {
      const num = Number(args[0]);
      const den = Number(args[1]);
      if (!Number.isFinite(num) || !Number.isFinite(den) || den === 0) return ERRORS.DIV0;
      return Math.trunc(num / den);
    },

    // ---- Angles ----
    RADIANS: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      return n * Math.PI / 180;
    },
    DEGREES: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      return n * 180 / Math.PI;
    },

    // ---- Combinatorics / Factorials ----
    FACT: async (args) => {
      const n = Math.trunc(Number(args[0]));
      if (!Number.isFinite(n) || n < 0) return ERRORS.NUM;
      if (n > 170) return ERRORS.NUM;
      let result = 1;
      for (let i = 2; i <= n; i++) result *= i;
      return result;
    },
    FACTDOUBLE: async (args) => {
      const n = Math.trunc(Number(args[0]));
      if (!Number.isFinite(n) || n < -1) return ERRORS.NUM;
      if (n <= 0) return 1;
      let result = 1;
      for (let i = n; i > 0; i -= 2) result *= i;
      return result;
    },
    COMBIN: async (args) => {
      const n = Math.trunc(Number(args[0]));
      const k = Math.trunc(Number(args[1]));
      if (!Number.isFinite(n) || !Number.isFinite(k) || n < 0 || k < 0 || k > n) return ERRORS.NUM;
      let result = 1;
      for (let i = 0; i < k; i++) result = result * (n - i) / (i + 1);
      return Math.round(result);
    },
    COMBINA: async (args) => {
      const n = Math.trunc(Number(args[0]));
      const k = Math.trunc(Number(args[1]));
      if (!Number.isFinite(n) || !Number.isFinite(k)) return ERRORS.VALUE;
      if (n < 0 || k < 0) return ERRORS.NUM;
      const total = n + k - 1;
      let result = 1;
      for (let i = 0; i < k; i++) result = result * (total - i) / (i + 1);
      return Math.round(result);
    },
    PERMUT: async (args) => {
      const n = Math.trunc(Number(args[0]));
      const k = Math.trunc(Number(args[1]));
      if (!Number.isFinite(n) || !Number.isFinite(k)) return ERRORS.VALUE;
      if (n < 0 || k < 0 || k > n) return ERRORS.NUM;
      let result = 1;
      for (let i = 0; i < k; i++) result *= (n - i);
      return result;
    },
    MULTINOMIAL: async (args) => {
      const nums = args.map((a: any) => Math.trunc(Number(a)));
      if (nums.some((n: number) => !Number.isFinite(n) || n < 0)) return ERRORS.NUM;
      const sum = nums.reduce((s: number, v: number) => s + v, 0);
      let result = 1;
      let denom = 1;
      for (let i = 1; i <= sum; i++) result *= i;
      for (const nn of nums) { let f = 1; for (let i = 2; i <= nn; i++) f *= i; denom *= f; }
      return Math.round(result / denom);
    },

    // ---- GCD / LCM ----
    GCD: async (args) => {
      const gcd2 = (a: number, b: number): number => { a = Math.abs(a); b = Math.abs(b); while (b) { [a, b] = [b, a % b]; } return a; };
      let result = 0;
      for (const a of args) {
        const n = Math.trunc(Number(a));
        if (!Number.isFinite(n) || n < 0) return ERRORS.NUM;
        result = gcd2(result, n);
      }
      return result;
    },
    LCM: async (args) => {
      const gcd2 = (a: number, b: number): number => { a = Math.abs(a); b = Math.abs(b); while (b) { [a, b] = [b, a % b]; } return a; };
      let result = 1; let any = false;
      for (const a of args) {
        const n = Math.trunc(Number(a));
        if (!Number.isFinite(n) || n < 0) return ERRORS.NUM;
        if (n === 0) return 0;
        result = (result * n) / gcd2(result, n);
        any = true;
      }
      return any ? result : 0;
    },

    // ---- Trigonometry ----
    SIN: async (args) => {
      const n = Number(args[0]);
      return Number.isFinite(n) ? Math.sin(n) : ERRORS.VALUE;
    },
    COS: async (args) => {
      const n = Number(args[0]);
      return Number.isFinite(n) ? Math.cos(n) : ERRORS.VALUE;
    },
    TAN: async (args) => {
      const n = Number(args[0]);
      return Number.isFinite(n) ? Math.tan(n) : ERRORS.VALUE;
    },
    ASIN: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n) || n < -1 || n > 1) return ERRORS.NUM;
      return Math.asin(n);
    },
    ACOS: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n) || n < -1 || n > 1) return ERRORS.NUM;
      return Math.acos(n);
    },
    ATAN: async (args) => {
      const n = Number(args[0]);
      return Number.isFinite(n) ? Math.atan(n) : ERRORS.VALUE;
    },
    ATAN2: async (args) => {
      const x = Number(args[0]);
      const y = Number(args[1]);
      if (!Number.isFinite(x) || !Number.isFinite(y)) return ERRORS.VALUE;
      if (x === 0 && y === 0) return ERRORS.DIV0;
      return Math.atan2(y, x);
    },
    SINH: async (args) => {
      const n = Number(args[0]);
      return Number.isFinite(n) ? Math.sinh(n) : ERRORS.VALUE;
    },
    COSH: async (args) => {
      const n = Number(args[0]);
      return Number.isFinite(n) ? Math.cosh(n) : ERRORS.VALUE;
    },
    TANH: async (args) => {
      const n = Number(args[0]);
      return Number.isFinite(n) ? Math.tanh(n) : ERRORS.VALUE;
    },
    ACOT: async (args) => {
      const n = Number(args[0]);
      return Number.isFinite(n) ? Math.atan(1 / n) : ERRORS.VALUE;
    },
    COT: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n) || n === 0) return ERRORS.DIV0;
      return 1 / Math.tan(n);
    },
    CSC: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n) || n === 0) return ERRORS.DIV0;
      return 1 / Math.sin(n);
    },
    SEC: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      const c = Math.cos(n);
      return c === 0 ? ERRORS.DIV0 : 1 / c;
    },
    SQRTPI: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n) || n < 0) return ERRORS.NUM;
      return Math.sqrt(n * Math.PI);
    },

    // ---- Series / Base conversion ----
    SERIESSUM: async (args) => {
      const x = Number(args[0]);
      const n = Number(args[1]);
      const m = Number(args[2]);
      const coeffs = args[3];
      if (!Number.isFinite(x) || !Number.isFinite(n) || !Number.isFinite(m)) return ERRORS.VALUE;
      if (!Array.isArray(coeffs)) return ERRORS.VALUE;
      let result = 0;
      for (let i = 0; i < coeffs.length; i++) {
        const c = Number(coeffs[i]);
        if (Number.isNaN(c)) return ERRORS.VALUE;
        result += c * Math.pow(x, n + i * m);
      }
      return result;
    },
    BASE: async (args) => {
      const num = Math.trunc(Number(args[0]));
      const radix = Math.trunc(Number(args[1]));
      const minLen = args[2] !== undefined ? Math.trunc(Number(args[2])) : 0;
      if (!Number.isFinite(num) || num < 0) return ERRORS.NUM;
      if (radix < 2 || radix > 36) return ERRORS.NUM;
      let result = num.toString(radix).toUpperCase();
      if (minLen > 0) result = result.padStart(minLen, '0');
      return result;
    },
    'DECIMAL': async (args) => {
      const text = String(args[0]);
      const radix = Math.trunc(Number(args[1]));
      if (radix < 2 || radix > 36) return ERRORS.NUM;
      const result = parseInt(text, radix);
      if (Number.isNaN(result)) return ERRORS.NUM;
      return result;
    },
    ROMAN: async (args) => {
      const num = Math.trunc(Number(args[0]));
      if (!Number.isFinite(num) || num < 0 || num > 3999) return ERRORS.VALUE;
      if (num === 0) return '';
      const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
      const syms = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I'];
      let result = ''; let nn = num;
      for (let i = 0; i < vals.length; i++) {
        while (nn >= vals[i]) { result += syms[i]; nn -= vals[i]; }
      }
      return result;
    },
    ARABIC: async (args) => {
      const s = String(args[0]).toUpperCase().trim();
      if (s === '') return 0;
      const map: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
      let result = 0;
      for (let i = 0; i < s.length; i++) {
        const curr = map[s[i]];
        if (curr === undefined) return ERRORS.VALUE;
        const next = i + 1 < s.length ? map[s[i + 1]] : 0;
        if (curr < (next || 0)) result -= curr;
        else result += curr;
      }
      return result;
    },

    // ---- Subtotal / Aggregate ----
    SUBTOTAL: async (args) => {
      const funcNum = Number(args[0]);
      const vals: any[] = args.slice(1);
      const nums: number[] = [];
      for (const v of vals) {
        const n = Number(v);
        if (!Number.isNaN(n)) nums.push(n);
      }
      const fn = funcNum > 100 ? funcNum - 100 : funcNum;
      switch (fn) {
        case 1: { const s = nums.reduce((a, b) => a + b, 0); return nums.length ? s / nums.length : ERRORS.DIV0; }
        case 2: return nums.length;
        case 3: return vals.filter((v: any) => v !== '' && v !== null && v !== undefined).length;
        case 4: return nums.length ? Math.max(...nums) : 0;
        case 5: return nums.length ? Math.min(...nums) : 0;
        case 6: return nums.reduce((a, b) => a * b, 1);
        case 7: {
          if (nums.length < 2) return ERRORS.DIV0;
          const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
          const ssq = nums.reduce((a, v) => a + (v - mean) ** 2, 0);
          return Math.sqrt(ssq / (nums.length - 1));
        }
        case 8: {
          if (nums.length === 0) return ERRORS.DIV0;
          const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
          const ssq = nums.reduce((a, v) => a + (v - mean) ** 2, 0);
          return Math.sqrt(ssq / nums.length);
        }
        case 9: return nums.reduce((a, b) => a + b, 0);
        case 10: {
          if (nums.length < 2) return ERRORS.DIV0;
          const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
          return nums.reduce((a, v) => a + (v - mean) ** 2, 0) / (nums.length - 1);
        }
        case 11: {
          if (nums.length === 0) return ERRORS.DIV0;
          const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
          return nums.reduce((a, v) => a + (v - mean) ** 2, 0) / nums.length;
        }
        default: return ERRORS.VALUE;
      }
    },
    AGGREGATE: async (args) => {
      const fnNum = Math.trunc(Number(args[0]));
      const _options = Math.trunc(Number(args[1]));
      const values: number[] = [];
      for (let i = 2; i < args.length; i++) {
        const n = Number(args[i]);
        if (!Number.isNaN(n)) values.push(n);
      }
      if (values.length === 0) return ERRORS.VALUE;
      switch (fnNum) {
        case 1: return values.reduce((a, b) => a + b, 0) / values.length;
        case 2: return values.length;
        case 3: return values.length;
        case 4: return Math.max(...values);
        case 5: return Math.min(...values);
        case 6: { let p = 1; for (const v of values) p *= v; return p; }
        case 7: {
          const mean = values.reduce((a, b) => a + b, 0) / values.length;
          const ssq = values.reduce((a, b) => a + (b - mean) ** 2, 0);
          return Math.sqrt(ssq / (values.length - 1));
        }
        case 8: {
          const mean = values.reduce((a, b) => a + b, 0) / values.length;
          const ssq = values.reduce((a, b) => a + (b - mean) ** 2, 0);
          return Math.sqrt(ssq / values.length);
        }
        case 9: return values.reduce((a, b) => a + b, 0);
        case 10: {
          const mean = values.reduce((a, b) => a + b, 0) / values.length;
          return values.reduce((a, b) => a + (b - mean) ** 2, 0) / (values.length - 1);
        }
        case 11: {
          const mean = values.reduce((a, b) => a + b, 0) / values.length;
          return values.reduce((a, b) => a + (b - mean) ** 2, 0) / values.length;
        }
        case 12: {
          const sorted = [...values].sort((a, b) => a - b);
          const mid = Math.floor(sorted.length / 2);
          return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
        }
        default: return ERRORS.VALUE;
      }
    },
  };
}
