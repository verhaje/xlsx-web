// statistical.ts - Statistical, counting, and conditional aggregate functions

import { ERRORS } from '../Errors';
import { matchCriteria } from './helpers';
import type { BuiltinFunction } from '../../types';

export function getStatisticalFunctions(): Record<string, BuiltinFunction> {
  return {
    // ---- Conditional Aggregates ----
    COUNTIF: async (args) => {
      try {
        const range = args[0];
        const criteria = args[1];
        if (!Array.isArray(range)) return ERRORS.VALUE;
        let count = 0;
        for (let i = 0; i < range.length; i++) {
          if (matchCriteria(range[i], criteria)) count++;
        }
        return count;
      } catch { return ERRORS.VALUE; }
    },
    COUNTIFS: async (args) => {
      try {
        if (!Array.isArray(args) || args.length === 0) return 0;
        if (args.length % 2 !== 0) return ERRORS.VALUE;
        const pairs: { range: any[]; crit: any }[] = [];
        for (let i = 0; i < args.length; i += 2) {
          const range = args[i];
          const crit = args[i + 1];
          if (!Array.isArray(range)) return ERRORS.VALUE;
          pairs.push({ range, crit });
        }
        const len = pairs[0].range.length;
        for (const p of pairs) if (p.range.length !== len) return ERRORS.VALUE;
        let count = 0;
        for (let i = 0; i < len; i += 1) {
          let ok = true;
          for (const p of pairs) {
            if (!matchCriteria(p.range[i], p.crit)) { ok = false; break; }
          }
          if (ok) count += 1;
        }
        return count;
      } catch { return ERRORS.VALUE; }
    },
    SUMIF: async (args) => {
      try {
        const criteriaRange = args[0];
        const criteria = args[1];
        const sumRange = args[2] !== undefined ? args[2] : criteriaRange;
        if (!Array.isArray(criteriaRange)) return ERRORS.VALUE;
        if (!Array.isArray(sumRange)) return ERRORS.VALUE;
        let sum = 0;
        const len = Math.min(criteriaRange.length, sumRange.length);
        for (let i = 0; i < len; i++) {
          if (matchCriteria(criteriaRange[i], criteria)) {
            const n = Number(sumRange[i]);
            if (!Number.isNaN(n)) sum += n;
          }
        }
        return sum;
      } catch { return ERRORS.VALUE; }
    },
    SUMIFS: async (args) => {
      try {
        if (!Array.isArray(args) || args.length < 3) return ERRORS.VALUE;
        if ((args.length - 1) % 2 !== 0) return ERRORS.VALUE;
        const sumRange = args[0];
        if (!Array.isArray(sumRange)) return ERRORS.VALUE;
        const pairs: { range: any[]; crit: any }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = args[i];
          const crit = args[i + 1];
          if (!Array.isArray(range)) return ERRORS.VALUE;
          pairs.push({ range, crit });
        }
        const len = sumRange.length;
        for (const p of pairs) if (p.range.length !== len) return ERRORS.VALUE;
        let sum = 0;
        for (let i = 0; i < len; i++) {
          let ok = true;
          for (const p of pairs) { if (!matchCriteria(p.range[i], p.crit)) { ok = false; break; } }
          if (ok) { const n = Number(sumRange[i]); if (!Number.isNaN(n)) sum += n; }
        }
        return sum;
      } catch { return ERRORS.VALUE; }
    },
    AVERAGEIF: async (args) => {
      try {
        const criteriaRange = args[0];
        const criteria = args[1];
        const avgRange = args[2] !== undefined ? args[2] : criteriaRange;
        if (!Array.isArray(criteriaRange)) return ERRORS.VALUE;
        if (!Array.isArray(avgRange)) return ERRORS.VALUE;
        let sum = 0; let count = 0;
        const len = Math.min(criteriaRange.length, avgRange.length);
        for (let i = 0; i < len; i++) {
          if (matchCriteria(criteriaRange[i], criteria)) {
            const n = Number(avgRange[i]);
            if (!Number.isNaN(n)) { sum += n; count++; }
          }
        }
        return count === 0 ? ERRORS.DIV0 : sum / count;
      } catch { return ERRORS.VALUE; }
    },
    AVERAGEIFS: async (args) => {
      try {
        if (!Array.isArray(args) || args.length < 3) return ERRORS.VALUE;
        if ((args.length - 1) % 2 !== 0) return ERRORS.VALUE;
        const avgRange = args[0];
        if (!Array.isArray(avgRange)) return ERRORS.VALUE;
        const pairs: { range: any[]; crit: any }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = args[i];
          const crit = args[i + 1];
          if (!Array.isArray(range)) return ERRORS.VALUE;
          pairs.push({ range, crit });
        }
        const len = avgRange.length;
        for (const p of pairs) if (p.range.length !== len) return ERRORS.VALUE;
        let sum = 0; let count = 0;
        for (let i = 0; i < len; i++) {
          let ok = true;
          for (const p of pairs) { if (!matchCriteria(p.range[i], p.crit)) { ok = false; break; } }
          if (ok) { const n = Number(avgRange[i]); if (!Number.isNaN(n)) { sum += n; count++; } }
        }
        return count === 0 ? ERRORS.DIV0 : sum / count;
      } catch { return ERRORS.VALUE; }
    },
    MAXIFS: async (args) => {
      try {
        if (!Array.isArray(args) || args.length < 3) return ERRORS.VALUE;
        if ((args.length - 1) % 2 !== 0) return ERRORS.VALUE;
        const maxRange = args[0];
        if (!Array.isArray(maxRange)) return ERRORS.VALUE;
        const pairs: { range: any[]; crit: any }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = args[i];
          const crit = args[i + 1];
          if (!Array.isArray(range)) return ERRORS.VALUE;
          pairs.push({ range, crit });
        }
        const len = maxRange.length;
        for (const p of pairs) if (p.range.length !== len) return ERRORS.VALUE;
        let result = -Infinity; let found = false;
        for (let i = 0; i < len; i++) {
          let ok = true;
          for (const p of pairs) { if (!matchCriteria(p.range[i], p.crit)) { ok = false; break; } }
          if (ok) { const n = Number(maxRange[i]); if (!Number.isNaN(n)) { if (n > result) result = n; found = true; } }
        }
        return found ? result : 0;
      } catch { return ERRORS.VALUE; }
    },
    MINIFS: async (args) => {
      try {
        if (!Array.isArray(args) || args.length < 3) return ERRORS.VALUE;
        if ((args.length - 1) % 2 !== 0) return ERRORS.VALUE;
        const minRange = args[0];
        if (!Array.isArray(minRange)) return ERRORS.VALUE;
        const pairs: { range: any[]; crit: any }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = args[i];
          const crit = args[i + 1];
          if (!Array.isArray(range)) return ERRORS.VALUE;
          pairs.push({ range, crit });
        }
        const len = minRange.length;
        for (const p of pairs) if (p.range.length !== len) return ERRORS.VALUE;
        let result = Infinity; let found = false;
        for (let i = 0; i < len; i++) {
          let ok = true;
          for (const p of pairs) { if (!matchCriteria(p.range[i], p.crit)) { ok = false; break; } }
          if (ok) { const n = Number(minRange[i]); if (!Number.isNaN(n)) { if (n < result) result = n; found = true; } }
        }
        return found ? result : 0;
      } catch { return ERRORS.VALUE; }
    },

    // ---- Counting ----
    COUNTA: async (args) => {
      let count = 0;
      for (const a of args) { if (a !== '' && a !== null && a !== undefined) count++; }
      return count;
    },
    COUNTBLANK: async (args) => {
      let count = 0;
      for (const a of args) { if (a === '' || a === null || a === undefined) count++; }
      return count;
    },

    // ---- Descriptive Statistics ----
    MEDIAN: async (args) => {
      const nums: number[] = [];
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) nums.push(n); }
      if (nums.length === 0) return ERRORS.NUM;
      nums.sort((a, b) => a - b);
      const mid = nums.length >> 1;
      return (nums.length & 1) ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
    },
    STDEV: async (args) => {
      const nums = args.map((a: any) => Number(a)).filter((n: number) => !Number.isNaN(n));
      const n = nums.length;
      if (n < 2) return ERRORS.DIV0;
      const mean = nums.reduce((sum: number, v: number) => sum + v, 0) / n;
      let sumSq = 0;
      for (const v of nums) sumSq += Math.pow(v - mean, 2);
      return Math.sqrt(sumSq / (n - 1));
    },
    'STDEV.S': async (args) => {
      const nums = args.map((a: any) => Number(a)).filter((n: number) => !Number.isNaN(n));
      const n = nums.length;
      if (n < 2) return ERRORS.DIV0;
      const mean = nums.reduce((s: number, v: number) => s + v, 0) / n;
      let sumSq = 0;
      for (const v of nums) sumSq += (v - mean) ** 2;
      return Math.sqrt(sumSq / (n - 1));
    },
    'STDEV.P': async (args) => {
      const nums = args.map((a: any) => Number(a)).filter((n: number) => !Number.isNaN(n));
      const n = nums.length;
      if (n === 0) return ERRORS.DIV0;
      const mean = nums.reduce((s: number, v: number) => s + v, 0) / n;
      let sumSq = 0;
      for (const v of nums) sumSq += (v - mean) ** 2;
      return Math.sqrt(sumSq / n);
    },
    STDEVP: async (args) => {
      const nums = args.map((a: any) => Number(a)).filter((n: number) => !Number.isNaN(n));
      const n = nums.length;
      if (n === 0) return ERRORS.DIV0;
      const mean = nums.reduce((s: number, v: number) => s + v, 0) / n;
      let sumSq = 0;
      for (const v of nums) sumSq += (v - mean) ** 2;
      return Math.sqrt(sumSq / n);
    },
    VAR: async (args) => {
      const nums = args.map((a: any) => Number(a)).filter((n: number) => !Number.isNaN(n));
      const n = nums.length;
      if (n < 2) return ERRORS.DIV0;
      const mean = nums.reduce((s: number, v: number) => s + v, 0) / n;
      return nums.reduce((s: number, v: number) => s + (v - mean) ** 2, 0) / (n - 1);
    },
    'VAR.S': async (args) => {
      const nums = args.map((a: any) => Number(a)).filter((n: number) => !Number.isNaN(n));
      const n = nums.length;
      if (n < 2) return ERRORS.DIV0;
      const mean = nums.reduce((s: number, v: number) => s + v, 0) / n;
      return nums.reduce((s: number, v: number) => s + (v - mean) ** 2, 0) / (n - 1);
    },
    'VAR.P': async (args) => {
      const nums = args.map((a: any) => Number(a)).filter((n: number) => !Number.isNaN(n));
      const n = nums.length;
      if (n === 0) return ERRORS.DIV0;
      const mean = nums.reduce((s: number, v: number) => s + v, 0) / n;
      return nums.reduce((s: number, v: number) => s + (v - mean) ** 2, 0) / n;
    },
    VARP: async (args) => {
      const nums = args.map((a: any) => Number(a)).filter((n: number) => !Number.isNaN(n));
      const n = nums.length;
      if (n === 0) return ERRORS.DIV0;
      const mean = nums.reduce((s: number, v: number) => s + v, 0) / n;
      return nums.reduce((s: number, v: number) => s + (v - mean) ** 2, 0) / n;
    },

    // ---- Ranking / Ordering ----
    LARGE: async (args) => {
      const data = args[0];
      const k = Math.trunc(Number(args[1]));
      if (!Array.isArray(data)) return ERRORS.VALUE;
      const nums = data.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (k < 1 || k > nums.length) return ERRORS.NUM;
      nums.sort((a: number, b: number) => b - a);
      return nums[k - 1];
    },
    SMALL: async (args) => {
      const data = args[0];
      const k = Math.trunc(Number(args[1]));
      if (!Array.isArray(data)) return ERRORS.VALUE;
      const nums = data.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (k < 1 || k > nums.length) return ERRORS.NUM;
      nums.sort((a: number, b: number) => a - b);
      return nums[k - 1];
    },
    RANK: async (args) => {
      const num = Number(args[0]);
      const ref = args[1];
      const order = args[2] !== undefined ? Number(args[2]) : 0;
      if (!Array.isArray(ref)) return ERRORS.VALUE;
      if (Number.isNaN(num)) return ERRORS.VALUE;
      const nums = ref.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (order === 0) {
        nums.sort((a: number, b: number) => b - a);
      } else {
        nums.sort((a: number, b: number) => a - b);
      }
      const idx = nums.indexOf(num);
      return idx === -1 ? ERRORS.NA : idx + 1;
    },
    'RANK.EQ': async (args) => {
      const num = Number(args[0]);
      const ref = args[1];
      const order = args[2] !== undefined ? Number(args[2]) : 0;
      if (!Array.isArray(ref)) return ERRORS.VALUE;
      if (Number.isNaN(num)) return ERRORS.VALUE;
      const nums = ref.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (order === 0) {
        nums.sort((a: number, b: number) => b - a);
      } else {
        nums.sort((a: number, b: number) => a - b);
      }
      const idx = nums.indexOf(num);
      return idx === -1 ? ERRORS.NA : idx + 1;
    },
    'RANK.AVG': async (args) => {
      const num = Number(args[0]);
      const ref = args[1];
      const order = args[2] !== undefined ? Number(args[2]) : 0;
      if (!Array.isArray(ref)) return ERRORS.VALUE;
      if (Number.isNaN(num)) return ERRORS.VALUE;
      const nums = ref.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (order === 0) {
        nums.sort((a: number, b: number) => b - a);
      } else {
        nums.sort((a: number, b: number) => a - b);
      }
      const positions: number[] = [];
      for (let i = 0; i < nums.length; i++) { if (nums[i] === num) positions.push(i + 1); }
      if (positions.length === 0) return ERRORS.NA;
      return positions.reduce((a, b) => a + b, 0) / positions.length;
    },

    // ---- Percentile / Quartile ----
    'PERCENTILE': async (args) => {
      const data = args[0];
      const k = Number(args[1]);
      if (!Array.isArray(data)) return ERRORS.VALUE;
      if (k < 0 || k > 1) return ERRORS.NUM;
      const nums = data.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (nums.length === 0) return ERRORS.NUM;
      nums.sort((a: number, b: number) => a - b);
      const idx = k * (nums.length - 1);
      const lo = Math.floor(idx);
      const hi = Math.ceil(idx);
      if (lo === hi) return nums[lo];
      return nums[lo] + (nums[hi] - nums[lo]) * (idx - lo);
    },
    'PERCENTILE.INC': async (args) => {
      const data = args[0];
      const k = Number(args[1]);
      if (!Array.isArray(data)) return ERRORS.VALUE;
      if (k < 0 || k > 1) return ERRORS.NUM;
      const nums = data.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (nums.length === 0) return ERRORS.NUM;
      nums.sort((a: number, b: number) => a - b);
      const idx = k * (nums.length - 1);
      const lo = Math.floor(idx);
      const hi = Math.ceil(idx);
      if (lo === hi) return nums[lo];
      return nums[lo] + (nums[hi] - nums[lo]) * (idx - lo);
    },
    'PERCENTILE.EXC': async (args) => {
      const data = args[0];
      const k = Number(args[1]);
      if (!Array.isArray(data)) return ERRORS.VALUE;
      const nums = data.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      const n = nums.length;
      if (n === 0) return ERRORS.NUM;
      if (k <= 1 / (n + 1) || k >= n / (n + 1)) return ERRORS.NUM;
      nums.sort((a: number, b: number) => a - b);
      const idx = k * (n + 1) - 1;
      const lo = Math.floor(idx);
      const hi = Math.ceil(idx);
      if (lo === hi) return nums[lo];
      return nums[lo] + (nums[hi] - nums[lo]) * (idx - lo);
    },
    QUARTILE: async (args) => {
      const data = args[0];
      const quart = Math.trunc(Number(args[1]));
      if (!Array.isArray(data)) return ERRORS.VALUE;
      if (quart < 0 || quart > 4) return ERRORS.NUM;
      const nums = data.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (nums.length === 0) return ERRORS.NUM;
      nums.sort((a: number, b: number) => a - b);
      const k = quart / 4;
      const idx = k * (nums.length - 1);
      const lo = Math.floor(idx);
      const hi = Math.ceil(idx);
      if (lo === hi) return nums[lo];
      return nums[lo] + (nums[hi] - nums[lo]) * (idx - lo);
    },
    'QUARTILE.INC': async (args) => {
      const data = args[0];
      const quart = Math.trunc(Number(args[1]));
      if (!Array.isArray(data)) return ERRORS.VALUE;
      if (quart < 0 || quart > 4) return ERRORS.NUM;
      const nums = data.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (nums.length === 0) return ERRORS.NUM;
      nums.sort((a: number, b: number) => a - b);
      const k = quart / 4;
      const idx = k * (nums.length - 1);
      const lo = Math.floor(idx);
      const hi = Math.ceil(idx);
      if (lo === hi) return nums[lo];
      return nums[lo] + (nums[hi] - nums[lo]) * (idx - lo);
    },
    FREQUENCY: async (args) => {
      const data = args[0];
      const bins = args[1];
      if (!Array.isArray(data) || !Array.isArray(bins)) return ERRORS.VALUE;
      const nums = data.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      const binNums = bins.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n)).sort((a: number, b: number) => a - b);
      const result = new Array(binNums.length + 1).fill(0);
      for (const n of nums) {
        let placed = false;
        for (let i = 0; i < binNums.length; i++) {
          if (n <= binNums[i]) { result[i]++; placed = true; break; }
        }
        if (!placed) result[binNums.length]++;
      }
      return result;
    },
    MODE: async (args) => {
      const nums: number[] = [];
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) nums.push(n); }
      if (nums.length === 0) return ERRORS.NA;
      const freq = new Map<number, number>();
      for (const n of nums) freq.set(n, (freq.get(n) || 0) + 1);
      let maxFreq = 0; let modeVal = nums[0];
      for (const [val, count] of freq) { if (count > maxFreq) { maxFreq = count; modeVal = val; } }
      if (maxFreq <= 1) return ERRORS.NA;
      return modeVal;
    },
    'MODE.SNGL': async (args) => {
      const nums: number[] = [];
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) nums.push(n); }
      if (nums.length === 0) return ERRORS.NA;
      const freq = new Map<number, number>();
      for (const n of nums) freq.set(n, (freq.get(n) || 0) + 1);
      let maxFreq = 0; let modeVal = nums[0];
      for (const [val, count] of freq) { if (count > maxFreq) { maxFreq = count; modeVal = val; } }
      if (maxFreq <= 1) return ERRORS.NA;
      return modeVal;
    },

    // ---- A-variants (include text/boolean) ----
    AVERAGEA: async (args) => {
      let total = 0; let count = 0;
      for (const a of args) {
        if (a === '' || a === null || a === undefined) continue;
        if (typeof a === 'boolean' || a === 'TRUE' || a === 'FALSE') { total += (a === true || a === 'TRUE') ? 1 : 0; count++; }
        else if (typeof a === 'string') { total += 0; count++; }
        else { const n = Number(a); if (!Number.isNaN(n)) { total += n; count++; } }
      }
      return count === 0 ? ERRORS.DIV0 : total / count;
    },
    MAXA: async (args) => {
      let max = -Infinity; let found = false;
      for (const a of args) {
        if (a === '' || a === null || a === undefined) continue;
        let n: number;
        if (typeof a === 'boolean' || a === 'TRUE' || a === 'FALSE') { n = (a === true || a === 'TRUE') ? 1 : 0; }
        else if (typeof a === 'string') { n = 0; }
        else { n = Number(a); if (Number.isNaN(n)) continue; }
        if (n > max) max = n;
        found = true;
      }
      return found ? max : 0;
    },
    MINA: async (args) => {
      let min = Infinity; let found = false;
      for (const a of args) {
        if (a === '' || a === null || a === undefined) continue;
        let n: number;
        if (typeof a === 'boolean' || a === 'TRUE' || a === 'FALSE') { n = (a === true || a === 'TRUE') ? 1 : 0; }
        else if (typeof a === 'string') { n = 0; }
        else { n = Number(a); if (Number.isNaN(n)) continue; }
        if (n < min) min = n;
        found = true;
      }
      return found ? min : 0;
    },

    // ---- Correlation / Regression ----
    CORREL: async (args) => {
      const arrX = args[0];
      const arrY = args[1];
      if (!Array.isArray(arrX) || !Array.isArray(arrY)) return ERRORS.VALUE;
      const len = Math.min(arrX.length, arrY.length);
      if (len < 1) return ERRORS.DIV0;
      const xs: number[] = []; const ys: number[] = [];
      for (let i = 0; i < len; i++) {
        const x = Number(arrX[i]); const y = Number(arrY[i]);
        if (!Number.isNaN(x) && !Number.isNaN(y)) { xs.push(x); ys.push(y); }
      }
      const n = xs.length;
      if (n < 1) return ERRORS.DIV0;
      const mx = xs.reduce((a, b) => a + b, 0) / n;
      const my = ys.reduce((a, b) => a + b, 0) / n;
      let sxy = 0, sx2 = 0, sy2 = 0;
      for (let i = 0; i < n; i++) {
        sxy += (xs[i] - mx) * (ys[i] - my);
        sx2 += (xs[i] - mx) ** 2;
        sy2 += (ys[i] - my) ** 2;
      }
      const denom = Math.sqrt(sx2 * sy2);
      return denom === 0 ? ERRORS.DIV0 : sxy / denom;
    },
    'COVARIANCE.P': async (args) => {
      const arrX = args[0]; const arrY = args[1];
      if (!Array.isArray(arrX) || !Array.isArray(arrY)) return ERRORS.VALUE;
      const xs: number[] = []; const ys: number[] = [];
      const len = Math.min(arrX.length, arrY.length);
      for (let i = 0; i < len; i++) {
        const x = Number(arrX[i]); const y = Number(arrY[i]);
        if (!Number.isNaN(x) && !Number.isNaN(y)) { xs.push(x); ys.push(y); }
      }
      const n = xs.length; if (n === 0) return ERRORS.DIV0;
      const mx = xs.reduce((a, b) => a + b, 0) / n;
      const my = ys.reduce((a, b) => a + b, 0) / n;
      let sxy = 0;
      for (let i = 0; i < n; i++) sxy += (xs[i] - mx) * (ys[i] - my);
      return sxy / n;
    },
    'COVARIANCE.S': async (args) => {
      const arrX = args[0]; const arrY = args[1];
      if (!Array.isArray(arrX) || !Array.isArray(arrY)) return ERRORS.VALUE;
      const xs: number[] = []; const ys: number[] = [];
      const len = Math.min(arrX.length, arrY.length);
      for (let i = 0; i < len; i++) {
        const x = Number(arrX[i]); const y = Number(arrY[i]);
        if (!Number.isNaN(x) && !Number.isNaN(y)) { xs.push(x); ys.push(y); }
      }
      const n = xs.length; if (n < 2) return ERRORS.DIV0;
      const mx = xs.reduce((a, b) => a + b, 0) / n;
      const my = ys.reduce((a, b) => a + b, 0) / n;
      let sxy = 0;
      for (let i = 0; i < n; i++) sxy += (xs[i] - mx) * (ys[i] - my);
      return sxy / (n - 1);
    },
    SLOPE: async (args) => {
      const ys = args[0]; const xs = args[1];
      if (!Array.isArray(ys) || !Array.isArray(xs)) return ERRORS.VALUE;
      const len = Math.min(ys.length, xs.length);
      const xn: number[] = []; const yn: number[] = [];
      for (let i = 0; i < len; i++) {
        const x = Number(xs[i]); const y = Number(ys[i]);
        if (!Number.isNaN(x) && !Number.isNaN(y)) { xn.push(x); yn.push(y); }
      }
      const n = xn.length; if (n < 1) return ERRORS.DIV0;
      const mx = xn.reduce((a, b) => a + b, 0) / n;
      const my = yn.reduce((a, b) => a + b, 0) / n;
      let sxy = 0, sx2 = 0;
      for (let i = 0; i < n; i++) { sxy += (xn[i] - mx) * (yn[i] - my); sx2 += (xn[i] - mx) ** 2; }
      return sx2 === 0 ? ERRORS.DIV0 : sxy / sx2;
    },
    INTERCEPT: async (args) => {
      const ys = args[0]; const xs = args[1];
      if (!Array.isArray(ys) || !Array.isArray(xs)) return ERRORS.VALUE;
      const len = Math.min(ys.length, xs.length);
      const xn: number[] = []; const yn: number[] = [];
      for (let i = 0; i < len; i++) {
        const x = Number(xs[i]); const y = Number(ys[i]);
        if (!Number.isNaN(x) && !Number.isNaN(y)) { xn.push(x); yn.push(y); }
      }
      const n = xn.length; if (n < 1) return ERRORS.DIV0;
      const mx = xn.reduce((a, b) => a + b, 0) / n;
      const my = yn.reduce((a, b) => a + b, 0) / n;
      let sxy = 0, sx2 = 0;
      for (let i = 0; i < n; i++) { sxy += (xn[i] - mx) * (yn[i] - my); sx2 += (xn[i] - mx) ** 2; }
      if (sx2 === 0) return ERRORS.DIV0;
      const slope = sxy / sx2;
      return my - slope * mx;
    },
    RSQ: async (args) => {
      const ys = args[0]; const xs = args[1];
      if (!Array.isArray(ys) || !Array.isArray(xs)) return ERRORS.VALUE;
      const len = Math.min(ys.length, xs.length);
      const xn: number[] = []; const yn: number[] = [];
      for (let i = 0; i < len; i++) {
        const x = Number(xs[i]); const y = Number(ys[i]);
        if (!Number.isNaN(x) && !Number.isNaN(y)) { xn.push(x); yn.push(y); }
      }
      const n = xn.length; if (n < 1) return ERRORS.DIV0;
      const mx = xn.reduce((a, b) => a + b, 0) / n;
      const my = yn.reduce((a, b) => a + b, 0) / n;
      let sxy = 0, sx2 = 0, sy2 = 0;
      for (let i = 0; i < n; i++) {
        sxy += (xn[i] - mx) * (yn[i] - my);
        sx2 += (xn[i] - mx) ** 2;
        sy2 += (yn[i] - my) ** 2;
      }
      const denom = sx2 * sy2;
      return denom === 0 ? ERRORS.DIV0 : (sxy * sxy) / denom;
    },
    FORECAST: async (args) => {
      const x = Number(args[0]);
      const ys = args[1]; const xs = args[2];
      if (!Number.isFinite(x)) return ERRORS.VALUE;
      if (!Array.isArray(ys) || !Array.isArray(xs)) return ERRORS.VALUE;
      const len = Math.min(ys.length, xs.length);
      const xn: number[] = []; const yn: number[] = [];
      for (let i = 0; i < len; i++) {
        const xv = Number(xs[i]); const yv = Number(ys[i]);
        if (!Number.isNaN(xv) && !Number.isNaN(yv)) { xn.push(xv); yn.push(yv); }
      }
      const n = xn.length; if (n < 1) return ERRORS.DIV0;
      const mx = xn.reduce((a, b) => a + b, 0) / n;
      const my = yn.reduce((a, b) => a + b, 0) / n;
      let sxy = 0, sx2 = 0;
      for (let i = 0; i < n; i++) { sxy += (xn[i] - mx) * (yn[i] - my); sx2 += (xn[i] - mx) ** 2; }
      if (sx2 === 0) return ERRORS.DIV0;
      const slope = sxy / sx2;
      const intercept = my - slope * mx;
      return intercept + slope * x;
    },
    'FORECAST.LINEAR': async (args) => {
      const x = Number(args[0]);
      const ys = args[1]; const xs = args[2];
      if (!Number.isFinite(x)) return ERRORS.VALUE;
      if (!Array.isArray(ys) || !Array.isArray(xs)) return ERRORS.VALUE;
      const len = Math.min(ys.length, xs.length);
      const xn: number[] = []; const yn: number[] = [];
      for (let i = 0; i < len; i++) {
        const xv = Number(xs[i]); const yv = Number(ys[i]);
        if (!Number.isNaN(xv) && !Number.isNaN(yv)) { xn.push(xv); yn.push(yv); }
      }
      const n = xn.length; if (n < 1) return ERRORS.DIV0;
      const mx = xn.reduce((a, b) => a + b, 0) / n;
      const my = yn.reduce((a, b) => a + b, 0) / n;
      let sxy = 0, sx2 = 0;
      for (let i = 0; i < n; i++) { sxy += (xn[i] - mx) * (yn[i] - my); sx2 += (xn[i] - mx) ** 2; }
      if (sx2 === 0) return ERRORS.DIV0;
      const slope = sxy / sx2;
      const intercept = my - slope * mx;
      return intercept + slope * x;
    },

    // ---- Distribution / Transform ----
    FISHER: async (args) => {
      const x = Number(args[0]);
      if (!Number.isFinite(x) || x <= -1 || x >= 1) return ERRORS.NUM;
      return 0.5 * Math.log((1 + x) / (1 - x));
    },
    FISHERINV: async (args) => {
      const y = Number(args[0]);
      if (!Number.isFinite(y)) return ERRORS.NUM;
      const e2y = Math.exp(2 * y);
      return (e2y - 1) / (e2y + 1);
    },
    GEOMEAN: async (args) => {
      const nums: number[] = [];
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) nums.push(n); }
      if (nums.length === 0) return ERRORS.NUM;
      if (nums.some(n => n <= 0)) return ERRORS.NUM;
      const logSum = nums.reduce((s, v) => s + Math.log(v), 0);
      return Math.exp(logSum / nums.length);
    },
    HARMEAN: async (args) => {
      const nums: number[] = [];
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) nums.push(n); }
      if (nums.length === 0) return ERRORS.NUM;
      if (nums.some(n => n <= 0)) return ERRORS.NUM;
      const recipSum = nums.reduce((s, v) => s + 1 / v, 0);
      return nums.length / recipSum;
    },
    TRIMMEAN: async (args) => {
      const data = args[0];
      const pct = Number(args[1]);
      if (!Array.isArray(data)) return ERRORS.VALUE;
      if (!Number.isFinite(pct) || pct < 0 || pct >= 1) return ERRORS.NUM;
      const nums = data.map((v: any) => Number(v)).filter((n: number) => !Number.isNaN(n));
      if (nums.length === 0) return ERRORS.NUM;
      nums.sort((a: number, b: number) => a - b);
      const trimCount = Math.floor(nums.length * pct / 2);
      const trimmed = nums.slice(trimCount, nums.length - trimCount);
      if (trimmed.length === 0) return ERRORS.NUM;
      return trimmed.reduce((a: number, b: number) => a + b, 0) / trimmed.length;
    },
    DEVSQ: async (args) => {
      const nums: number[] = [];
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) nums.push(n); }
      if (nums.length === 0) return ERRORS.NUM;
      const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
      return nums.reduce((s, v) => s + (v - mean) ** 2, 0);
    },
    AVEDEV: async (args) => {
      const nums: number[] = [];
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) nums.push(n); }
      if (nums.length === 0) return ERRORS.NUM;
      const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
      return nums.reduce((s, v) => s + Math.abs(v - mean), 0) / nums.length;
    },
    SKEW: async (args) => {
      const nums: number[] = [];
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) nums.push(n); }
      const n = nums.length;
      if (n < 3) return ERRORS.DIV0;
      const mean = nums.reduce((a, b) => a + b, 0) / n;
      const s = Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1));
      if (s === 0) return ERRORS.DIV0;
      const m3 = nums.reduce((a, b) => a + ((b - mean) / s) ** 3, 0);
      return (n / ((n - 1) * (n - 2))) * m3;
    },
    KURT: async (args) => {
      const nums: number[] = [];
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) nums.push(n); }
      const n = nums.length;
      if (n < 4) return ERRORS.DIV0;
      const mean = nums.reduce((a, b) => a + b, 0) / n;
      const s = Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1));
      if (s === 0) return ERRORS.DIV0;
      const m4 = nums.reduce((a, b) => a + ((b - mean) / s) ** 4, 0);
      return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * m4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
    },
    STANDARDIZE: async (args) => {
      const x = Number(args[0]);
      const mean = Number(args[1]);
      const stddev = Number(args[2]);
      if (!Number.isFinite(x) || !Number.isFinite(mean) || !Number.isFinite(stddev)) return ERRORS.VALUE;
      if (stddev <= 0) return ERRORS.NUM;
      return (x - mean) / stddev;
    },
    'CONFIDENCE.NORM': async (args) => {
      const alpha = Number(args[0]);
      const stddev = Number(args[1]);
      const n = Number(args[2]);
      if (!Number.isFinite(alpha) || !Number.isFinite(stddev) || !Number.isFinite(n)) return ERRORS.VALUE;
      if (alpha <= 0 || alpha >= 1 || stddev <= 0 || n < 1) return ERRORS.NUM;
      const p = 1 - alpha / 2;
      const t = Math.sqrt(-2 * Math.log(1 - p));
      const z = t - (2.515517 + 0.802853 * t + 0.010328 * t * t) / (1 + 1.432788 * t + 0.189269 * t * t + 0.001308 * t * t * t);
      return z * stddev / Math.sqrt(n);
    },

    // ---- Probability Distributions ----
    'NORM.S.DIST': async (args) => {
      const z = Number(args[0]);
      const cumulative = args[1] !== undefined ? Boolean(args[1]) : true;
      if (!Number.isFinite(z)) return ERRORS.NUM;
      if (!cumulative) {
        return Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
      }
      const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741;
      const a4 = -1.453152027, a5 = 1.061405429, pConst = 0.3275911;
      const sign = z < 0 ? -1 : 1;
      const xVal = Math.abs(z) / Math.sqrt(2);
      const tVal = 1 / (1 + pConst * xVal);
      const y = 1 - (((((a5 * tVal + a4) * tVal) + a3) * tVal + a2) * tVal + a1) * tVal * Math.exp(-xVal * xVal);
      return 0.5 * (1 + sign * y);
    },
    'NORM.DIST': async (args) => {
      const x = Number(args[0]);
      const mean = Number(args[1]);
      const stddev = Number(args[2]);
      const cumulative = args[3] !== undefined ? Boolean(args[3]) : true;
      if (!Number.isFinite(x) || !Number.isFinite(mean) || !Number.isFinite(stddev)) return ERRORS.NUM;
      if (stddev <= 0) return ERRORS.NUM;
      if (!cumulative) {
        const zz = (x - mean) / stddev;
        return Math.exp(-0.5 * zz * zz) / (stddev * Math.sqrt(2 * Math.PI));
      }
      const zz = (x - mean) / stddev;
      const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741;
      const a4 = -1.453152027, a5 = 1.061405429, pConst = 0.3275911;
      const sign = zz < 0 ? -1 : 1;
      const xVal = Math.abs(zz) / Math.sqrt(2);
      const tVal = 1 / (1 + pConst * xVal);
      const y = 1 - (((((a5 * tVal + a4) * tVal) + a3) * tVal + a2) * tVal + a1) * tVal * Math.exp(-xVal * xVal);
      return 0.5 * (1 + sign * y);
    },
    'NORM.S.INV': async (args) => {
      const p = Number(args[0]);
      if (!Number.isFinite(p) || p <= 0 || p >= 1) return ERRORS.NUM;
      const a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00];
      const b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01];
      const c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00];
      const d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00];
      const pLow = 0.02425, pHigh = 1 - pLow;
      let q: number, r: number, result: number;
      if (p < pLow) {
        q = Math.sqrt(-2 * Math.log(p));
        result = (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
      } else if (p <= pHigh) {
        q = p - 0.5; r = q * q;
        result = (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q / (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1);
      } else {
        q = Math.sqrt(-2 * Math.log(1 - p));
        result = -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
      }
      return result;
    },
    'NORM.INV': async (args) => {
      const p = Number(args[0]);
      const mean = Number(args[1]);
      const stddev = Number(args[2]);
      if (!Number.isFinite(p) || !Number.isFinite(mean) || !Number.isFinite(stddev)) return ERRORS.NUM;
      if (p <= 0 || p >= 1 || stddev <= 0) return ERRORS.NUM;
      const a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00];
      const b2 = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01];
      const c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00];
      const d2 = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00];
      const pLow = 0.02425, pHigh = 1 - pLow;
      let q: number, r: number, z: number;
      if (p < pLow) {
        q = Math.sqrt(-2 * Math.log(p));
        z = (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d2[0]*q+d2[1])*q+d2[2])*q+d2[3])*q+1);
      } else if (p <= pHigh) {
        q = p - 0.5; r = q * q;
        z = (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q / (((((b2[0]*r+b2[1])*r+b2[2])*r+b2[3])*r+b2[4])*r+1);
      } else {
        q = Math.sqrt(-2 * Math.log(1 - p));
        z = -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d2[0]*q+d2[1])*q+d2[2])*q+d2[3])*q+1);
      }
      return mean + z * stddev;
    },
    'POISSON.DIST': async (args) => {
      const x = Math.trunc(Number(args[0]));
      const mean = Number(args[1]);
      const cumulative = args[2] !== undefined ? Boolean(args[2]) : true;
      if (!Number.isFinite(x) || !Number.isFinite(mean)) return ERRORS.NUM;
      if (x < 0 || mean < 0) return ERRORS.NUM;
      if (!cumulative) {
        let fact = 1; for (let i = 2; i <= x; i++) fact *= i;
        return Math.exp(-mean) * Math.pow(mean, x) / fact;
      }
      let sum = 0;
      for (let k = 0; k <= x; k++) {
        let fact = 1; for (let i = 2; i <= k; i++) fact *= i;
        sum += Math.pow(mean, k) / fact;
      }
      return Math.exp(-mean) * sum;
    },
    'BINOM.DIST': async (args) => {
      const successes = Math.trunc(Number(args[0]));
      const trials = Math.trunc(Number(args[1]));
      const prob = Number(args[2]);
      const cumulative = args[3] !== undefined ? Boolean(args[3]) : true;
      if (!Number.isFinite(successes) || !Number.isFinite(trials) || !Number.isFinite(prob)) return ERRORS.NUM;
      if (successes < 0 || trials < 0 || successes > trials || prob < 0 || prob > 1) return ERRORS.NUM;
      const comb = (nn2: number, kk2: number): number => {
        if (kk2 > nn2 - kk2) kk2 = nn2 - kk2;
        let r = 1;
        for (let i = 0; i < kk2; i++) r = r * (nn2 - i) / (i + 1);
        return r;
      };
      if (!cumulative) {
        return comb(trials, successes) * Math.pow(prob, successes) * Math.pow(1 - prob, trials - successes);
      }
      let sum = 0;
      for (let k = 0; k <= successes; k++) {
        sum += comb(trials, k) * Math.pow(prob, k) * Math.pow(1 - prob, trials - k);
      }
      return sum;
    },
    'WEIBULL.DIST': async (args) => {
      const x = Number(args[0]);
      const alpha = Number(args[1]);
      const beta = Number(args[2]);
      const cumulative = args[3] !== undefined ? Boolean(args[3]) : true;
      if (!Number.isFinite(x) || !Number.isFinite(alpha) || !Number.isFinite(beta)) return ERRORS.NUM;
      if (x < 0 || alpha <= 0 || beta <= 0) return ERRORS.NUM;
      if (cumulative) {
        return 1 - Math.exp(-Math.pow(x / beta, alpha));
      }
      return (alpha / beta) * Math.pow(x / beta, alpha - 1) * Math.exp(-Math.pow(x / beta, alpha));
    },
    'EXPON.DIST': async (args) => {
      const x = Number(args[0]);
      const lambda = Number(args[1]);
      const cumulative = args[2] !== undefined ? Boolean(args[2]) : true;
      if (!Number.isFinite(x) || !Number.isFinite(lambda)) return ERRORS.NUM;
      if (x < 0 || lambda <= 0) return ERRORS.NUM;
      if (cumulative) return 1 - Math.exp(-lambda * x);
      return lambda * Math.exp(-lambda * x);
    },
    'LOGNORM.DIST': async (args) => {
      const x = Number(args[0]);
      const mean = Number(args[1]);
      const stddev = Number(args[2]);
      const cumulative = args[3] !== undefined ? Boolean(args[3]) : true;
      if (!Number.isFinite(x) || !Number.isFinite(mean) || !Number.isFinite(stddev)) return ERRORS.NUM;
      if (x <= 0 || stddev <= 0) return ERRORS.NUM;
      const zz = (Math.log(x) - mean) / stddev;
      if (!cumulative) {
        return Math.exp(-0.5 * zz * zz) / (x * stddev * Math.sqrt(2 * Math.PI));
      }
      const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741;
      const a4 = -1.453152027, a5 = 1.061405429, pConst = 0.3275911;
      const sign = zz < 0 ? -1 : 1;
      const xVal = Math.abs(zz) / Math.sqrt(2);
      const tVal = 1 / (1 + pConst * xVal);
      const y = 1 - (((((a5 * tVal + a4) * tVal) + a3) * tVal + a2) * tVal + a1) * tVal * Math.exp(-xVal * xVal);
      return 0.5 * (1 + sign * y);
    },
  };
}
