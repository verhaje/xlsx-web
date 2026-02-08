import { ERRORS, isError } from './errors.js';
import { excelSerialToParts, todayToExcelSerial, ymdToExcelSerial } from './date-utils.js';

// Built-in functions registry
export function createBuiltins() {
  return {
    // Math functions
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
      for (let i = 0; i < args.length; i++) { const n = Number(args[i]); if (!Number.isNaN(n)) { if (n < min) min = n; found = true; } }
      return found ? min : 0;
    },
    MAX: async (args) => {
      let max = -Infinity; let found = false;
      for (let i = 0; i < args.length; i++) { const n = Number(args[i]); if (!Number.isNaN(n)) { if (n > max) max = n; found = true; } }
      return found ? max : 0;
    },
    COUNT: async (args) => {
      let count = 0;
      for (const a of args) { const n = Number(a); if (!Number.isNaN(n)) count++; }
      return count;
    },
    COUNTIFS: async (args) => {
      // Expect args as alternating [rangeArray, criteria, rangeArray, criteria, ...]
      try {
        if (!Array.isArray(args) || args.length === 0) return 0;
        if (args.length % 2 !== 0) return ERRORS.VALUE;
        const pairs = [];
        for (let i = 0; i < args.length; i += 2) {
          const range = args[i];
          const crit = args[i + 1];
          if (!Array.isArray(range)) return ERRORS.VALUE;
          pairs.push({ range, crit });
        }
        const len = pairs[0].range.length;
        for (const p of pairs) if (p.range.length !== len) return ERRORS.VALUE;

        const match = (val, crit) => {
          // null/empty handling
          if (crit === null || crit === undefined || crit === '') return val === '' || val === null || val === undefined;
          // criteria string starting with operator
          if (typeof crit === 'string') {
            const m = /^([<>]=?|=)\s*(.*)$/s.exec(crit);
            if (m) {
              const op = m[1];
              const rhs = m[2];
              const rn = Number(rhs);
              const isNum = !Number.isNaN(rn);
              const vn = Number(val);
              if (isNum && !Number.isNaN(vn)) {
                switch (op) {
                  case '>': return vn > rn;
                  case '<': return vn < rn;
                  case '>=': return vn >= rn;
                  case '<=': return vn <= rn;
                  case '=': return vn === rn;
                }
              }
              // string compare fallback
              const vs = String(val);
              switch (op) {
                case '>': return vs > rhs;
                case '<': return vs < rhs;
                case '>=': return vs >= rhs;
                case '<=': return vs <= rhs;
                case '=': return vs === rhs;
              }
            }
            // wildcard support (* and ?)
            if (/[\*\?]/.test(crit)) {
              // escape regex special chars except * and ?, then replace wildcards
              const esc = crit.replace(/[-\\^$+?.()|[\]{}]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
              try {
                const re = new RegExp(`^${esc}$`, 'i');
                return re.test(String(val));
              } catch (e) {
                return false;
              }
            }
            // equality
            return String(val) === crit;
          }
          // numeric or boolean
          return val == crit;
        };

        let count = 0;
        for (let i = 0; i < len; i += 1) {
          let ok = true;
          for (const p of pairs) {
            const v = p.range[i];
            if (!match(v, p.crit)) { ok = false; break; }
          }
          if (ok) count += 1;
        }
        return count;
      } catch (e) {
        return ERRORS.VALUE;
      }
    },
    // SUMIF(criteria_range, criteria, [sum_range])
    SUMIF: async (args) => {
      try {
        const criteriaRange = args[0];
        const criteria = args[1];
        const sumRange = args[2] !== undefined ? args[2] : criteriaRange;

        if (!Array.isArray(criteriaRange)) return ERRORS.VALUE;
        if (!Array.isArray(sumRange)) return ERRORS.VALUE;

        const match = (val, crit) => {
          if (crit === null || crit === undefined || crit === '') return val === '' || val === null || val === undefined;
          if (typeof crit === 'string') {
            const m = /^(<>|[<>]=?|=)\s*(.*)$/s.exec(crit);
            if (m) {
              const op = m[1];
              const rhs = m[2];
              const rn = Number(rhs);
              const isRhsNum = !Number.isNaN(rn) && rhs !== '';
              const vn = Number(val);
              const isValNum = !Number.isNaN(vn) && val !== '' && val !== null;
              // For numeric criteria, only compare numeric values
              if (isRhsNum) {
                if (!isValNum) return false;
                switch (op) {
                  case '>': return vn > rn;
                  case '<': return vn < rn;
                  case '>=': return vn >= rn;
                  case '<=': return vn <= rn;
                  case '=': return vn === rn;
                  case '<>': return vn !== rn;
                }
              }
              // String comparison for non-numeric criteria
              const vs = String(val ?? '');
              switch (op) {
                case '>': return vs > rhs;
                case '<': return vs < rhs;
                case '>=': return vs >= rhs;
                case '<=': return vs <= rhs;
                case '=': return vs.toLowerCase() === rhs.toLowerCase();
                case '<>': return vs.toLowerCase() !== rhs.toLowerCase();
              }
            }
            if (/[\*\?]/.test(crit)) {
              const esc = crit.replace(/[-\\^$+.()|[\]{}]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
              try {
                const re = new RegExp(`^${esc}$`, 'i');
                return re.test(String(val ?? ''));
              } catch (e) { return false; }
            }
            return String(val ?? '').toLowerCase() === crit.toLowerCase();
          }
          return val == crit;
        };

        let sum = 0;
        const len = Math.min(criteriaRange.length, sumRange.length);
        for (let i = 0; i < len; i++) {
          if (match(criteriaRange[i], criteria)) {
            const n = Number(sumRange[i]);
            if (!Number.isNaN(n)) sum += n;
          }
        }
        return sum;
      } catch (e) {
        return ERRORS.VALUE;
      }
    },
    // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
    SUMIFS: async (args) => {
      try {
        if (!Array.isArray(args) || args.length < 3) return ERRORS.VALUE;
        if ((args.length - 1) % 2 !== 0) return ERRORS.VALUE;

        const sumRange = args[0];
        if (!Array.isArray(sumRange)) return ERRORS.VALUE;

        const pairs = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = args[i];
          const crit = args[i + 1];
          if (!Array.isArray(range)) return ERRORS.VALUE;
          pairs.push({ range, crit });
        }

        const len = sumRange.length;
        for (const p of pairs) if (p.range.length !== len) return ERRORS.VALUE;

        const match = (val, crit) => {
          if (crit === null || crit === undefined || crit === '') return val === '' || val === null || val === undefined;
          if (typeof crit === 'string') {
            const m = /^(<>|[<>]=?|=)\s*(.*)$/s.exec(crit);
            if (m) {
              const op = m[1];
              const rhs = m[2];
              const rn = Number(rhs);
              const isRhsNum = !Number.isNaN(rn) && rhs !== '';
              const vn = Number(val);
              const isValNum = !Number.isNaN(vn) && val !== '' && val !== null;
              if (isRhsNum) {
                if (!isValNum) return false;
                switch (op) {
                  case '>': return vn > rn;
                  case '<': return vn < rn;
                  case '>=': return vn >= rn;
                  case '<=': return vn <= rn;
                  case '=': return vn === rn;
                  case '<>': return vn !== rn;
                }
              }
              const vs = String(val ?? '');
              switch (op) {
                case '>': return vs > rhs;
                case '<': return vs < rhs;
                case '>=': return vs >= rhs;
                case '<=': return vs <= rhs;
                case '=': return vs.toLowerCase() === rhs.toLowerCase();
                case '<>': return vs.toLowerCase() !== rhs.toLowerCase();
              }
            }
            if (/[\*\?]/.test(crit)) {
              const esc = crit.replace(/[-\\^$+.()|[\]{}]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
              try {
                const re = new RegExp(`^${esc}$`, 'i');
                return re.test(String(val ?? ''));
              } catch (e) { return false; }
            }
            return String(val ?? '').toLowerCase() === crit.toLowerCase();
          }
          return val == crit;
        };

        let sum = 0;
        for (let i = 0; i < len; i++) {
          let ok = true;
          for (const p of pairs) {
            if (!match(p.range[i], p.crit)) { ok = false; break; }
          }
          if (ok) {
            const n = Number(sumRange[i]);
            if (!Number.isNaN(n)) sum += n;
          }
        }
        return sum;
      } catch (e) {
        return ERRORS.VALUE;
      }
    },
    // AVERAGEIF(criteria_range, criteria, [average_range])
    AVERAGEIF: async (args) => {
      try {
        const criteriaRange = args[0];
        const criteria = args[1];
        const avgRange = args[2] !== undefined ? args[2] : criteriaRange;

        if (!Array.isArray(criteriaRange)) return ERRORS.VALUE;
        if (!Array.isArray(avgRange)) return ERRORS.VALUE;

        const match = (val, crit) => {
          if (crit === null || crit === undefined || crit === '') return val === '' || val === null || val === undefined;
          if (typeof crit === 'string') {
            const m = /^(<>|[<>]=?|=)\s*(.*)$/s.exec(crit);
            if (m) {
              const op = m[1];
              const rhs = m[2];
              const rn = Number(rhs);
              const isRhsNum = !Number.isNaN(rn) && rhs !== '';
              const vn = Number(val);
              const isValNum = !Number.isNaN(vn) && val !== '' && val !== null;
              if (isRhsNum) {
                if (!isValNum) return false;
                switch (op) {
                  case '>': return vn > rn;
                  case '<': return vn < rn;
                  case '>=': return vn >= rn;
                  case '<=': return vn <= rn;
                  case '=': return vn === rn;
                  case '<>': return vn !== rn;
                }
              }
              const vs = String(val ?? '');
              switch (op) {
                case '>': return vs > rhs;
                case '<': return vs < rhs;
                case '>=': return vs >= rhs;
                case '<=': return vs <= rhs;
                case '=': return vs.toLowerCase() === rhs.toLowerCase();
                case '<>': return vs.toLowerCase() !== rhs.toLowerCase();
              }
            }
            if (/[\*\?]/.test(crit)) {
              const esc = crit.replace(/[-\\^$+.()|[\]{}]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
              try {
                const re = new RegExp(`^${esc}$`, 'i');
                return re.test(String(val ?? ''));
              } catch (e) { return false; }
            }
            return String(val ?? '').toLowerCase() === crit.toLowerCase();
          }
          return val == crit;
        };

        let sum = 0;
        let count = 0;
        const len = Math.min(criteriaRange.length, avgRange.length);
        for (let i = 0; i < len; i++) {
          if (match(criteriaRange[i], criteria)) {
            const n = Number(avgRange[i]);
            if (!Number.isNaN(n)) { sum += n; count++; }
          }
        }
        return count === 0 ? ERRORS.DIV0 : sum / count;
      } catch (e) {
        return ERRORS.VALUE;
      }
    },
    // AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
    AVERAGEIFS: async (args) => {
      try {
        if (!Array.isArray(args) || args.length < 3) return ERRORS.VALUE;
        if ((args.length - 1) % 2 !== 0) return ERRORS.VALUE;

        const avgRange = args[0];
        if (!Array.isArray(avgRange)) return ERRORS.VALUE;

        const pairs = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = args[i];
          const crit = args[i + 1];
          if (!Array.isArray(range)) return ERRORS.VALUE;
          pairs.push({ range, crit });
        }

        const len = avgRange.length;
        for (const p of pairs) if (p.range.length !== len) return ERRORS.VALUE;

        const match = (val, crit) => {
          if (crit === null || crit === undefined || crit === '') return val === '' || val === null || val === undefined;
          if (typeof crit === 'string') {
            const m = /^(<>|[<>]=?|=)\s*(.*)$/s.exec(crit);
            if (m) {
              const op = m[1];
              const rhs = m[2];
              const rn = Number(rhs);
              const isRhsNum = !Number.isNaN(rn) && rhs !== '';
              const vn = Number(val);
              const isValNum = !Number.isNaN(vn) && val !== '' && val !== null;
              if (isRhsNum) {
                if (!isValNum) return false;
                switch (op) {
                  case '>': return vn > rn;
                  case '<': return vn < rn;
                  case '>=': return vn >= rn;
                  case '<=': return vn <= rn;
                  case '=': return vn === rn;
                  case '<>': return vn !== rn;
                }
              }
              const vs = String(val ?? '');
              switch (op) {
                case '>': return vs > rhs;
                case '<': return vs < rhs;
                case '>=': return vs >= rhs;
                case '<=': return vs <= rhs;
                case '=': return vs.toLowerCase() === rhs.toLowerCase();
                case '<>': return vs.toLowerCase() !== rhs.toLowerCase();
              }
            }
            if (/[\*\?]/.test(crit)) {
              const esc = crit.replace(/[-\\^$+.()|[\]{}]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
              try {
                const re = new RegExp(`^${esc}$`, 'i');
                return re.test(String(val ?? ''));
              } catch (e) { return false; }
            }
            return String(val ?? '').toLowerCase() === crit.toLowerCase();
          }
          return val == crit;
        };

        let sum = 0;
        let count = 0;
        for (let i = 0; i < len; i++) {
          let ok = true;
          for (const p of pairs) {
            if (!match(p.range[i], p.crit)) { ok = false; break; }
          }
          if (ok) {
            const n = Number(avgRange[i]);
            if (!Number.isNaN(n)) { sum += n; count++; }
          }
        }
        return count === 0 ? ERRORS.DIV0 : sum / count;
      } catch (e) {
        return ERRORS.VALUE;
      }
    },
    ABS: async (args) => Math.abs(Number(args[0]) || 0),
    ROUND: async (args) => {
      const num = Number(args[0]) || 0;
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

    // Logical functions
    IF: async (args) => {
      const cond = args[0];
      const truthy = cond && cond !== 0 && cond !== '' && cond !== false && cond !== 'FALSE';
      return truthy ? (args[1] !== undefined ? args[1] : true) : (args[2] !== undefined ? args[2] : false);
    },
    AND: async (args) => {
      for (const a of args) { if (!a || a === 0 || a === '' || a === false || a === 'FALSE') return false; }
      return true;
    },
    OR: async (args) => {
      for (const a of args) { if (a && a !== 0 && a !== '' && a !== false && a !== 'FALSE') return true; }
      return false;
    },
    NOT: async (args) => {
      const v = args[0];
      return !(v && v !== 0 && v !== '' && v !== false && v !== 'FALSE');
    },
    TRUE: async () => true,
    FALSE: async () => false,
    IFERROR: async (args) => isError(args[0]) ? (args[1] !== undefined ? args[1] : '') : args[0],
    ISERROR: async (args) => isError(args[0]),
    ISBLANK: async (args) => args[0] === '' || args[0] === null || args[0] === undefined,
    ISNUMBER: async (args) => typeof args[0] === 'number' && !Number.isNaN(args[0]),
    ISTEXT: async (args) => typeof args[0] === 'string',

    // Date/time functions (Excel serial numbers with 1900 leap-year bug)
    DATE: async (args) => {
      if (args.length < 3) return ERRORS.VALUE;
      const year = Number(args[0]);
      const month = Number(args[1]);
      const day = Number(args[2]);
      if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return ERRORS.VALUE;
      const serial = ymdToExcelSerial(year, month, day);
      return Number.isFinite(serial) ? serial : ERRORS.VALUE;
    },
    YEAR: async (args) => {
      const v = args[0];
      if (v instanceof Date) return v.getUTCFullYear();
      const n = Number(v);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      const parts = excelSerialToParts(n);
      return parts ? parts.year : ERRORS.VALUE;
    },
    MONTH: async (args) => {
      const v = args[0];
      if (v instanceof Date) return v.getUTCMonth() + 1;
      const n = Number(v);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      const parts = excelSerialToParts(n);
      return parts ? parts.month : ERRORS.VALUE;
    },
    DAY: async (args) => {
      const v = args[0];
      if (v instanceof Date) return v.getUTCDate();
      const n = Number(v);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      const parts = excelSerialToParts(n);
      return parts ? parts.day : ERRORS.VALUE;
    },
    TODAY: async () => todayToExcelSerial(),

    // Text functions
    CONCAT: async (args) => args.map(a => String(a ?? '')).join(''),
    CONCATENATE: async (args) => args.map(a => String(a ?? '')).join(''),
    LEFT: async (args) => String(args[0] ?? '').slice(0, Number(args[1]) || 1),
    RIGHT: async (args) => { const s = String(args[0] ?? ''); const n = Number(args[1]) || 1; return s.slice(-n); },
    MID: async (args) => { const s = String(args[0] ?? ''); const start = (Number(args[1]) || 1) - 1; const len = Number(args[2]) || 1; return s.slice(start, start + len); },
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

    // Statistical functions
    MEDIAN: async (args) => {
      // collect numeric values avoiding extra allocations
      const nums = [];
      for (let i = 0; i < args.length; i++) { const n = Number(args[i]); if (!Number.isNaN(n)) nums.push(n); }
      if (nums.length === 0) return ERRORS.NUM;
      nums.sort((a, b) => a - b);
      const mid = nums.length >> 1;
      return (nums.length & 1) ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
    },
    STDEV: async (args) => {
      const nums = args.map(a => Number(a)).filter(n => !Number.isNaN(n));
      const n = nums.length;
      if (n < 2) return ERRORS.DIV0;
      const mean = nums.reduce((sum, v) => sum + v, 0) / n;
      let sumSq = 0;
      for (const v of nums) sumSq += Math.pow(v - mean, 2);
      const variance = sumSq / (n - 1);
      return Math.sqrt(variance);
    },

    // Lookup functions - full implementation
    // MATCH(lookup_value, lookup_array, [match_type])
    // match_type: 1 (default) = less than or equal (array must be ascending)
    //             0 = exact match
    //            -1 = greater than or equal (array must be descending)
    MATCH: async (args, meta) => {
      const lookupValue = args[0];
      const lookupArray = args[1]; // expected to be array of values
      const matchType = args[2] !== undefined ? Number(args[2]) : 1;

      if (!Array.isArray(lookupArray) || lookupArray.length === 0) {
        return ERRORS.NA;
      }

      // Get row/col info from meta if available
      let rows = 1, cols = lookupArray.length;
      if (meta && meta[1]) {
        rows = meta[1].rows || 1;
        cols = meta[1].cols || lookupArray.length;
      }

      // For MATCH, array must be 1-dimensional (single row or column)
      if (rows > 1 && cols > 1) {
        return ERRORS.NA; // MATCH requires a 1-D array
      }

      const arr = lookupArray;
      const isNumLookup = typeof lookupValue === 'number' || !Number.isNaN(Number(lookupValue));
      const numLookup = isNumLookup ? Number(lookupValue) : null;

      // Helper for value comparison
      const compareValues = (a, b) => {
        const aNum = Number(a);
        const bNum = Number(b);
        const aIsNum = !Number.isNaN(aNum) && a !== '' && a !== null;
        const bIsNum = !Number.isNaN(bNum) && b !== '' && b !== null;

        if (aIsNum && bIsNum) {
          return aNum - bNum;
        }
        // Strings compare case-insensitively
        const aStr = String(a ?? '').toLowerCase();
        const bStr = String(b ?? '').toLowerCase();
        return aStr.localeCompare(bStr);
      };

      // Exact match (match_type = 0)
      if (matchType === 0) {
        for (let i = 0; i < arr.length; i++) {
          const val = arr[i];
          if (isNumLookup) {
            const valNum = Number(val);
            if (!Number.isNaN(valNum) && valNum === numLookup) {
              return i + 1; // 1-based index
            }
          } else {
            // Case-insensitive string match
            if (String(val ?? '').toLowerCase() === String(lookupValue ?? '').toLowerCase()) {
              return i + 1;
            }
          }
        }
        return ERRORS.NA;
      }

      // Approximate match (match_type = 1 or -1)
      // match_type = 1: find largest value <= lookup_value (ascending order assumed)
      // match_type = -1: find smallest value >= lookup_value (descending order assumed)
      let bestIdx = -1;
      let bestVal = null;

      for (let i = 0; i < arr.length; i++) {
        const val = arr[i];
        if (val === '' || val === null || val === undefined) continue;

        const cmp = compareValues(val, lookupValue);

        if (matchType >= 1) {
          // Looking for largest value <= lookupValue
          if (cmp <= 0) {
            if (bestIdx === -1 || compareValues(val, bestVal) > 0) {
              bestIdx = i;
              bestVal = val;
            }
          }
        } else {
          // match_type = -1: Looking for smallest value >= lookupValue
          if (cmp >= 0) {
            if (bestIdx === -1 || compareValues(val, bestVal) < 0) {
              bestIdx = i;
              bestVal = val;
            }
          }
        }
      }

      if (bestIdx === -1) return ERRORS.NA;
      return bestIdx + 1; // 1-based index
    },

    // INDEX(array, row_num, [col_num])
    INDEX: async (args, meta) => {
      const array = args[0];
      const rowNum = args[1] !== undefined ? Number(args[1]) : 1;
      const colNum = args[2] !== undefined ? Number(args[2]) : 1;

      if (!Array.isArray(array) || array.length === 0) {
        // Single value passed
        if (rowNum === 1 && colNum === 1) return array;
        return ERRORS.REF;
      }

      // Get dimensions from meta
      let rows = 1, cols = array.length;
      if (meta && meta[0]) {
        rows = meta[0].rows || 1;
        cols = meta[0].cols || array.length;
      }

      // Validate bounds
      if (rowNum < 0 || colNum < 0) return ERRORS.VALUE;
      if (rowNum > rows || colNum > cols) return ERRORS.REF;

      // Handle row_num = 0 or col_num = 0 (return entire row/column - simplified: just return first match)
      if (rowNum === 0 && colNum === 0) return ERRORS.REF;

      // Calculate 0-based index in flattened array
      // Array is stored row-major: [row1col1, row1col2, ..., row2col1, row2col2, ...]
      const r = rowNum === 0 ? 1 : rowNum;
      const c = colNum === 0 ? 1 : colNum;

      const idx = (r - 1) * cols + (c - 1);
      if (idx < 0 || idx >= array.length) return ERRORS.REF;

      const result = array[idx];
      return result !== undefined ? result : '';
    },

    // VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
    VLOOKUP: async (args, meta) => {
      const lookupValue = args[0];
      const tableArray = args[1];
      const colIndexNum = Number(args[2]);
      // range_lookup: TRUE/1 (default) = approximate match, FALSE/0 = exact match
      const rangeLookup = args[3] === undefined ? true :
        (args[3] === false || args[3] === 0 || args[3] === 'FALSE' || String(args[3]).toUpperCase() === 'FALSE' ? false : true);

      if (!Array.isArray(tableArray) || tableArray.length === 0) {
        return ERRORS.NA;
      }

      // Get dimensions from meta
      let rows = 1, cols = tableArray.length;
      if (meta && meta[1]) {
        rows = meta[1].rows || 1;
        cols = meta[1].cols || tableArray.length;
      }

      // Validate col_index_num
      if (colIndexNum < 1) return ERRORS.VALUE;
      if (colIndexNum > cols) return ERRORS.REF;

      // Extract first column for searching
      const firstCol = [];
      for (let r = 0; r < rows; r++) {
        firstCol.push(tableArray[r * cols]);
      }

      const isNumLookup = typeof lookupValue === 'number' || !Number.isNaN(Number(lookupValue));
      const numLookup = isNumLookup ? Number(lookupValue) : null;

      // Helper for value comparison
      const compareValues = (a, b) => {
        const aNum = Number(a);
        const bNum = Number(b);
        const aIsNum = !Number.isNaN(aNum) && a !== '' && a !== null;
        const bIsNum = !Number.isNaN(bNum) && b !== '' && b !== null;

        if (aIsNum && bIsNum) {
          return aNum - bNum;
        }
        const aStr = String(a ?? '').toLowerCase();
        const bStr = String(b ?? '').toLowerCase();
        return aStr.localeCompare(bStr);
      };

      let foundRow = -1;

      if (!rangeLookup) {
        // Exact match
        for (let i = 0; i < firstCol.length; i++) {
          const val = firstCol[i];
          if (isNumLookup) {
            const valNum = Number(val);
            if (!Number.isNaN(valNum) && valNum === numLookup) {
              foundRow = i;
              break;
            }
          } else {
            if (String(val ?? '').toLowerCase() === String(lookupValue ?? '').toLowerCase()) {
              foundRow = i;
              break;
            }
          }
        }
      } else {
        // Approximate match - find largest value <= lookupValue
        let bestRow = -1;
        let bestVal = null;
        for (let i = 0; i < firstCol.length; i++) {
          const val = firstCol[i];
          if (val === '' || val === null || val === undefined) continue;
          const cmp = compareValues(val, lookupValue);
          if (cmp <= 0) {
            if (bestRow === -1 || compareValues(val, bestVal) > 0) {
              bestRow = i;
              bestVal = val;
            }
          }
        }
        foundRow = bestRow;
      }

      if (foundRow === -1) return ERRORS.NA;

      // Return value from the specified column
      const resultIdx = foundRow * cols + (colIndexNum - 1);
      if (resultIdx < 0 || resultIdx >= tableArray.length) return ERRORS.REF;

      const result = tableArray[resultIdx];
      return result !== undefined ? result : '';
    },

    // HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
    HLOOKUP: async (args, meta) => {
      const lookupValue = args[0];
      const tableArray = args[1];
      const rowIndexNum = Number(args[2]);
      const rangeLookup = args[3] === undefined ? true :
        (args[3] === false || args[3] === 0 || args[3] === 'FALSE' || String(args[3]).toUpperCase() === 'FALSE' ? false : true);

      if (!Array.isArray(tableArray) || tableArray.length === 0) {
        return ERRORS.NA;
      }

      // Get dimensions from meta
      let rows = 1, cols = tableArray.length;
      if (meta && meta[1]) {
        rows = meta[1].rows || 1;
        cols = meta[1].cols || tableArray.length;
      }

      // Validate row_index_num
      if (rowIndexNum < 1) return ERRORS.VALUE;
      if (rowIndexNum > rows) return ERRORS.REF;

      // Extract first row for searching
      const firstRow = [];
      for (let c = 0; c < cols; c++) {
        firstRow.push(tableArray[c]);
      }

      const isNumLookup = typeof lookupValue === 'number' || !Number.isNaN(Number(lookupValue));
      const numLookup = isNumLookup ? Number(lookupValue) : null;

      // Helper for value comparison
      const compareValues = (a, b) => {
        const aNum = Number(a);
        const bNum = Number(b);
        const aIsNum = !Number.isNaN(aNum) && a !== '' && a !== null;
        const bIsNum = !Number.isNaN(bNum) && b !== '' && b !== null;

        if (aIsNum && bIsNum) {
          return aNum - bNum;
        }
        const aStr = String(a ?? '').toLowerCase();
        const bStr = String(b ?? '').toLowerCase();
        return aStr.localeCompare(bStr);
      };

      let foundCol = -1;

      if (!rangeLookup) {
        // Exact match
        for (let i = 0; i < firstRow.length; i++) {
          const val = firstRow[i];
          if (isNumLookup) {
            const valNum = Number(val);
            if (!Number.isNaN(valNum) && valNum === numLookup) {
              foundCol = i;
              break;
            }
          } else {
            if (String(val ?? '').toLowerCase() === String(lookupValue ?? '').toLowerCase()) {
              foundCol = i;
              break;
            }
          }
        }
      } else {
        // Approximate match - find largest value <= lookupValue
        let bestCol = -1;
        let bestVal = null;
        for (let i = 0; i < firstRow.length; i++) {
          const val = firstRow[i];
          if (val === '' || val === null || val === undefined) continue;
          const cmp = compareValues(val, lookupValue);
          if (cmp <= 0) {
            if (bestCol === -1 || compareValues(val, bestVal) > 0) {
              bestCol = i;
              bestVal = val;
            }
          }
        }
        foundCol = bestCol;
      }

      if (foundCol === -1) return ERRORS.NA;

      // Return value from the specified row
      const resultIdx = (rowIndexNum - 1) * cols + foundCol;
      if (resultIdx < 0 || resultIdx >= tableArray.length) return ERRORS.REF;

      const result = tableArray[resultIdx];
      return result !== undefined ? result : '';
    },
  };
}
