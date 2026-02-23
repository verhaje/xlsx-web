// lookup.ts - Lookup, reference, and dynamic array functions

import { ERRORS } from '../Errors';
import { compareValues } from './helpers';
import type { BuiltinFunction } from '../../types';

export function getLookupFunctions(): Record<string, BuiltinFunction> {
  return {
    MATCH: async (args, meta) => {
      const lookupValue = args[0];
      const lookupArray = args[1];
      const matchType = args[2] !== undefined ? Number(args[2]) : 1;
      if (!Array.isArray(lookupArray) || lookupArray.length === 0) return ERRORS.NA;

      let rows = 1, cols = lookupArray.length;
      if (meta && meta[1]) { rows = meta[1].rows || 1; cols = meta[1].cols || lookupArray.length; }
      if (rows > 1 && cols > 1) return ERRORS.NA;

      const arr = lookupArray;
      const isNumLookup = typeof lookupValue === 'number' || !Number.isNaN(Number(lookupValue));
      const numLookup = isNumLookup ? Number(lookupValue) : null;

      if (matchType === 0) {
        for (let i = 0; i < arr.length; i++) {
          const val = arr[i];
          if (isNumLookup) {
            const valNum = Number(val);
            if (!Number.isNaN(valNum) && valNum === numLookup) return i + 1;
          } else {
            if (String(val ?? '').toLowerCase() === String(lookupValue ?? '').toLowerCase()) return i + 1;
          }
        }
        return ERRORS.NA;
      }

      let bestIdx = -1;
      let bestVal: any = null;
      for (let i = 0; i < arr.length; i++) {
        const val = arr[i];
        if (val === '' || val === null || val === undefined) continue;
        const cmp = compareValues(val, lookupValue);
        if (matchType >= 1) {
          if (cmp <= 0 && (bestIdx === -1 || compareValues(val, bestVal) > 0)) { bestIdx = i; bestVal = val; }
        } else {
          if (cmp >= 0 && (bestIdx === -1 || compareValues(val, bestVal) < 0)) { bestIdx = i; bestVal = val; }
        }
      }
      return bestIdx === -1 ? ERRORS.NA : bestIdx + 1;
    },

    INDEX: async (args, meta) => {
      const array = args[0];
      const rowNum = args[1] !== undefined ? Number(args[1]) : 1;
      const colNum = args[2] !== undefined ? Number(args[2]) : 1;
      if (!Array.isArray(array) || array.length === 0) {
        if (rowNum === 1 && colNum === 1) return array;
        return ERRORS.REF;
      }
      let rows = 1, cols = array.length;
      if (meta && meta[0]) { rows = meta[0].rows || 1; cols = meta[0].cols || array.length; }
      if (rowNum < 0 || colNum < 0) return ERRORS.VALUE;
      if (rowNum > rows || colNum > cols) return ERRORS.REF;
      if (rowNum === 0 && colNum === 0) return ERRORS.REF;
      const r = rowNum === 0 ? 1 : rowNum;
      const c = colNum === 0 ? 1 : colNum;
      const idx = (r - 1) * cols + (c - 1);
      if (idx < 0 || idx >= array.length) return ERRORS.REF;
      return array[idx] !== undefined ? array[idx] : '';
    },

    VLOOKUP: async (args, meta) => {
      const lookupValue = args[0];
      const tableArray = args[1];
      const colIndexNum = Number(args[2]);
      const rangeLookup = args[3] === undefined ? true :
        (args[3] === false || args[3] === 0 || args[3] === 'FALSE' || String(args[3]).toUpperCase() === 'FALSE' ? false : true);
      if (!Array.isArray(tableArray) || tableArray.length === 0) return ERRORS.NA;

      let rows = 1, cols = tableArray.length;
      if (meta && meta[1]) { rows = meta[1].rows || 1; cols = meta[1].cols || tableArray.length; }
      if (colIndexNum < 1) return ERRORS.VALUE;
      if (colIndexNum > cols) return ERRORS.REF;

      const firstCol: any[] = [];
      for (let r = 0; r < rows; r++) firstCol.push(tableArray[r * cols]);

      const isNumLookup = typeof lookupValue === 'number' || !Number.isNaN(Number(lookupValue));
      const numLookup = isNumLookup ? Number(lookupValue) : null;
      let foundRow = -1;

      if (!rangeLookup) {
        for (let i = 0; i < firstCol.length; i++) {
          const val = firstCol[i];
          if (isNumLookup) {
            const valNum = Number(val);
            if (!Number.isNaN(valNum) && valNum === numLookup) { foundRow = i; break; }
          } else {
            if (String(val ?? '').toLowerCase() === String(lookupValue ?? '').toLowerCase()) { foundRow = i; break; }
          }
        }
      } else {
        let bestRow = -1; let bestVal: any = null;
        for (let i = 0; i < firstCol.length; i++) {
          const val = firstCol[i];
          if (val === '' || val === null || val === undefined) continue;
          const cmp = compareValues(val, lookupValue);
          if (cmp <= 0 && (bestRow === -1 || compareValues(val, bestVal) > 0)) { bestRow = i; bestVal = val; }
        }
        foundRow = bestRow;
      }

      if (foundRow === -1) return ERRORS.NA;
      const resultIdx = foundRow * cols + (colIndexNum - 1);
      if (resultIdx < 0 || resultIdx >= tableArray.length) return ERRORS.REF;
      return tableArray[resultIdx] !== undefined ? tableArray[resultIdx] : '';
    },

    HLOOKUP: async (args, meta) => {
      const lookupValue = args[0];
      const tableArray = args[1];
      const rowIndexNum = Number(args[2]);
      const rangeLookup = args[3] === undefined ? true :
        (args[3] === false || args[3] === 0 || args[3] === 'FALSE' || String(args[3]).toUpperCase() === 'FALSE' ? false : true);
      if (!Array.isArray(tableArray) || tableArray.length === 0) return ERRORS.NA;

      let rows = 1, cols = tableArray.length;
      if (meta && meta[1]) { rows = meta[1].rows || 1; cols = meta[1].cols || tableArray.length; }
      if (rowIndexNum < 1) return ERRORS.VALUE;
      if (rowIndexNum > rows) return ERRORS.REF;

      const firstRow: any[] = [];
      for (let c = 0; c < cols; c++) firstRow.push(tableArray[c]);

      const isNumLookup = typeof lookupValue === 'number' || !Number.isNaN(Number(lookupValue));
      const numLookup = isNumLookup ? Number(lookupValue) : null;
      let foundCol = -1;

      if (!rangeLookup) {
        for (let i = 0; i < firstRow.length; i++) {
          const val = firstRow[i];
          if (isNumLookup) {
            const valNum = Number(val);
            if (!Number.isNaN(valNum) && valNum === numLookup) { foundCol = i; break; }
          } else {
            if (String(val ?? '').toLowerCase() === String(lookupValue ?? '').toLowerCase()) { foundCol = i; break; }
          }
        }
      } else {
        let bestCol = -1; let bestVal: any = null;
        for (let i = 0; i < firstRow.length; i++) {
          const val = firstRow[i];
          if (val === '' || val === null || val === undefined) continue;
          const cmp = compareValues(val, lookupValue);
          if (cmp <= 0 && (bestCol === -1 || compareValues(val, bestVal) > 0)) { bestCol = i; bestVal = val; }
        }
        foundCol = bestCol;
      }

      if (foundCol === -1) return ERRORS.NA;
      const resultIdx = (rowIndexNum - 1) * cols + foundCol;
      if (resultIdx < 0 || resultIdx >= tableArray.length) return ERRORS.REF;
      return tableArray[resultIdx] !== undefined ? tableArray[resultIdx] : '';
    },

    XLOOKUP: async (args, meta) => {
      const lookupValue = args[0];
      const lookupArray = args[1];
      const returnArray = args[2];
      const ifNotFound = args[3] !== undefined ? args[3] : ERRORS.NA;
      const matchMode = args[4] !== undefined ? Number(args[4]) : 0;
      const searchMode = args[5] !== undefined ? Number(args[5]) : 1;
      if (!Array.isArray(lookupArray) || !Array.isArray(returnArray)) return ERRORS.VALUE;
      if (lookupArray.length === 0) return ERRORS.NA;

      const isNumLookup = typeof lookupValue === 'number' || (!Number.isNaN(Number(lookupValue)) && lookupValue !== '' && lookupValue !== null);
      const numLookup = isNumLookup ? Number(lookupValue) : null;
      let foundIdx = -1;

      const start = (searchMode === -1 || searchMode === -2) ? lookupArray.length - 1 : 0;
      const end = (searchMode === -1 || searchMode === -2) ? -1 : lookupArray.length;
      const step = (searchMode === -1 || searchMode === -2) ? -1 : 1;

      if (matchMode === 0) {
        for (let i = start; i !== end; i += step) {
          const val = lookupArray[i];
          if (isNumLookup) {
            const vn = Number(val);
            if (!Number.isNaN(vn) && vn === numLookup) { foundIdx = i; break; }
          } else {
            if (String(val ?? '').toLowerCase() === String(lookupValue ?? '').toLowerCase()) { foundIdx = i; break; }
          }
        }
      } else if (matchMode === -1) {
        let bestIdx = -1; let bestVal: any = null;
        for (let i = 0; i < lookupArray.length; i++) {
          const val = lookupArray[i];
          if (val === '' || val === null || val === undefined) continue;
          const cmp = compareValues(val, lookupValue);
          if (cmp === 0) { bestIdx = i; break; }
          if (cmp < 0 && (bestIdx === -1 || compareValues(val, bestVal) > 0)) { bestIdx = i; bestVal = val; }
        }
        foundIdx = bestIdx;
      } else if (matchMode === 1) {
        let bestIdx = -1; let bestVal: any = null;
        for (let i = 0; i < lookupArray.length; i++) {
          const val = lookupArray[i];
          if (val === '' || val === null || val === undefined) continue;
          const cmp = compareValues(val, lookupValue);
          if (cmp === 0) { bestIdx = i; break; }
          if (cmp > 0 && (bestIdx === -1 || compareValues(val, bestVal) < 0)) { bestIdx = i; bestVal = val; }
        }
        foundIdx = bestIdx;
      } else if (matchMode === 2) {
        const pattern = String(lookupValue ?? '').replace(/[-\\^$+.()|[\]{}]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
        try {
          const re = new RegExp(`^${pattern}$`, 'i');
          for (let i = start; i !== end; i += step) {
            if (re.test(String(lookupArray[i] ?? ''))) { foundIdx = i; break; }
          }
        } catch { return ERRORS.VALUE; }
      }

      if (foundIdx === -1) return ifNotFound;
      if (foundIdx < returnArray.length) return returnArray[foundIdx];
      return ERRORS.NA;
    },

    XMATCH: async (args) => {
      const lookupValue = args[0];
      const lookupArray = args[1];
      const matchMode = args[2] !== undefined ? Number(args[2]) : 0;
      const searchMode = args[3] !== undefined ? Number(args[3]) : 1;
      if (!Array.isArray(lookupArray) || lookupArray.length === 0) return ERRORS.NA;

      const isNumLookup = typeof lookupValue === 'number' || (!Number.isNaN(Number(lookupValue)) && lookupValue !== '' && lookupValue !== null);
      const numLookup = isNumLookup ? Number(lookupValue) : null;

      const start = (searchMode === -1 || searchMode === -2) ? lookupArray.length - 1 : 0;
      const end = (searchMode === -1 || searchMode === -2) ? -1 : lookupArray.length;
      const step = (searchMode === -1 || searchMode === -2) ? -1 : 1;

      if (matchMode === 0) {
        for (let i = start; i !== end; i += step) {
          const val = lookupArray[i];
          if (isNumLookup) {
            const vn = Number(val);
            if (!Number.isNaN(vn) && vn === numLookup) return i + 1;
          } else {
            if (String(val ?? '').toLowerCase() === String(lookupValue ?? '').toLowerCase()) return i + 1;
          }
        }
        return ERRORS.NA;
      } else if (matchMode === -1) {
        let bestIdx = -1; let bestVal: any = null;
        for (let i = 0; i < lookupArray.length; i++) {
          const val = lookupArray[i];
          if (val === '' || val === null || val === undefined) continue;
          const cmp = compareValues(val, lookupValue);
          if (cmp === 0) return i + 1;
          if (cmp < 0 && (bestIdx === -1 || compareValues(val, bestVal) > 0)) { bestIdx = i; bestVal = val; }
        }
        return bestIdx === -1 ? ERRORS.NA : bestIdx + 1;
      } else if (matchMode === 1) {
        let bestIdx = -1; let bestVal: any = null;
        for (let i = 0; i < lookupArray.length; i++) {
          const val = lookupArray[i];
          if (val === '' || val === null || val === undefined) continue;
          const cmp = compareValues(val, lookupValue);
          if (cmp === 0) return i + 1;
          if (cmp > 0 && (bestIdx === -1 || compareValues(val, bestVal) < 0)) { bestIdx = i; bestVal = val; }
        }
        return bestIdx === -1 ? ERRORS.NA : bestIdx + 1;
      }
      return ERRORS.NA;
    },

    LOOKUP: async (args) => {
      const lookupValue = args[0];
      const lookupVector = args[1];
      const resultVector = args[2] !== undefined ? args[2] : lookupVector;
      if (!Array.isArray(lookupVector)) return ERRORS.NA;
      if (!Array.isArray(resultVector)) return ERRORS.NA;

      let bestIdx = -1; let bestVal: any = null;
      for (let i = 0; i < lookupVector.length; i++) {
        const val = lookupVector[i];
        if (val === '' || val === null || val === undefined) continue;
        const cmp = compareValues(val, lookupValue);
        if (cmp <= 0 && (bestIdx === -1 || compareValues(val, bestVal) > 0)) { bestIdx = i; bestVal = val; }
      }
      if (bestIdx === -1) return ERRORS.NA;
      return bestIdx < resultVector.length ? resultVector[bestIdx] : ERRORS.NA;
    },

    ROW: async (args, meta) => {
      if (meta && meta[0] && meta[0].row !== undefined) return meta[0].row;
      if (args.length === 0) return 1;
      return ERRORS.VALUE;
    },
    ROWS: async (args, meta) => {
      if (meta && meta[0] && meta[0].rows !== undefined) return meta[0].rows;
      if (Array.isArray(args[0])) return args[0].length;
      return 1;
    },
    COLUMN: async (args, meta) => {
      if (meta && meta[0] && meta[0].col !== undefined) return meta[0].col;
      if (args.length === 0) return 1;
      return ERRORS.VALUE;
    },
    COLUMNS: async (args, meta) => {
      if (meta && meta[0] && meta[0].cols !== undefined) return meta[0].cols;
      if (Array.isArray(args[0])) return args[0].length;
      return 1;
    },
    CHOOSE: async (args) => {
      const idx = Math.trunc(Number(args[0]));
      if (idx < 1 || idx >= args.length) return ERRORS.VALUE;
      return args[idx];
    },
    ADDRESS: async (args) => {
      const row = Math.trunc(Number(args[0]));
      const col = Math.trunc(Number(args[1]));
      const absType = args[2] !== undefined ? Number(args[2]) : 1;
      const a1 = args[3] !== undefined ? args[3] : true;
      const sheet = args[4] !== undefined ? String(args[4]) : '';
      if (row < 1 || col < 1) return ERRORS.VALUE;
      let colStr = '';
      let c = col;
      while (c > 0) { c--; colStr = String.fromCharCode(65 + (c % 26)) + colStr; c = Math.floor(c / 26); }
      let result = '';
      if (a1 === false || a1 === 0 || a1 === 'FALSE') {
        const rAbs = absType === 1 || absType === 2;
        const cAbs = absType === 1 || absType === 3;
        result = (rAbs ? `R${row}` : `R[${row}]`) + (cAbs ? `C${col}` : `C[${col}]`);
      } else {
        const colAbs = absType === 1 || absType === 2 ? '$' : '';
        const rowAbs = absType === 1 || absType === 3 ? '$' : '';
        result = `${colAbs}${colStr}${rowAbs}${row}`;
      }
      if (sheet) result = `${sheet}!${result}`;
      return result;
    },
    TRANSPOSE: async (args, meta) => {
      const arr = args[0];
      if (!Array.isArray(arr)) return arr;
      let rows = 1, cols = arr.length;
      if (meta && meta[0]) { rows = meta[0].rows || 1; cols = meta[0].cols || arr.length; }
      const result: any[] = [];
      for (let c = 0; c < cols; c++) {
        for (let r = 0; r < rows; r++) {
          result.push(arr[r * cols + c] ?? '');
        }
      }
      return result;
    },

    // ---- Dynamic Array functions ----
    SORT: async (args) => {
      const arr = args[0];
      if (!Array.isArray(arr)) return ERRORS.VALUE;
      const sortOrder = args[2] !== undefined ? Number(args[2]) : 1;
      const sorted = [...arr];
      sorted.sort((a: any, b: any) => {
        const na = Number(a); const nb = Number(b);
        if (!Number.isNaN(na) && !Number.isNaN(nb)) return sortOrder === 1 ? na - nb : nb - na;
        return sortOrder === 1 ? String(a).localeCompare(String(b)) : String(b).localeCompare(String(a));
      });
      return sorted;
    },
    UNIQUE: async (args) => {
      const arr = args[0];
      if (!Array.isArray(arr)) return ERRORS.VALUE;
      const exactlyOnce = args[2] !== undefined ? Boolean(args[2]) : false;
      if (exactlyOnce) {
        const counts = new Map<any, number>();
        for (const v of arr) counts.set(v, (counts.get(v) || 0) + 1);
        return arr.filter((v: any) => counts.get(v) === 1);
      }
      return [...new Set(arr)];
    },
    FILTER: async (args) => {
      const arr = args[0];
      const include = args[1];
      const ifEmpty = args[2] !== undefined ? args[2] : ERRORS.VALUE;
      if (!Array.isArray(arr) || !Array.isArray(include)) return ERRORS.VALUE;
      const result: any[] = [];
      for (let i = 0; i < arr.length && i < include.length; i++) {
        if (include[i] === true || include[i] === 1 || include[i] === '1') result.push(arr[i]);
      }
      return result.length > 0 ? result : ifEmpty;
    },
    SORTBY: async (args) => {
      const arr = args[0];
      const byArr = args[1];
      const order = args[2] !== undefined ? Number(args[2]) : 1;
      if (!Array.isArray(arr) || !Array.isArray(byArr)) return ERRORS.VALUE;
      const indices = arr.map((_: any, i: number) => i);
      indices.sort((a: number, b: number) => {
        const va = byArr[a]; const vb = byArr[b];
        const na = Number(va); const nb = Number(vb);
        if (!Number.isNaN(na) && !Number.isNaN(nb)) return order === 1 ? na - nb : nb - na;
        return order === 1 ? String(va ?? '').localeCompare(String(vb ?? '')) : String(vb ?? '').localeCompare(String(va ?? ''));
      });
      return indices.map((i: number) => arr[i]);
    },
    FORMULATEXT: async (_args, meta) => {
      if (meta && meta[0] && meta[0].formula) return meta[0].formula;
      return ERRORS.NA;
    },
    AREAS: async (args) => {
      return args.length || 1;
    },
  };
}
