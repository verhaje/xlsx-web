// information.ts - Information and type-checking functions

import { ERRORS, isError } from '../Errors';
import type { BuiltinFunction } from '../../types';

export function getInformationFunctions(): Record<string, BuiltinFunction> {
  return {
    ISERROR: async (args) => isError(args[0]),
    ISBLANK: async (args) => args[0] === '' || args[0] === null || args[0] === undefined,
    ISNUMBER: async (args) => typeof args[0] === 'number' && !Number.isNaN(args[0]),
    ISTEXT: async (args) => typeof args[0] === 'string',
    TYPE: async (args) => {
      const v = args[0];
      if (typeof v === 'number') return 1;
      if (typeof v === 'string') {
        if (isError(v)) return 16;
        return 2;
      }
      if (typeof v === 'boolean') return 4;
      if (Array.isArray(v)) return 64;
      return 1;
    },
    N: async (args) => {
      const v = args[0];
      if (typeof v === 'number') return v;
      if (typeof v === 'boolean') return v ? 1 : 0;
      if (v === 'TRUE') return 1;
      if (v === 'FALSE') return 0;
      if (isError(v)) return v;
      return 0;
    },
    NA: async () => ERRORS.NA,
    'ERROR.TYPE': async (args) => {
      const v = args[0];
      switch (v) {
        case ERRORS.NULL: return 1;
        case ERRORS.DIV0: return 2;
        case ERRORS.VALUE: return 3;
        case ERRORS.REF: return 4;
        case ERRORS.NAME: return 5;
        case ERRORS.NUM: return 6;
        case ERRORS.NA: return 7;
        default: return ERRORS.NA;
      }
    },
    ISEVEN: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      return Math.trunc(n) % 2 === 0;
    },
    ISODD: async (args) => {
      const n = Number(args[0]);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      return Math.trunc(n) % 2 !== 0;
    },
    ISLOGICAL: async (args) => typeof args[0] === 'boolean' || args[0] === 'TRUE' || args[0] === 'FALSE',
    ISNA: async (args) => args[0] === ERRORS.NA,
    ISNONTEXT: async (args) => typeof args[0] !== 'string',
    ISERR: async (args) => isError(args[0]) && args[0] !== ERRORS.NA,
    ISREF: async (_args, meta) => !!(meta && meta[0] && meta[0].isRef),
    ISFORMULA: async (_args, meta) => !!(meta && meta[0] && meta[0].isFormula),
    SHEET: async () => 1,
    SHEETS: async () => 1,
    CELL: async () => {
      // Stub — requires environment context not available in formula engine
      return '';
    },
    INFO: async (args) => {
      const type = String(args[0]).toLowerCase();
      switch (type) {
        case 'osversion': return 'Windows';
        case 'release': return '16.0';
        case 'system': return 'pcdos';
        default: return '';
      }
    },
  };
}
