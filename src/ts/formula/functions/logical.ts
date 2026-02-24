// logical.ts - Logical and conditional functions

import { ERRORS, isError } from '../Errors';
import type { BuiltinFunction } from '../../types';

export function getLogicalFunctions(): Record<string, BuiltinFunction> {
  return {
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
    IFNA: async (args) => args[0] === ERRORS.NA ? (args[1] !== undefined ? args[1] : '') : args[0],
    IFS: async (args) => {
      for (let i = 0; i < args.length - 1; i += 2) {
        const cond = args[i];
        if (cond && cond !== 0 && cond !== '' && cond !== false && cond !== 'FALSE') {
          return args[i + 1] !== undefined ? args[i + 1] : '';
        }
      }
      return ERRORS.NA;
    },
    SWITCH: async (args) => {
      if (args.length < 2) return ERRORS.VALUE;
      const expr = args[0];
      for (let i = 1; i < args.length - 1; i += 2) {
        if (expr === args[i] || (typeof expr === 'string' && typeof args[i] === 'string' && expr.toLowerCase() === args[i].toLowerCase()))
          return args[i + 1] !== undefined ? args[i + 1] : '';
      }
      return (args.length % 2 === 0) ? args[args.length - 1] : ERRORS.NA;
    },
    XOR: async (args) => {
      let trueCount = 0;
      for (const a of args) { if (a && a !== 0 && a !== '' && a !== false && a !== 'FALSE') trueCount++; }
      return trueCount % 2 === 1;
    },
  };
}
