// BuiltinFunctions.ts - Aggregator that merges all category function modules

import type { BuiltinFunction } from '../types';
import {
  getMathFunctions,
  getLogicalFunctions,
  getTextFunctions,
  getDateFunctions,
  getLookupFunctions,
  getStatisticalFunctions,
  getInformationFunctions,
  getFinancialFunctions,
} from './functions';

/**
 * Creates the registry of all built-in Excel functions.
 *
 * Each category is defined in its own module under ./functions/:
 *   math, logical, text, date, lookup, statistical, information, financial
 *
 * Shared helpers (matchCriteria, compareValues) live in ./functions/helpers.
 */
export class BuiltinFunctions {
  /**
   * Return a map of function name -> async handler.
   */
  static create(): Record<string, BuiltinFunction> {
    return {
      ...getMathFunctions(),
      ...getLogicalFunctions(),
      ...getTextFunctions(),
      ...getDateFunctions(),
      ...getLookupFunctions(),
      ...getStatisticalFunctions(),
      ...getInformationFunctions(),
      ...getFinancialFunctions(),
    };
  }
}