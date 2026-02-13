// FormulaShifter.ts - Shift A1 cell references in formulas (for shared formulas)

import { CellReference } from '../core/CellReference';

/**
 * Shifts cell references within formula strings to support shared formulas.
 * Respects absolute ($) references.
 */
export class FormulaShifter {
  /**
   * Shift a single A1-style reference by row/col offset, respecting $ absolute markers.
   */
  static shiftA1Ref(ref: string, rowOffset: number, colOffset: number): string {
    const match = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(ref);
    if (!match) return ref;
    const colAbs = match[1] === '$';
    const colLetters = match[2];
    const rowAbs = match[3] === '$';
    const rowNum = parseInt(match[4], 10);
    let colIndex = CellReference.parse(`${colLetters}1`).col;
    let rowIndex = rowNum;
    if (!colAbs) colIndex += colOffset;
    if (!rowAbs) rowIndex += rowOffset;
    if (colIndex < 1 || rowIndex < 1) return '#REF!';
    const colName = CellReference.columnIndexToName(colIndex);
    return `${colAbs ? '$' : ''}${colName}${rowAbs ? '$' : ''}${rowIndex}`;
  }

  /**
   * Shift all cell references in a formula string.
   */
  static shiftFormulaRefs(formula: string, rowOffset: number, colOffset: number): string {
    if (!rowOffset && !colOffset) return formula;
    const refRegex = /((?:'[^']+'|[A-Za-z0-9_]+)!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/g;

    const adjustChunk = (chunk: string): string =>
      chunk.replace(refRegex, (match, sheetPrefix: string | undefined, startRef: string, endRef?: string) => {
        const shiftedStart = FormulaShifter.shiftA1Ref(startRef, rowOffset, colOffset);
        const shiftedEnd = endRef ? FormulaShifter.shiftA1Ref(endRef, rowOffset, colOffset) : null;
        return `${sheetPrefix || ''}${shiftedStart}${shiftedEnd ? `:${shiftedEnd}` : ''}`;
      });

    let output = '';
    let buffer = '';
    let inString = false;

    for (let i = 0; i < formula.length; i += 1) {
      const ch = formula[i];
      if (ch === '"') {
        if (inString && formula[i + 1] === '"') {
          output += '""';
          i += 1;
          continue;
        }
        if (!inString) {
          output += adjustChunk(buffer);
          buffer = '';
          inString = true;
          output += '"';
        } else {
          output += '"';
          inString = false;
        }
        continue;
      }
      if (inString) output += ch;
      else buffer += ch;
    }

    if (buffer) output += adjustChunk(buffer);
    return output;
  }

  /**
   * Derive a shared formula for a target cell from an anchor cell's formula.
   */
  static deriveSharedFormula(baseFormula: string, anchorRef: string, targetRef: string): string {
    if (!baseFormula || !anchorRef || !targetRef) return baseFormula || '';
    const anchor = CellReference.parse(anchorRef);
    const target = CellReference.parse(targetRef);
    if (!anchor.row || !anchor.col || !target.row || !target.col) return baseFormula;
    const rowOffset = target.row - anchor.row;
    const colOffset = target.col - anchor.col;
    return FormulaShifter.shiftFormulaRefs(baseFormula, rowOffset, colOffset);
  }
}
