// CellEvaluator.ts - Cell evaluation with cross-sheet caching

import type JSZip from 'jszip';
import type { SheetInfo, SheetModel, SheetCacheEntry, IFormulaEngine } from '../types';
import { XmlParser } from '../core/XmlParser';
import { CellReference } from '../core/CellReference';
import { SheetParser } from '../workbook/SheetParser';
import { FormulaShifter } from './FormulaShifter';

interface CellFormulaInfo {
  formulaText: string | null;
  sharedInfo: { si: string; anchorRef?: string; baseFormula?: string } | null;
}

/**
 * Manages cell evaluation, cross-sheet references, and caching during rendering.
 */
export class CellEvaluator {
  private zip: JSZip;
  private sheet: SheetInfo;
  private sheets: SheetInfo[];
  private sharedStrings: string[];
  private formulaEngine: IFormulaEngine;

  /** Per-sheet cached cell maps and sheet docs */
  private sheetCache: Map<string, any>;
  /** Promise coalescing for sheet building */
  private buildingSheets: Map<string, Promise<SheetCacheEntry>>;
  /** Cells currently being evaluated (circular reference detection) */
  private evaluatingCells: Set<string>;
  /** Promise coalescing for cell evaluation */
  private pendingCells: Map<string, Promise<any>>;

  constructor(
    zip: JSZip,
    sheet: SheetInfo,
    sheets: SheetInfo[],
    sharedStrings: string[],
    formulaEngine: IFormulaEngine,
    sheetCache: Map<string, any>
  ) {
    this.zip = zip;
    this.sheet = sheet;
    this.sheets = sheets;
    this.sharedStrings = sharedStrings;
    this.formulaEngine = formulaEngine;
    this.sheetCache = sheetCache;
    this.buildingSheets = new Map();
    this.evaluatingCells = new Set();
    this.pendingCells = new Map();
  }

  /**
   * Get the formula text for a cell element, resolving shared formulas if needed.
   */
  static getCellFormula(cellEl: Element, cellRef: string, sheetModel: SheetModel): string | null {
    const fNode = cellEl.getElementsByTagName('f')[0];
    if (!fNode) return null;
    const fType = fNode.getAttribute('t');
    if (fType === 'shared' && sheetModel?.sharedCells) {
      const sharedInfo = sheetModel.sharedCells.get(cellRef);
      if (sharedInfo?.baseFormula && sharedInfo?.anchorRef) {
        return FormulaShifter.deriveSharedFormula(sharedInfo.baseFormula, sharedInfo.anchorRef, cellRef);
      }
    }
    return fNode.textContent || '';
  }

  /**
   * Get a cache key for a shared formula instance.
   */
  static getSharedCacheKey(
    sheetName: string,
    sharedInfo: { si: string; anchorRef?: string } | null,
    targetRef: string
  ): string | null {
    if (!sharedInfo?.anchorRef) return null;
    const anchor = CellReference.parse(sharedInfo.anchorRef);
    const target = CellReference.parse(targetRef);
    if (!anchor.row || !anchor.col || !target.row || !target.col) return null;
    return `${sheetName}::${sharedInfo.si}::${target.row - anchor.row},${target.col - anchor.col}`;
  }

  /**
   * Evaluate a cell by A1 reference, resolving cross-sheet refs and using caches.
   */
  async evaluateCellByRef(refA1: string, contextSheetName?: string): Promise<any> {
    let targetSheetName = contextSheetName || this.sheet.name;
    let a1 = refA1;

    if (typeof refA1 === 'string' && refA1.includes('!')) {
      const parts = refA1.split('!');
      targetSheetName = parts.slice(0, -1).join('!');
      a1 = parts[parts.length - 1];
    }
    targetSheetName = CellEvaluator.stripSheetQuotes(targetSheetName);

    let targetSheetObj: SheetInfo | null = null;
    if (targetSheetName === this.sheet.name) {
      targetSheetObj = this.sheet;
    } else if (Array.isArray(this.sheets)) {
      targetSheetObj = this.sheets.find((s) => s.name === targetSheetName) || null;
    }

    const info = targetSheetObj
      ? await this.buildCellMapForSheet(targetSheetObj)
      : { cellMap: new Map<string, Element>(), sheetDoc: null, sheetModel: null } as any;

    a1 = CellReference.mapMergedRef(a1, info.sheetModel);
    const parsed = CellReference.parse(a1);
    if (!parsed.col || !parsed.row) return '';

    const key = `${targetSheetName}::${parsed.row}-${parsed.col}`;

    // Check values cache
    if (this.sheetCache.has('__values') && this.sheetCache.get('__values').has(key)) {
      return this.sheetCache.get('__values').get(key);
    }
    // Prevent thundering herd
    if (this.pendingCells.has(key)) return this.pendingCells.get(key);
    // Ensure values map
    if (!this.sheetCache.has('__values')) this.sheetCache.set('__values', new Map());
    // Circular reference guard
    if (this.evaluatingCells.has(key)) return '#REF!';
    this.evaluatingCells.add(key);

    const promise = (async () => {
      const cmap: Map<string, Element> = info.cellMap || new Map();
      const cellEl = cmap.get(`${parsed.row}-${parsed.col}`);
      if (!cellEl) {
        this.sheetCache.get('__values').set(key, '');
        return '';
      }

      let val: any;
      const cellRef = CellReference.toA1(parsed.row, parsed.col);
      const formulaText = CellEvaluator.getCellFormula(cellEl, cellRef, info.sheetModel);

      if (formulaText !== null && this.formulaEngine) {
        const sharedInfo = info.sheetModel?.sharedCells?.get(cellRef) || null;
        const cacheKey = CellEvaluator.getSharedCacheKey(targetSheetName, sharedInfo, cellRef);
        if (cacheKey && this.sheetCache.get('__sharedValues')?.has(cacheKey)) {
          val = this.sheetCache.get('__sharedValues').get(cacheKey);
        } else {
          val = await this.formulaEngine.evaluateFormula(formulaText, {
            resolveCell: async (r: string) => await this.evaluateCellByRef(r, targetSheetName),
            sharedStrings: this.sharedStrings,
            zip: this.zip,
            sheetDoc: info.sheetDoc,
          });
          if (cacheKey) {
            if (!this.sheetCache.has('__sharedValues')) this.sheetCache.set('__sharedValues', new Map());
            this.sheetCache.get('__sharedValues').set(cacheKey, val);
          }
        }
      } else {
        val = SheetParser.extractCellValue(cellEl, this.sharedStrings);
      }

      this.sheetCache.get('__values').set(key, val);
      return val;
    })();

    this.pendingCells.set(key, promise);
    try {
      return await promise;
    } finally {
      this.evaluatingCells.delete(key);
      this.pendingCells.delete(key);
    }
  }

  // ---- Private helpers ----

  private async buildCellMapForSheet(sheetObj: SheetInfo): Promise<SheetCacheEntry> {
    if (!sheetObj?.target) return { cellMap: new Map(), sheetDoc: null, maxRow: 0, maxCol: 0 };
    if (this.sheetCache.has(sheetObj.name)) return this.sheetCache.get(sheetObj.name);
    if (this.buildingSheets.has(sheetObj.name)) return this.buildingSheets.get(sheetObj.name)!;

    const promise = (async (): Promise<SheetCacheEntry> => {
      try {
        const xml = await XmlParser.readZipText(this.zip, sheetObj.target!);
        const doc = XmlParser.parseXml(xml);
        const data = doc.getElementsByTagName('sheetData')[0];
        const rows = data ? Array.from(data.getElementsByTagName('row')) : [];
        const cmap = new Map<string, Element>();
        let mRow = 0;
        let mCol = 0;
        rows.forEach((r) => {
          const rowIndex = parseInt(r.getAttribute('r') || '0', 10);
          const actualRow = Number.isNaN(rowIndex) || rowIndex === 0 ? mRow + 1 : rowIndex;
          mRow = Math.max(mRow, actualRow);
          const cells = Array.from(r.getElementsByTagName('c'));
          cells.forEach((cell) => {
            const ref = cell.getAttribute('r') || '';
            const { col } = CellReference.parse(ref);
            if (col === 0) return;
            mCol = Math.max(mCol, col);
            cmap.set(CellReference.cellKey(actualRow, col), cell);
          });
        });
        const model = SheetParser.buildSheetModel(doc);
        const info: SheetCacheEntry = { cellMap: cmap, sheetDoc: doc, maxRow: mRow, maxCol: mCol, sheetModel: model };
        this.sheetCache.set(sheetObj.name, info);
        return info;
      } catch {
        return { cellMap: new Map(), sheetDoc: null, maxRow: 0, maxCol: 0 };
      } finally {
        this.buildingSheets.delete(sheetObj.name);
      }
    })();

    this.buildingSheets.set(sheetObj.name, promise);
    return promise;
  }

  private static stripSheetQuotes(name: string): string {
    if (name.length >= 2 && name.startsWith("'") && name.endsWith("'")) return name.slice(1, -1);
    return name;
  }
}
