// SheetRenderer.ts - Main DOM rendering for a sheet

import type JSZip from 'jszip';
import type {
  SheetInfo, SheetModel, StyleResult, ColumnStyleRange,
  ImageAnchor, CellStyleResult, RenderSheetOptions,
  CfEntry, SheetCacheEntry, CssProperties, IFormulaEngine,
} from '../types';
import { XmlParser } from '../core/XmlParser';
import { CellReference } from '../core/CellReference';
import { SheetParser } from '../workbook/SheetParser';
import { StyleApplicator } from '../styles/StyleApplicator';
import { CellEvaluator } from './CellEvaluator';
import { ConditionalFormatting } from './ConditionalFormatting';
import { ImageRenderer } from './ImageRenderer';
import { FormulaShifter } from './FormulaShifter';

/**
 * Renders a sheet to a DOM table, evaluating formulas and applying styles and conditional formatting.
 */
export class SheetRenderer {
  /**
   * Render a sheet into the given table container.
   * Returns the created <table> element (so callers can attach resizers etc.).
   */
  static async renderSheet(options: RenderSheetOptions): Promise<HTMLTableElement | void> {
    const { zip, sheet, sharedStrings, styles, tableContainer, sheetNameEl, formulaEngine, sheets } = options;
    if (!sheet.target) return;

    const sheetXml = await XmlParser.readZipText(zip, sheet.target);
    const sheetDoc = XmlParser.parseXml(sheetXml);
    const sheetData = sheetDoc.getElementsByTagName('sheetData')[0];
    const sheetModel = SheetParser.buildSheetModel(sheetDoc);
    const columnStyleRanges = SheetParser.parseColumnStyles(sheetDoc);

    // Load images for this sheet
    const images = await SheetRenderer.loadSheetImages(zip, sheet.target);

    // Parse conditional formatting rules
    const cfMap = ConditionalFormatting.parseCfRules(sheetDoc, styles);

    if (!sheetData) {
      tableContainer.innerHTML = '<div class="status">No data found in sheet.</div>';
      return;
    }

    // Build cell map and extents
    const rows = Array.from(sheetData.getElementsByTagName('row'));
    const cellMap = new Map<string, Element>();
    const rowStyleMap = new Map<number, string>();
    let maxRow = 0;
    let maxCol = 0;

    rows.forEach((row) => {
      const rowIndex = parseInt(row.getAttribute('r') || '0', 10);
      const actualRow = Number.isNaN(rowIndex) || rowIndex === 0 ? maxRow + 1 : rowIndex;
      maxRow = Math.max(maxRow, actualRow);
      const rowStyle = row.getAttribute('s');
      if (rowStyle != null) rowStyleMap.set(actualRow, rowStyle);
      const cells = Array.from(row.getElementsByTagName('c'));
      cells.forEach((cell) => {
        const ref = cell.getAttribute('r') || '';
        const { col } = CellReference.parse(ref);
        if (col === 0) return;
        maxCol = Math.max(maxCol, col);
        cellMap.set(CellReference.cellKey(actualRow, col), cell);
      });
    });

    // Sheet cache (evaluation caching across sheets)
    const sheetCache = new Map<string, any>();
    sheetCache.set(sheet.name, { cellMap, sheetDoc, maxRow, maxCol, sheetModel });

    // Cell evaluator
    const cellEvaluator = new CellEvaluator(zip, sheet, sheets, sharedStrings, formulaEngine, sheetCache);

    // Conditional formatting evaluator
    const cfEvaluator = new ConditionalFormatting(
      styles,
      (ref: string, ctx?: string) => cellEvaluator.evaluateCellByRef(ref, ctx),
      formulaEngine,
      sharedStrings,
      zip,
      sheetDoc
    );

    // Build column style index
    const colStyleByIndex: (string | undefined)[] = [];
    if (columnStyleRanges.length) {
      columnStyleRanges.forEach(({ min, max, style }) => {
        for (let c = min; c <= max; c += 1) colStyleByIndex[c] = style;
      });
    }

    // Create table structure
    const table = document.createElement('table');
    table.className = 'sheet-table';

    const colgroup = document.createElement('colgroup');
    for (let c = 1; c <= maxCol; c += 1) {
      const colEl = document.createElement('col');
      colEl.dataset.colIndex = String(c);
      colgroup.appendChild(colEl);
    }
    table.appendChild(colgroup);

    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const corner = document.createElement('th');
    corner.textContent = '';
    headerRow.appendChild(corner);
    for (let col = 1; col <= maxCol; col += 1) {
      const th = document.createElement('th');
      th.textContent = CellReference.columnIndexToName(col);
      th.dataset.colIndex = String(col);
      headerRow.appendChild(th);
    }
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    // Yield helper for large renders
    const yieldToEventLoop = () => new Promise<void>((res) => setTimeout(res, 0));
    const batchSize = 50;

    // Show loading overlay
    const loadingOverlay = document.createElement('div');
    loadingOverlay.className = 'sheet-loading-overlay';
    loadingOverlay.innerHTML =
      '<div class="sheet-loading-spinner"></div><div class="sheet-loading-text">Rendering sheet\u2026</div>';
    tableContainer.innerHTML = '';
    tableContainer.appendChild(loadingOverlay);
    await yieldToEventLoop();

    // Render rows in batches
    for (let startRow = 1; startRow <= maxRow; startRow += batchSize) {
      const endRow = Math.min(maxRow, startRow + batchSize - 1);
      const frag = document.createDocumentFragment();
      const deferredWork: DeferredWorkItem[] = [];

      for (let row = startRow; row <= endRow; row += 1) {
        const tr = document.createElement('tr');
        const rowHeader = document.createElement('th');
        rowHeader.textContent = row.toString();
        tr.appendChild(rowHeader);
        rowHeader.dataset.rowIndex = String(row);

        for (let col = 1; col <= maxCol; col += 1) {
          const key = `${row}-${col}`;
          if (sheetModel.coveredMap.has(key)) continue;

          const td = document.createElement('td');
          td.dataset.rowIndex = String(row);
          td.dataset.colIndex = String(col);
          const cellEl = cellMap.get(key);
          let value: any = '';
          let cellStyle: CellStyleResult | null = null;
          const rowStyleIndex = rowStyleMap.get(row);
          const colStyleIndex = colStyleByIndex[col];
          const fallbackStyleIndex = rowStyleIndex != null ? rowStyleIndex : colStyleIndex;
          let needsFormulaEval = false;

          if (cellEl) {
            const cellRef = CellReference.toA1(row, col);
            const formulaText = CellEvaluator.getCellFormula(cellEl, cellRef, sheetModel);

            if (formulaText !== null && formulaEngine) {
              // Check shared-formula cache synchronously
              const sharedInfo = sheetModel.sharedCells.get(cellRef) || null;
              const cacheKey = CellEvaluator.getSharedCacheKey(sheet.name, sharedInfo, cellRef);
              if (cacheKey && sheetCache.get('__sharedValues')?.has(cacheKey)) {
                value = sheetCache.get('__sharedValues').get(cacheKey);
              } else {
                needsFormulaEval = true;
                deferredWork.push({
                  td, row, col, key, cellEl, cellRef, formulaText, cacheKey,
                  fallbackStyleIndex, cfOnly: false, value: '', cfEntries: [],
                });
              }
              td.dataset.formula = formulaText;
            } else {
              value = SheetParser.extractCellValue(cellEl, sharedStrings);
            }
            td.dataset.value = value;
            cellStyle = StyleApplicator.getCellStyle(
              cellEl, styles.cellXfs, styles.fonts, styles.themeColors,
              styles.fills, styles.borders, styles.cellStyleXfs, fallbackStyleIndex
            );
          } else if (fallbackStyleIndex != null) {
            cellStyle = StyleApplicator.getCellStyle(
              null, styles.cellXfs, styles.fonts, styles.themeColors,
              styles.fills, styles.borders, styles.cellStyleXfs, fallbackStyleIndex
            );
          }

          if (!needsFormulaEval) {
            SheetRenderer.applyCellContent(td, value, cellStyle, sheetModel, key);
            const cfEntries = cfMap.get(key) || [];
            if (cfEntries.length) {
              deferredWork.push({ td, row, col, key, cfOnly: true, value, cfEntries,
                cellEl: null as any, cellRef: '', formulaText: '', cacheKey: null, fallbackStyleIndex });
            }
          }
          tr.appendChild(td);
        }
        frag.appendChild(tr);
      }

      // Evaluate deferred formulas and conditional formatting concurrently
      if (deferredWork.length > 0) {
        await Promise.all(
          deferredWork.map(async (item) => {
            if (item.cfOnly) {
              for (const entry of item.cfEntries) {
                const result = await cfEvaluator.evaluate(entry, item.value, item.row, item.col);
                if (result.matched) {
                  ConditionalFormatting.applyCfToTd(item.td, result);
                  if (entry.rule.stopIfTrue) break;
                }
              }
              return;
            }

            const { td, row, col, key, cellEl, cellRef, formulaText, cacheKey, fallbackStyleIndex } = item;
            let val: any;

            // Re-check shared cache
            if (cacheKey && sheetCache.get('__sharedValues')?.has(cacheKey)) {
              val = sheetCache.get('__sharedValues').get(cacheKey);
            } else {
              val = await formulaEngine.evaluateFormula(formulaText, {
                resolveCell: async (r: string) => await cellEvaluator.evaluateCellByRef(r),
                sharedStrings,
                zip,
                sheetDoc,
              });
              if (cacheKey) {
                if (!sheetCache.has('__sharedValues')) sheetCache.set('__sharedValues', new Map());
                sheetCache.get('__sharedValues').set(cacheKey, val);
              }
            }

            td.dataset.value = val;
            const cellStyle = StyleApplicator.getCellStyle(
              cellEl, styles.cellXfs, styles.fonts, styles.themeColors,
              styles.fills, styles.borders, styles.cellStyleXfs, fallbackStyleIndex
            );
            SheetRenderer.applyCellContent(td, val, cellStyle, sheetModel, key);

            // Apply conditional formatting
            const cfEntries = cfMap.get(key) || [];
            if (cfEntries.length) {
              for (const entry of cfEntries) {
                const result = await cfEvaluator.evaluate(entry, val, row, col);
                if (result.matched) {
                  ConditionalFormatting.applyCfToTd(td, result);
                  if (entry.rule.stopIfTrue) break;
                }
              }
            }
          })
        );
      }

      tbody.appendChild(frag);
      await yieldToEventLoop();
    }

    table.appendChild(tbody);
    sheetNameEl.textContent = sheet.name;

    // Remove loading overlay and show table
    tableContainer.innerHTML = '';
    tableContainer.appendChild(table);

    // Render images
    if (images.length > 0) {
      ImageRenderer.renderImages(tableContainer, images);
    }

    return table;
  }

  // ---- Private helpers ----

  /**
   * Load images referenced by a sheet.
   */
  private static async loadSheetImages(zip: JSZip, target: string): Promise<ImageAnchor[]> {
    const sheetRels = await SheetParser.loadSheetRelationships(zip, target);
    const images: ImageAnchor[] = [];
    for (const [relId, relTarget] of sheetRels) {
      if (relTarget.includes('drawings/')) {
        const drawingImages = await SheetParser.loadDrawing(zip, XmlParser.normalizeTargetPath(relTarget));
        for (const image of drawingImages) {
          const mediaData = await SheetParser.loadMedia(zip, image.embed);
          if (mediaData) {
            images.push({ ...image, dataUrl: mediaData });
          }
        }
      }
    }
    return images;
  }

  /**
   * Apply value, style, and merge-cell spans to a table cell.
   */
  private static applyCellContent(
    td: HTMLTableCellElement,
    value: any,
    cellStyle: CellStyleResult | null,
    sheetModel: SheetModel,
    key: string
  ): void {
    const anchorSpan = sheetModel.anchorMap.get(key);
    if (anchorSpan) {
      if (anchorSpan.colSpan > 1) td.colSpan = anchorSpan.colSpan;
      if (anchorSpan.rowSpan > 1) td.rowSpan = anchorSpan.rowSpan;
    }

    if (cellStyle && cellStyle.type === 'runs') {
      if (cellStyle.baseline) StyleApplicator.applyCssToElement(td, cellStyle.baseline);
      if (value === undefined || value === null || value === '') {
        td.textContent = '';
        td.classList.add('empty-cell');
      } else {
        cellStyle.runs.forEach((run) => {
          const span = document.createElement('span');
          span.textContent = run.text;
          StyleApplicator.applyCssToElement(span, run.style);
          td.appendChild(span);
        });
      }
    } else {
      if (value === undefined || value === null || value === '') {
        td.textContent = '';
        td.classList.add('empty-cell');
      } else {
        td.textContent = value;
      }
      if (cellStyle && cellStyle.style) StyleApplicator.applyCssToElement(td, cellStyle.style);
    }
  }
}

// ---- Internal types ----

interface DeferredWorkItem {
  td: HTMLTableCellElement;
  row: number;
  col: number;
  key: string;
  cfOnly: boolean;
  value: any;
  cfEntries: CfEntry[];
  cellEl: Element;
  cellRef: string;
  formulaText: string;
  cacheKey: string | null;
  fallbackStyleIndex: string | undefined;
}
