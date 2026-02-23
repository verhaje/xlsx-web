// SheetRenderer.ts - Main DOM rendering for a sheet

import type JSZip from 'jszip';
import type {
  SheetInfo, SheetModel, StyleResult, ColumnStyleRange,
  ImageAnchor, ChartAnchor, CellStyleResult, RenderSheetOptions,
  CfEntry, SheetCacheEntry, CssProperties, IFormulaEngine,
} from '../types';
import { XmlParser } from '../core/XmlParser';
import { CellReference } from '../core/CellReference';
import { SheetParser } from '../workbook/SheetParser';
import { StyleApplicator } from '../styles/StyleApplicator';
import { CellEvaluator } from './CellEvaluator';
import { ConditionalFormatting } from './ConditionalFormatting';
import { ImageRenderer } from './ImageRenderer';
import { ChartRenderer } from './ChartRenderer';
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

    const sheetDoc = await XmlParser.readZipXml(zip, sheet.target);
    const sheetData = sheetDoc.getElementsByTagName('sheetData')[0];
    // Reuse the SheetModel from WorkbookManager if provided, avoiding
    // a redundant parse of merge cells and shared formulas.
    const sheetModel = options.sheetModel ?? SheetParser.buildSheetModel(sheetDoc);
    const columnStyleRanges = SheetParser.parseColumnStyles(sheetDoc);

    // Load images and charts in parallel, sharing the relationship map
    const [images, charts] = await SheetRenderer.loadSheetAssets(zip, sheet.target);

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

    // Extend grid by +20 rows and +20 columns beyond defined data
    const EXTRA_ROWS = 20;
    const EXTRA_COLS = 20;
    const renderMaxRow = maxRow + EXTRA_ROWS;
    const renderMaxCol = maxCol + EXTRA_COLS;

    // Sheet cache (evaluation caching across sheets)
    const sheetCache = new Map<string, any>();
    sheetCache.set(sheet.name, { cellMap, sheetDoc, maxRow: renderMaxRow, maxCol: renderMaxCol, sheetModel });

    // Shared range cache for all formulas in this sheet (avoids re-resolving identical ranges)
    const rangeCache = new Map<string, any[]>();

    // Cell evaluator
    const cellEvaluator = new CellEvaluator(zip, sheet, sheets, sharedStrings, formulaEngine, sheetCache, rangeCache);

    // Conditional formatting evaluator
    const cfEvaluator = new ConditionalFormatting(
      styles,
      (ref: string, ctx?: string) => cellEvaluator.evaluateCellByRef(ref, ctx),
      formulaEngine,
      sharedStrings,
      zip,
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
    // Add a <col> for the row-header column so <col> elements align 1:1 with table columns
    const rowHeaderCol = document.createElement('col');
    rowHeaderCol.className = 'row-header-col';
    colgroup.appendChild(rowHeaderCol);
    for (let c = 1; c <= renderMaxCol; c += 1) {
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
    for (let col = 1; col <= renderMaxCol; col += 1) {
      const th = document.createElement('th');
      th.textContent = CellReference.columnIndexToName(col);
      th.dataset.colIndex = String(col);
      headerRow.appendChild(th);
    }
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    // Yield helper for large renders – use setTimeout to schedule a macrotask.
    // Using rAF here is counterproductive: all subsequent awaited microtasks
    // (from Promise.all / async-await) chain inside the *same* rAF callback,
    // so the browser never gets a chance to paint until the entire batch
    // finishes. setTimeout(0) pushes continuation to the next task, allowing
    // the browser to paint the loading spinner and keep the UI responsive.
    const yieldToEventLoop = () =>
      new Promise<void>((res) => {
        setTimeout(res, 0);
      });
    // Larger batches reduce yield overhead (each yield triggers layout); 500 is
    // a good trade-off between responsiveness and throughput.
    const batchSize = 500;

    // Show loading overlay
    const loadingOverlay = document.createElement('div');
    loadingOverlay.className = 'sheet-loading-overlay';
    loadingOverlay.innerHTML =
      '<div class=\"sheet-loading-spinner\"></div><div class=\"sheet-loading-text\">Rendering sheet\u2026</div>';
    tableContainer.replaceChildren(loadingOverlay);
    await yieldToEventLoop();
    // Wrap the entire rendering pipeline in try/finally so that the loading
    // overlay is always replaced even if a formula evaluation throws.
    try {
    // Render rows in batches
    for (let startRow = 1; startRow <= renderMaxRow; startRow += batchSize) {
      const endRow = Math.min(renderMaxRow, startRow + batchSize - 1);
      const frag = document.createDocumentFragment();
      const deferredWork: DeferredWorkItem[] = [];

      for (let row = startRow; row <= endRow; row += 1) {
        const tr = document.createElement('tr');
        const rowHeader = document.createElement('th');
        rowHeader.textContent = row.toString();
        tr.appendChild(rowHeader);
        rowHeader.dataset.rowIndex = String(row);

        for (let col = 1; col <= renderMaxCol; col += 1) {
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

      // Evaluate deferred formulas and conditional formatting
      if (deferredWork.length > 0) {
        // Split deferred work: CF-only items can be concurrent, formula items are sequential
        const cfOnlyItems = deferredWork.filter((item) => item.cfOnly);
        const formulaItems = deferredWork.filter((item) => !item.cfOnly);

        // Evaluate formulas in parallel chunks for throughput.
        // Range-cache is shared, so earlier evaluations still benefit later ones.
        // Yield to the event loop periodically to keep the UI responsive on
        // formula-heavy sheets (prevents the 27s hang seen in the trace).
        const FORMULA_CHUNK = 32;
        let chunksSinceYield = 0;
        const YIELD_EVERY_N_CHUNKS = 4; // yield every ~128 formulas
        for (let fi = 0; fi < formulaItems.length; fi += FORMULA_CHUNK) {
          const chunk = formulaItems.slice(fi, fi + FORMULA_CHUNK);
          await Promise.all(
            chunk.map(async (item) => {
              const { td, row, col, key, cellEl, cellRef, formulaText, cacheKey, fallbackStyleIndex } = item;
              let val: any;

              // Re-check shared cache
              if (cacheKey && sheetCache.get('__sharedValues')?.has(cacheKey)) {
                val = sheetCache.get('__sharedValues').get(cacheKey);
              } else {
                val = await formulaEngine.evaluateFormula(formulaText, {
                  resolveCell: async (r: string) => await cellEvaluator.evaluateCellByRef(r),
                  resolveCellsBatch: async (refs: string[]) => await cellEvaluator.evaluateCellsBatch(refs),
                  resolveRange: async (sn: string | undefined, sr: number, sc: number, er: number, ec: number) =>
                    await cellEvaluator.resolveRange(sn, sr, sc, er, ec),
                  sharedStrings,
                  zip,
                  rangeCache,
                  getSheetMaxRow: (sheetName?: string) => {
                    // Use the sheet cache to find actual data extent; avoids
                    // defaulting whole-column ranges to 5000 rows which causes
                    // millions of unnecessary cell iterations and GC pressure.
                    const name = sheetName || sheet.name;
                    const entry = sheetCache.get(name) as SheetCacheEntry | undefined;
                    if (entry && entry.maxRow && entry.maxRow > 0) return entry.maxRow;
                    return 200; // conservative fallback for not-yet-loaded sheets
                  },
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

              // Apply conditional formatting for this formula cell
              const cfEntries = cfMap.get(key) || [];
              if (cfEntries.length) {
                for (const entry of cfEntries) {
                  const result = await cfEvaluator.evaluate(entry, val, row, col);
                  if (result.matched) {
                    ConditionalFormatting.applyCfToTd(item.td, result);
                    if (entry.rule.stopIfTrue) break;
                  }
                }
              }
            })
          );          // Yield to the event loop periodically to keep UI responsive
          chunksSinceYield++;
          if (chunksSinceYield >= YIELD_EVERY_N_CHUNKS) {
            chunksSinceYield = 0;
            await yieldToEventLoop();
          }        }

        // CF-only items – batch them with periodic yields to avoid blocking
        // the main thread with 500+ concurrent promise chains.
        const CF_CHUNK = 32;
        let cfChunksSinceYield = 0;
        const CF_YIELD_EVERY = 4;
        for (let ci = 0; ci < cfOnlyItems.length; ci += CF_CHUNK) {
          const cfChunk = cfOnlyItems.slice(ci, ci + CF_CHUNK);
          await Promise.all(
            cfChunk.map(async (item) => {
              for (const entry of item.cfEntries) {
                const result = await cfEvaluator.evaluate(entry, item.value, item.row, item.col);
                if (result.matched) {
                  ConditionalFormatting.applyCfToTd(item.td, result);
                  if (entry.rule.stopIfTrue) break;
                }
              }
            })
          );
          cfChunksSinceYield++;
          if (cfChunksSinceYield >= CF_YIELD_EVERY) {
            cfChunksSinceYield = 0;
            await yieldToEventLoop();
          }
        }
      }

      tbody.appendChild(frag);
      // Only yield periodically on large sheets to keep the UI responsive
      // without incurring constant layout invalidation overhead.
      if (renderMaxRow > batchSize && startRow + batchSize <= renderMaxRow) {
        await yieldToEventLoop();
      }
    }

    table.appendChild(tbody);
    sheetNameEl.textContent = sheet.name;

    // Remove loading overlay and show table – use replaceChildren to do a
    // single atomic DOM swap instead of innerHTML = '' + appendChild which
    // causes two separate layout invalidations.
    tableContainer.replaceChildren(table);

    } catch (err) {
      // Ensure loading overlay is removed even on error
      console.error('SheetRenderer: error during rendering', err);
      tableContainer.replaceChildren(table);
    }

    // Render images
    if (images.length > 0) {
      ImageRenderer.renderImages(tableContainer, images);
    }

    // Render charts
    if (charts.length > 0) {
      ChartRenderer.renderCharts(tableContainer, charts);
    }

    // Release heavy XML DOM objects and cellMaps from the render-time cache.
    // CellStore already holds all extracted data; keeping these would waste
    // hundreds of MB on large workbooks.
    for (const [, entry] of sheetCache) {
      if (entry && typeof entry === 'object') {
        entry.cellMap = undefined;
        entry.sheetDoc = undefined;
      }
    }

    return table;
  }

  // ---- Private helpers ----

  /**
   * Load images and charts for a sheet in one pass, sharing the relationship
   * map so the .rels file is only read once.
   */
  private static async loadSheetAssets(
    zip: JSZip,
    target: string
  ): Promise<[ImageAnchor[], ChartAnchor[]]> {
    const sheetRels = await SheetParser.loadSheetRelationships(zip, target);
    const images: ImageAnchor[] = [];
    const charts: ChartAnchor[] = [];

    // Collect unique drawing paths
    const drawingPaths = new Set<string>();
    for (const [, relTarget] of sheetRels) {
      if (relTarget.includes('drawings/')) {
        drawingPaths.add(XmlParser.normalizeTargetPath(relTarget));
      }
    }

    // Process each drawing once for both images and charts
    for (const drawingPath of drawingPaths) {
      const [drawingImages, drawingCharts, drawingRels] = await Promise.all([
        SheetParser.loadDrawing(zip, drawingPath),
        SheetParser.loadDrawingCharts(zip, drawingPath),
        SheetParser.loadDrawingRelationships(zip, drawingPath),
      ]);

      // Load image media data — resolve embed IDs through drawing rels
      for (const image of drawingImages) {
        const relTarget = drawingRels.get(image.embed);
        if (!relTarget) continue;
        const mediaPath = XmlParser.normalizeTargetPath(relTarget);
        const mediaData = await SheetParser.loadMedia(zip, mediaPath);
        if (mediaData) {
          images.push({ ...image, dataUrl: mediaData });
        }
      }

      charts.push(...drawingCharts);
    }

    return [images, charts];
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
