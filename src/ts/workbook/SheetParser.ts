// SheetParser.ts - Parses sheet-level data structures

import type JSZip from 'jszip';
import type {
  SheetModel, ColumnStyleRange, ImageAnchor, ChartAnchor,
  SharedFormulaEntry, SharedCellEntry, RangeRef,
  MergeAnchor, CellRef,
} from '../types';
import { XmlParser } from '../core/XmlParser';
import { CellReference } from '../core/CellReference';
import { ChartParser } from './ChartParser';

/**
 * Parses structural data from a sheet XML document: merge cells,
 * shared formulas, column styles, drawings, and images.
 */
export class SheetParser {
  /**
   * Build a complete SheetModel from a sheet XML document.
   */
  static buildSheetModel(sheetDoc: Document): SheetModel {
    const mergedRanges = SheetParser.parseMergeCells(sheetDoc);
    const { anchorMap, coveredMap } = SheetParser.buildMergedMaps(mergedRanges);
    const { sharedFormulas, sharedCells } = SheetParser.parseSharedFormulas(sheetDoc);
    return { mergedRanges, anchorMap, coveredMap, sharedFormulas, sharedCells };
  }

  /**
   * Parse merge cell ranges from the sheet document.
   */
  static parseMergeCells(sheetDoc: Document): RangeRef[] {
    const mergeCellsNode = sheetDoc.getElementsByTagName('mergeCells')[0];
    if (!mergeCellsNode) return [];
    const mergeNodes = Array.from(mergeCellsNode.getElementsByTagName('mergeCell'));
    const ranges: RangeRef[] = [];
    mergeNodes.forEach((node) => {
      const sqref = node.getAttribute('ref') || node.getAttribute('sqref');
      if (!sqref) return;
      const parts = sqref.trim().split(/\s+/);
      parts.forEach((part) => {
        const { start, end } = CellReference.parseRangeRef(part);
        if (!start.row || !start.col || !end.row || !end.col) return;
        ranges.push({ start, end });
      });
    });
    return ranges;
  }

  /**
   * Build anchorMap and coveredMap from merge ranges.
   */
  static buildMergedMaps(mergedRanges: RangeRef[]): {
    anchorMap: Map<string, MergeAnchor>;
    coveredMap: Map<string, string>;
  } {
    const anchorMap = new Map<string, MergeAnchor>();
    const coveredMap = new Map<string, string>();
    mergedRanges.forEach(({ start, end }) => {
      const rowSpan = end.row - start.row + 1;
      const colSpan = end.col - start.col + 1;
      const anchorKey = CellReference.cellKey(start.row, start.col);
      anchorMap.set(anchorKey, { rowSpan, colSpan });
      for (let r = start.row; r <= end.row; r += 1) {
        for (let c = start.col; c <= end.col; c += 1) {
          if (r === start.row && c === start.col) continue;
          coveredMap.set(CellReference.cellKey(r, c), anchorKey);
        }
      }
    });
    return { anchorMap, coveredMap };
  }

  /**
   * Parse shared formulas and their cell associations.
   */
  static parseSharedFormulas(sheetDoc: Document): {
    sharedFormulas: Map<string, SharedFormulaEntry>;
    sharedCells: Map<string, SharedCellEntry>;
  } {
    const sharedFormulas = new Map<string, SharedFormulaEntry>();
    const sharedCells = new Map<string, SharedCellEntry>();
    const fNodes = Array.from(sheetDoc.getElementsByTagName('f'));

    fNodes.forEach((fNode) => {
      if (fNode.getAttribute('t') !== 'shared') return;
      const si = fNode.getAttribute('si');
      if (!si) return;
      const parent = fNode.parentNode as Element | null;
      if (!parent || parent.nodeName !== 'c') return;
      const cellRef = parent.getAttribute('r');
      if (!cellRef) return;
      const text = fNode.textContent || '';
      if (text && !sharedFormulas.has(si)) {
        sharedFormulas.set(si, { si, anchorRef: cellRef, baseFormula: text });
      }
      sharedCells.set(cellRef, { si });
    });

    sharedCells.forEach((info, cellRef) => {
      const entry = sharedFormulas.get(info.si);
      if (!entry) return;
      sharedCells.set(cellRef, { ...info, anchorRef: entry.anchorRef, baseFormula: entry.baseFormula });
    });

    return { sharedFormulas, sharedCells };
  }

  /**
   * Parse column style definitions from the sheet document.
   */
  static parseColumnStyles(sheetDoc: Document): ColumnStyleRange[] {
    const colsNode = sheetDoc.getElementsByTagName('cols')[0];
    if (!colsNode) return [];
    const colNodes = Array.from(colsNode.getElementsByTagName('col'));
    const ranges: ColumnStyleRange[] = [];
    colNodes.forEach((col) => {
      const min = parseInt(col.getAttribute('min') || '0', 10);
      const max = parseInt(col.getAttribute('max') || '0', 10);
      const style = col.getAttribute('style');
      if (!min || !max || style == null) return;
      ranges.push({ min, max, style });
    });
    return ranges;
  }

  /**
   * Load sheet-level relationships.
   */
  static async loadSheetRelationships(zip: JSZip, sheetTarget: string): Promise<Map<string, string>> {
    const relsPath = sheetTarget.replace('.xml', '.xml.rels');
    try {
      const relsDoc = await XmlParser.readZipXml(zip, relsPath);
      return XmlParser.buildRelationshipMap(relsDoc);
    } catch {
      return new Map();
    }
  }

  /**
   * Load drawing anchors from a drawing XML document.
   */
  static async loadDrawing(zip: JSZip, drawingTarget: string): Promise<ImageAnchor[]> {
    try {
      const drawingDoc = await XmlParser.readZipXml(zip, drawingTarget);
      return SheetParser.parseDrawing(drawingDoc);
    } catch {
      return [];
    }
  }

  /**
   * Load a media file as a data URL.
   * @param mediaPath - fully resolved path inside the zip, e.g. "xl/media/image1.png"
   */
  static async loadMedia(zip: JSZip, mediaPath: string): Promise<string | null> {
    try {
      const mediaFile = zip.file(mediaPath);
      if (!mediaFile) return null;
      const data = await mediaFile.async('base64');
      const ext = mediaPath.split('.').pop()?.toLowerCase() || 'png';
      const mimeMap: Record<string, string> = {
        png: 'image/png',
        jpg: 'image/jpeg',
        jpeg: 'image/jpeg',
        gif: 'image/gif',
        bmp: 'image/bmp',
        svg: 'image/svg+xml',
        webp: 'image/webp',
        emf: 'image/x-emf',
        wmf: 'image/x-wmf',
        tiff: 'image/tiff',
        tif: 'image/tiff',
      };
      const mime = mimeMap[ext] || 'image/png';
      return `data:${mime};base64,${data}`;
    } catch {
      return null;
    }
  }

  /**
   * Load the relationship map for a drawing XML file.
   */
  static async loadDrawingRelationships(
    zip: JSZip,
    drawingTarget: string
  ): Promise<Map<string, string>> {
    const relsPath = drawingTarget.replace(/([^/]+)$/, '_rels/$1.rels');
    try {
      const relsDoc = await XmlParser.readZipXml(zip, relsPath);
      return XmlParser.buildRelationshipMap(relsDoc);
    } catch {
      return new Map();
    }
  }

  /**
   * Extract cell value from a cell element.
   */
  static extractCellValue(cell: Element, sharedStrings: string[]): string {
    const type = cell.getAttribute('t');
    if (type === 'inlineStr') {
      const inlineText = cell.getElementsByTagName('t')[0];
      return inlineText ? inlineText.textContent || '' : '';
    }
    const valueNode = cell.getElementsByTagName('v')[0];
    if (!valueNode) return '';
    const rawValue = valueNode.textContent || '';
    if (type === 's') {
      const index = parseInt(rawValue, 10);
      return Number.isNaN(index) ? rawValue : sharedStrings[index] || '';
    }
    if (type === 'b') return rawValue === '1' ? 'TRUE' : 'FALSE';
    return rawValue;
  }

  /**
   * Evaluate a simple numeric formula (no cell refs).
   */
  static evalSimpleFormula(formula: string): number | null {
    if (!formula) return null;
    const f = formula.trim();
    const cleaned = f.replace(/^\(|\)$/g, '');
    const num = Number(cleaned);
    return Number.isNaN(num) ? null : num;
  }

  // ---- Private helpers ----

  /**
   * Load charts referenced by a drawing XML file.
   * Parses drawing rels for chart relationships and loads each chart XML.
   */
  static async loadDrawingCharts(zip: JSZip, drawingTarget: string): Promise<ChartAnchor[]> {
    try {
      const drawingDoc = await XmlParser.readZipXml(zip, drawingTarget);

      // Parse chart anchors from drawing
      const chartAnchors = ChartParser.parseChartAnchors(drawingDoc);
      if (chartAnchors.length === 0) return [];

      // Load drawing rels to resolve chart relationship IDs
      const relsPath = drawingTarget.replace(/([^/]+)$/, '_rels/$1.rels');
      let drawingRels = new Map<string, string>();
      try {
        const relsDoc = await XmlParser.readZipXml(zip, relsPath);
        drawingRels = XmlParser.buildRelationshipMap(relsDoc);
      } catch {
        return [];
      }

      const charts: ChartAnchor[] = [];
      for (const anchor of chartAnchors) {
        const chartTarget = drawingRels.get(anchor.relId);
        if (!chartTarget) continue;

        const chartPath = XmlParser.normalizeTargetPath(chartTarget);
        const chartData = await ChartParser.parseChart(zip, chartPath);
        if (!chartData) continue;

        charts.push({
          from: anchor.from,
          to: anchor.to,
          relId: anchor.relId,
          chartPath,
          chartData,
        });
      }

      return charts;
    } catch {
      return [];
    }
  }

  private static parseDrawing(drawingDoc: Document): ImageAnchor[] {
    const images: ImageAnchor[] = [];
    const twoCellAnchors = Array.from(drawingDoc.getElementsByTagName('xdr:twoCellAnchor'));
    const oneCellAnchors = Array.from(drawingDoc.getElementsByTagName('xdr:oneCellAnchor'));

    const parseAnchor = (anchorEl: Element): ImageAnchor | null => {
      const fromEl = anchorEl.getElementsByTagName('xdr:from')[0];
      const toEl = anchorEl.getElementsByTagName('xdr:to')[0];
      const picEl = anchorEl.getElementsByTagName('xdr:pic')[0];
      if (!picEl) return null;

      const blipEl = picEl.getElementsByTagName('a:blip')[0];
      if (!blipEl) return null;

      const embed = blipEl.getAttribute('r:embed');
      if (!embed) return null;

      const from: CellRef = { col: 0, row: 0 };
      const to: CellRef = { col: 1, row: 1 };

      if (fromEl) {
        const colEl = fromEl.getElementsByTagName('xdr:col')[0];
        const rowEl = fromEl.getElementsByTagName('xdr:row')[0];
        from.col = parseInt(colEl?.textContent || '0', 10) + 1;
        from.row = parseInt(rowEl?.textContent || '0', 10) + 1;
      }

      if (toEl) {
        const colEl = toEl.getElementsByTagName('xdr:col')[0];
        const rowEl = toEl.getElementsByTagName('xdr:row')[0];
        to.col = parseInt(colEl?.textContent || '0', 10) + 1;
        to.row = parseInt(rowEl?.textContent || '0', 10) + 1;
      }

      return { embed, from, to, type: 'twoCellAnchor' };
    };

    twoCellAnchors.forEach((anchor) => {
      const image = parseAnchor(anchor);
      if (image) images.push(image);
    });

    oneCellAnchors.forEach((anchor) => {
      const image = parseAnchor(anchor);
      if (image) {
        image.type = 'oneCellAnchor';
        images.push(image);
      }
    });

    return images;
  }
}
