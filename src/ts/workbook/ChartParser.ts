// ChartParser.ts - Parses Excel chart XML (xl/charts/chartN.xml)

import type JSZip from 'jszip';
import type { ChartData, ChartType, ChartSeries, ChartSeriesPoint, ChartAxisInfo } from '../types';
import { XmlParser } from '../core/XmlParser';

/**
 * Default color palette used when the chart XML does not specify series colors.
 * Based on a standard Excel-like palette.
 */
const DEFAULT_COLORS = [
  '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5',
  '#70AD47', '#264478', '#9B57A0', '#636363', '#EB5757',
  '#00B0F0', '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0',
  '#9966FF', '#FF9F40', '#C9CBCF',
];

// Tag names used across chart XMLs for the different chart types
const CHART_TYPE_TAGS: Array<{ tag: string; type: ChartType }> = [
  { tag: 'c:barChart', type: 'bar' },
  { tag: 'c:bar3DChart', type: 'bar' },
  { tag: 'c:lineChart', type: 'line' },
  { tag: 'c:line3DChart', type: 'line' },
  { tag: 'c:pieChart', type: 'pie' },
  { tag: 'c:pie3DChart', type: 'pie' },
  { tag: 'c:areaChart', type: 'area' },
  { tag: 'c:area3DChart', type: 'area' },
  { tag: 'c:scatterChart', type: 'scatter' },
  { tag: 'c:doughnutChart', type: 'doughnut' },
  { tag: 'c:radarChart', type: 'radar' },
  { tag: 'c:ofPieChart', type: 'pie' },
];

/**
 * Parses chart XML files from an XLSX ZIP archive.
 */
export class ChartParser {
  /**
   * Parse a chart XML file and return its data.
   */
  static async parseChart(zip: JSZip, chartPath: string): Promise<ChartData | null> {
    try {
      const doc = await XmlParser.readZipXml(zip, chartPath);
      return ChartParser.parseChartDocument(doc);
    } catch {
      return null;
    }
  }

  /**
   * Parse a chart XML DOM document.
   */
  static parseChartDocument(doc: Document): ChartData {
    const chartEl = doc.getElementsByTagName('c:chart')[0];
    const title = ChartParser.parseChartTitle(chartEl || doc);
    const plotArea = chartEl
      ? chartEl.getElementsByTagName('c:plotArea')[0]
      : doc.getElementsByTagName('c:plotArea')[0];

    let chartType: ChartType = 'unknown';
    let chartTypeEl: Element | null = null;

    if (plotArea) {
      for (const mapping of CHART_TYPE_TAGS) {
        const els = plotArea.getElementsByTagName(mapping.tag);
        if (els.length > 0) {
          chartType = mapping.type;
          chartTypeEl = els[0];
          break;
        }
      }
    }

    // Detect bar direction (horizontal bar vs vertical column)
    if (chartType === 'bar' && chartTypeEl) {
      const barDirEl = chartTypeEl.getElementsByTagName('c:barDir')[0];
      const barDir = barDirEl?.getAttribute('val') || 'col';
      chartType = barDir === 'bar' ? 'bar' : 'col';
    }

    const series = chartTypeEl
      ? ChartParser.parseSeries(chartTypeEl, chartType)
      : [];

    // Parse axes
    const categoryAxis = ChartParser.parseAxis(plotArea, 'c:catAx');
    const valueAxis = ChartParser.parseAxis(plotArea, 'c:valAx');

    // Check if we used cached values
    const usesCache = series.some(s => s.points.length > 0);

    return {
      type: chartType,
      title,
      series,
      categoryAxis: categoryAxis || undefined,
      valueAxis: valueAxis || undefined,
      usesCache,
    };
  }

  /**
   * Extract the chart title text.
   */
  private static parseChartTitle(chartEl: Element): string {
    const titleEl = chartEl.getElementsByTagName('c:title')[0];
    if (!titleEl) return '';

    // Look for rich text in <c:tx><c:rich><a:p><a:r><a:t>
    const txEl = titleEl.getElementsByTagName('c:tx')[0];
    if (txEl) {
      const richEl = txEl.getElementsByTagName('c:rich')[0];
      if (richEl) {
        const textEls = richEl.getElementsByTagName('a:t');
        const parts: string[] = [];
        for (let i = 0; i < textEls.length; i += 1) {
          parts.push(textEls[i].textContent || '');
        }
        return parts.join('').trim();
      }
      // Fallback: string reference
      const strRefEl = txEl.getElementsByTagName('c:strRef')[0];
      if (strRefEl) {
        const cached = ChartParser.extractCachedStrings(strRefEl);
        if (cached.length > 0) return cached[0];
      }
    }

    return '';
  }

  /**
   * Parse all <c:ser> elements inside a chart type element.
   */
  private static parseSeries(chartTypeEl: Element, chartType: ChartType): ChartSeries[] {
    const serEls = chartTypeEl.getElementsByTagName('c:ser');
    const result: ChartSeries[] = [];

    for (let i = 0; i < serEls.length; i += 1) {
      // Only consider direct children of chartTypeEl (not nested)
      const serEl = serEls[i];
      if (serEl.parentElement !== chartTypeEl) continue;

      const name = ChartParser.parseSeriesName(serEl, i);
      const color = ChartParser.parseSeriesColor(serEl) || DEFAULT_COLORS[i % DEFAULT_COLORS.length];
      const categories = ChartParser.parseCategories(serEl, chartType);
      const values = ChartParser.parseValues(serEl, chartType);

      const points: ChartSeriesPoint[] = [];
      const maxLen = Math.max(categories.length, values.length);
      for (let j = 0; j < maxLen; j += 1) {
        points.push({
          category: j < categories.length ? categories[j] : `${j + 1}`,
          value: j < values.length ? values[j] : 0,
        });
      }

      result.push({ name, points, color });
    }

    return result;
  }

  /**
   * Extract the series name from <c:tx>.
   */
  private static parseSeriesName(serEl: Element, index: number): string {
    const txEl = serEl.getElementsByTagName('c:tx')[0];
    if (txEl) {
      // Try cached string reference first
      const strRefEl = txEl.getElementsByTagName('c:strRef')[0];
      if (strRefEl) {
        const cached = ChartParser.extractCachedStrings(strRefEl);
        if (cached.length > 0) return cached[0];
      }
      // Direct value
      const vEl = txEl.getElementsByTagName('c:v')[0];
      if (vEl) return vEl.textContent || `Series ${index + 1}`;
    }
    return `Series ${index + 1}`;
  }

  /**
   * Extract the series fill color from <c:spPr>.
   */
  private static parseSeriesColor(serEl: Element): string | undefined {
    const spPrEl = serEl.getElementsByTagName('c:spPr')[0];
    if (!spPrEl) return undefined;

    // Try <a:solidFill><a:srgbClr val="..." />
    const solidFill = spPrEl.getElementsByTagName('a:solidFill')[0];
    if (solidFill) {
      const srgb = solidFill.getElementsByTagName('a:srgbClr')[0];
      if (srgb) {
        const val = srgb.getAttribute('val');
        if (val) return `#${val}`;
      }
    }

    return undefined;
  }

  /**
   * Parse category labels from <c:cat> (string or numeric ref with cache).
   */
  private static parseCategories(serEl: Element, chartType: ChartType): string[] {
    // For scatter charts, categories are X values
    if (chartType === 'scatter') {
      const xValEl = serEl.getElementsByTagName('c:xVal')[0];
      if (xValEl) {
        const numRef = xValEl.getElementsByTagName('c:numRef')[0];
        if (numRef) return ChartParser.extractCachedNumbers(numRef).map(String);
        const numLit = xValEl.getElementsByTagName('c:numLit')[0];
        if (numLit) return ChartParser.extractLiteralNumbers(numLit).map(String);
      }
      return [];
    }

    const catEl = serEl.getElementsByTagName('c:cat')[0];
    if (!catEl) return [];

    // String reference with cache
    const strRefEl = catEl.getElementsByTagName('c:strRef')[0];
    if (strRefEl) return ChartParser.extractCachedStrings(strRefEl);

    // Numeric reference with cache
    const numRefEl = catEl.getElementsByTagName('c:numRef')[0];
    if (numRefEl) return ChartParser.extractCachedNumbers(numRefEl).map(String);

    // String literal
    const strLitEl = catEl.getElementsByTagName('c:strLit')[0];
    if (strLitEl) return ChartParser.extractLiteralStrings(strLitEl);

    // Numeric literal
    const numLitEl = catEl.getElementsByTagName('c:numLit')[0];
    if (numLitEl) return ChartParser.extractLiteralNumbers(numLitEl).map(String);

    return [];
  }

  /**
   * Parse values from <c:val> or <c:yVal> (for scatter charts).
   */
  private static parseValues(serEl: Element, chartType: ChartType): number[] {
    const valTag = chartType === 'scatter' ? 'c:yVal' : 'c:val';
    const valEl = serEl.getElementsByTagName(valTag)[0];
    if (!valEl) return [];

    // Numeric reference with cache
    const numRefEl = valEl.getElementsByTagName('c:numRef')[0];
    if (numRefEl) return ChartParser.extractCachedNumbers(numRefEl);

    // Numeric literal
    const numLitEl = valEl.getElementsByTagName('c:numLit')[0];
    if (numLitEl) return ChartParser.extractLiteralNumbers(numLitEl);

    return [];
  }

  /**
   * Extract cached string values from a <c:strRef> element.
   * Reads from <c:strCache><c:pt idx="N"><c:v>...</c:v></c:pt></c:strCache>.
   */
  private static extractCachedStrings(refEl: Element): string[] {
    const cache = refEl.getElementsByTagName('c:strCache')[0];
    if (!cache) return [];
    const pts = cache.getElementsByTagName('c:pt');
    const result: Array<{ idx: number; val: string }> = [];
    for (let i = 0; i < pts.length; i += 1) {
      const idx = parseInt(pts[i].getAttribute('idx') || '0', 10);
      const vEl = pts[i].getElementsByTagName('c:v')[0];
      result.push({ idx, val: vEl?.textContent || '' });
    }
    result.sort((a, b) => a.idx - b.idx);
    return result.map(r => r.val);
  }

  /**
   * Extract cached numeric values from a <c:numRef> element.
   * Reads from <c:numCache><c:pt idx="N"><c:v>...</c:v></c:pt></c:numCache>.
   */
  private static extractCachedNumbers(refEl: Element): number[] {
    const cache = refEl.getElementsByTagName('c:numCache')[0];
    if (!cache) return [];
    const pts = cache.getElementsByTagName('c:pt');
    const result: Array<{ idx: number; val: number }> = [];
    for (let i = 0; i < pts.length; i += 1) {
      const idx = parseInt(pts[i].getAttribute('idx') || '0', 10);
      const vEl = pts[i].getElementsByTagName('c:v')[0];
      const num = parseFloat(vEl?.textContent || '0');
      result.push({ idx, val: Number.isNaN(num) ? 0 : num });
    }
    result.sort((a, b) => a.idx - b.idx);
    return result.map(r => r.val);
  }

  /**
   * Extract literal string values from <c:strLit>.
   */
  private static extractLiteralStrings(litEl: Element): string[] {
    const pts = litEl.getElementsByTagName('c:pt');
    const result: Array<{ idx: number; val: string }> = [];
    for (let i = 0; i < pts.length; i += 1) {
      const idx = parseInt(pts[i].getAttribute('idx') || '0', 10);
      const vEl = pts[i].getElementsByTagName('c:v')[0];
      result.push({ idx, val: vEl?.textContent || '' });
    }
    result.sort((a, b) => a.idx - b.idx);
    return result.map(r => r.val);
  }

  /**
   * Extract literal numeric values from <c:numLit>.
   */
  private static extractLiteralNumbers(litEl: Element): number[] {
    const pts = litEl.getElementsByTagName('c:pt');
    const result: Array<{ idx: number; val: number }> = [];
    for (let i = 0; i < pts.length; i += 1) {
      const idx = parseInt(pts[i].getAttribute('idx') || '0', 10);
      const vEl = pts[i].getElementsByTagName('c:v')[0];
      const num = parseFloat(vEl?.textContent || '0');
      result.push({ idx, val: Number.isNaN(num) ? 0 : num });
    }
    result.sort((a, b) => a.idx - b.idx);
    return result.map(r => r.val);
  }

  /**
   * Parse an axis element (catAx or valAx).
   */
  private static parseAxis(plotArea: Element | null, tagName: string): ChartAxisInfo | null {
    if (!plotArea) return null;
    const axEl = plotArea.getElementsByTagName(tagName)[0];
    if (!axEl) return null;

    const axIdEl = axEl.getElementsByTagName('c:axId')[0];
    const id = axIdEl?.getAttribute('val') || '';

    let title = '';
    const titleEl = axEl.getElementsByTagName('c:title')[0];
    if (titleEl) {
      const textEls = titleEl.getElementsByTagName('a:t');
      const parts: string[] = [];
      for (let i = 0; i < textEls.length; i += 1) {
        parts.push(textEls[i].textContent || '');
      }
      title = parts.join('').trim();
    }

    return { title, id };
  }

  /**
   * Parse chart anchors from a drawing XML document.
   * Looks for <xdr:graphicFrame> elements that reference charts.
   */
  static parseChartAnchors(drawingDoc: Document): Array<{ from: { col: number; row: number }; to: { col: number; row: number }; relId: string }> {
    const anchors: Array<{ from: { col: number; row: number }; to: { col: number; row: number }; relId: string }> = [];

    const processTwoCellAnchors = (anchorEls: Element[]) => {
      for (const anchorEl of anchorEls) {
        const graphicFrame = anchorEl.getElementsByTagName('xdr:graphicFrame')[0];
        if (!graphicFrame) continue;

        const graphic = graphicFrame.getElementsByTagName('a:graphic')[0];
        if (!graphic) continue;

        const graphicData = graphic.getElementsByTagName('a:graphicData')[0];
        if (!graphicData) continue;

        const uri = graphicData.getAttribute('uri') || '';
        if (!uri.includes('chart')) continue;

        // Get the chart relationship ID
        const chartRef = graphicData.getElementsByTagName('c:chart')[0];
        if (!chartRef) continue;

        const relId = chartRef.getAttribute('r:id') || '';
        if (!relId) continue;

        // Parse anchor positions
        const fromEl = anchorEl.getElementsByTagName('xdr:from')[0];
        const toEl = anchorEl.getElementsByTagName('xdr:to')[0];

        const from = { col: 0, row: 0 };
        const to = { col: 1, row: 1 };

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

        anchors.push({ from, to, relId });
      }
    };

    const twoCellAnchors = Array.from(drawingDoc.getElementsByTagName('xdr:twoCellAnchor'));
    processTwoCellAnchors(twoCellAnchors);

    return anchors;
  }
}
