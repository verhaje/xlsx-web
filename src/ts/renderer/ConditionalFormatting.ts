// ConditionalFormatting.ts - Conditional formatting evaluation

import type {
  CfRule, CfEntry, CfResult, CfvoInfo, CfColorInfo,
  ColorScaleInfo, DataBarInfo, IconSetInfo,
  StyleResult, SheetModel, CellRef, CssProperties,
  IFormulaEngine,
} from '../types';
import { CellReference } from '../core/CellReference';
import { SheetParser } from '../workbook/SheetParser';
import { StyleApplicator } from '../styles/StyleApplicator';
import { ColorResolver } from '../styles/ColorResolver';
import { FormulaShifter } from './FormulaShifter';

// Icon set metadata
const ICON_SET_META: Record<string, Array<{ icon: string; color: string }>> = {
  '3Arrows': [
    { icon: '\u25BC', color: '#d93025' },
    { icon: '\u25B6', color: '#f29900' },
    { icon: '\u25B2', color: '#188038' },
  ],
  '3ArrowsGray': [
    { icon: '\u25BC', color: '#6e6e6e' },
    { icon: '\u25B6', color: '#9e9e9e' },
    { icon: '\u25B2', color: '#4a4a4a' },
  ],
  '3TrafficLights1': [
    { icon: '\u25CF', color: '#d93025' },
    { icon: '\u25CF', color: '#f29900' },
    { icon: '\u25CF', color: '#188038' },
  ],
  '3TrafficLights2': [
    { icon: '\u25CF', color: '#d93025' },
    { icon: '\u25CF', color: '#f29900' },
    { icon: '\u25CF', color: '#188038' },
  ],
  '3Signs': [
    { icon: '\u25A0', color: '#d93025' },
    { icon: '\u25A0', color: '#f29900' },
    { icon: '\u25A0', color: '#188038' },
  ],
  '3Symbols': [
    { icon: '\u2716', color: '#d93025' },
    { icon: '\u25CF', color: '#f29900' },
    { icon: '\u2714', color: '#188038' },
  ],
  '3Symbols2': [
    { icon: '\u2716', color: '#d93025' },
    { icon: '\u25CF', color: '#f29900' },
    { icon: '\u2714', color: '#188038' },
  ],
};

/**
 * Parses and evaluates conditional formatting rules for sheet cells.
 */
export class ConditionalFormatting {
  private styles: StyleResult;
  private evaluateCellByRef: (ref: string, contextSheetName?: string) => Promise<any>;
  private formulaEngine: IFormulaEngine | null;
  private sharedStrings: string[];
  private zip: any;
  private sheetDoc: Document;

  constructor(
    styles: StyleResult,
    evaluateCellByRef: (ref: string, contextSheetName?: string) => Promise<any>,
    formulaEngine: IFormulaEngine | null,
    sharedStrings: string[],
    zip: any,
    sheetDoc: Document
  ) {
    this.styles = styles;
    this.evaluateCellByRef = evaluateCellByRef;
    this.formulaEngine = formulaEngine;
    this.sharedStrings = sharedStrings;
    this.zip = zip;
    this.sheetDoc = sheetDoc;
  }

  /**
   * Parse conditional formatting rules from the sheet document into a per-cell map.
   */
  static parseCfRules(
    sheetDoc: Document,
    styles: StyleResult
  ): Map<string, CfEntry[]> {
    const cfMap = new Map<string, CfEntry[]>();
    let cfRuleOrder = 0;

    const parseCfvo = (node: Element | null): CfvoInfo | null => {
      if (!node) return null;
      return {
        type: node.getAttribute('type') || 'num',
        val: node.getAttribute('val'),
        gte: node.getAttribute('gte') !== '0',
      };
    };

    const parseColorNode = (node: Element | null): CfColorInfo | null => {
      if (!node) return null;
      return {
        rgb: node.getAttribute('rgb'),
        theme: node.getAttribute('theme'),
        tint: node.getAttribute('tint'),
      };
    };

    const cfNodes = Array.from(sheetDoc.getElementsByTagName('conditionalFormatting'));
    cfNodes.forEach((cfNode) => {
      const sqref = cfNode.getAttribute('sqref');
      if (!sqref) return;
      const rangeParts = sqref.trim().split(/\s+/);
      const cfRules = Array.from(cfNode.getElementsByTagName('cfRule'));

      cfRules.forEach((rule) => {
        const dxfId = rule.getAttribute('dxfId');
        const type = rule.getAttribute('type');
        const operator = rule.getAttribute('operator');
        const formulaNodes = Array.from(rule.getElementsByTagName('formula'));
        const formula = (formulaNodes[0] || {} as any).textContent || null;
        const formula1 = (rule.getElementsByTagName('formula1')[0] || {} as any).textContent || null;
        const formula2 = (rule.getElementsByTagName('formula2')[0] || {} as any).textContent || null;
        const priorityRaw = rule.getAttribute('priority');
        const priorityParsed = priorityRaw == null ? Number.MAX_SAFE_INTEGER : parseInt(priorityRaw, 10);
        const priority = Number.isNaN(priorityParsed) ? Number.MAX_SAFE_INTEGER : priorityParsed;
        const stopIfTrue = rule.getAttribute('stopIfTrue') === '1' || rule.getAttribute('stopIfTrue') === 'true';
        const text = rule.getAttribute('text');
        const top = rule.getAttribute('top');
        const bottom = rule.getAttribute('bottom');
        const percent = rule.getAttribute('percent');
        const rank = rule.getAttribute('rank');
        const aboveAverage = rule.getAttribute('aboveAverage');
        const equalAverage = rule.getAttribute('equalAverage');

        const colorScaleNode = rule.getElementsByTagName('colorScale')[0];
        const dataBarNode = rule.getElementsByTagName('dataBar')[0];
        const iconSetNode = rule.getElementsByTagName('iconSet')[0];

        const colorScale: ColorScaleInfo | null = colorScaleNode
          ? {
              cfvos: Array.from(colorScaleNode.getElementsByTagName('cfvo')).map(parseCfvo).filter(Boolean) as CfvoInfo[],
              colors: Array.from(colorScaleNode.getElementsByTagName('color')).map(parseColorNode).filter(Boolean) as CfColorInfo[],
            }
          : null;

        const dataBar: DataBarInfo | null = dataBarNode
          ? {
              cfvos: Array.from(dataBarNode.getElementsByTagName('cfvo')).map(parseCfvo).filter(Boolean) as CfvoInfo[],
              color: parseColorNode(dataBarNode.getElementsByTagName('color')[0]),
              showValue: dataBarNode.getAttribute('showValue') !== '0',
            }
          : null;

        const iconSet: IconSetInfo | null = iconSetNode
          ? {
              cfvos: Array.from(iconSetNode.getElementsByTagName('cfvo')).map(parseCfvo).filter(Boolean) as CfvoInfo[],
              iconSet: iconSetNode.getAttribute('iconSet') || '3TrafficLights1',
              showValue: iconSetNode.getAttribute('showValue') !== '0',
            }
          : null;

        let css: CssProperties | null = null;
        if (dxfId != null && styles.dxfs && styles.dxfs[parseInt(dxfId, 10)]) {
          css = StyleApplicator.dxfToCss(styles.dxfs[parseInt(dxfId, 10)], styles.themeColors);
        }
        const hasVisualRule = !!(colorScale || dataBar || iconSet);
        if (!css && !hasVisualRule) return;

        const ruleInfo: CfRule = {
          type,
          operator,
          formula,
          formula1,
          formula2,
          formulas: formulaNodes.map((node) => node.textContent || '').filter(Boolean),
          css,
          priority,
          stopIfTrue,
          text,
          top,
          bottom,
          percent,
          rank,
          aboveAverage,
          equalAverage,
          colorScale,
          dataBar,
          iconSet,
          order: cfRuleOrder,
          targets: new Map(),
        };
        cfRuleOrder += 1;

        rangeParts.forEach((part) => {
          const { start, end } = CellReference.parseRangeRef(part);
          if (!start.row || !start.col || !end.row || !end.col) return;
          const targetCells = CellReference.expandRange(part);
          targetCells.forEach(({ row, col }) => {
            const key = `${row}-${col}`;
            if (!ruleInfo.targets.has(key)) ruleInfo.targets.set(key, { row, col, key });
            if (!cfMap.has(key)) cfMap.set(key, []);
            cfMap.get(key)!.push({ rule: ruleInfo, anchor: start });
          });
        });
      });
    });

    // Sort by priority then order
    cfMap.forEach((entries) => {
      entries.sort((a, b) => {
        if (a.rule.priority !== b.rule.priority) return a.rule.priority - b.rule.priority;
        return a.rule.order - b.rule.order;
      });
    });

    return cfMap;
  }

  /**
   * Evaluate a single conditional formatting entry for a cell.
   */
  async evaluate(entry: CfEntry, cellValue: any, row: number, col: number): Promise<CfResult> {
    const { rule, anchor } = entry;
    if (!rule) return { matched: false };

    if (rule.type === 'cellIs') {
      return this.evaluateCellIs(rule, cellValue, row, col, anchor);
    }
    if (
      rule.type === 'containsText' ||
      rule.type === 'notContainsText' ||
      rule.type === 'beginsWith' ||
      rule.type === 'endsWith'
    ) {
      return this.evaluateTextRule(rule, cellValue);
    }
    if (rule.type === 'expression') {
      const expr = rule.formula || rule.formula1 || (rule.formulas && rule.formulas[0]) || null;
      const result = await this.resolveCfValue(expr, row, col, anchor);
      return { matched: ConditionalFormatting.isTruthyCfValue(result), css: rule.css || undefined };
    }
    if (rule.type === 'top10') {
      return this.evaluateTop10(rule, cellValue);
    }
    if (rule.type === 'aboveAverage') {
      return this.evaluateAboveAverage(rule, cellValue);
    }
    if (rule.type === 'colorScale') {
      const css = await this.evaluateColorScale(rule, cellValue, row, col, anchor);
      return css ? { matched: true, css } : { matched: false };
    }
    if (rule.type === 'dataBar') {
      const result = await this.evaluateDataBar(rule, cellValue, row, col, anchor);
      return result ? { matched: true, css: result.css, hideValue: result.hideValue } : { matched: false };
    }
    if (rule.type === 'iconSet') {
      const result = await this.evaluateIconSet(rule, cellValue, row, col, anchor);
      return result ? { matched: true, icon: result.icon, hideValue: result.hideValue } : { matched: false };
    }
    return { matched: false };
  }

  /**
   * Apply conditional formatting results to a table cell element.
   */
  static applyCfToTd(
    td: HTMLTableCellElement,
    result: CfResult
  ): void {
    if (result.css) StyleApplicator.applyCssToElement(td, result.css);
    if (result.icon) {
      const existing = td.querySelector('.cf-icon');
      if (existing) existing.remove();
      const iconSpan = document.createElement('span');
      iconSpan.className = 'cf-icon';
      iconSpan.textContent = result.icon.icon;
      iconSpan.style.marginRight = '4px';
      iconSpan.style.fontSize = '12px';
      iconSpan.style.color = result.icon.color || '#000';
      iconSpan.style.verticalAlign = 'middle';
      td.insertBefore(iconSpan, td.firstChild);
    }
    if (result.hideValue) {
      td.style.color = 'transparent';
      const children = Array.from(td.querySelectorAll('span'));
      children.forEach((span) => {
        if (!span.classList.contains('cf-icon')) (span as HTMLElement).style.color = 'transparent';
      });
    }
  }

  // ---- Private evaluation helpers ----

  private async evaluateCellIs(
    rule: CfRule,
    cellValue: any,
    row: number,
    col: number,
    anchor: CellRef
  ): Promise<CfResult> {
    const v1 = await this.resolveCfValue(rule.formula1 || rule.formula, row, col, anchor);
    const v2 = await this.resolveCfValue(rule.formula2, row, col, anchor);
    const num = Number(cellValue);
    const n1 = Number(v1);
    const n2 = Number(v2);
    if (Number.isNaN(num) || Number.isNaN(n1)) {
      const s = String(cellValue ?? '');
      const s1 = String(v1 ?? '');
      if (rule.operator === 'equal') return { matched: s === s1, css: rule.css || undefined };
      if (rule.operator === 'notEqual') return { matched: s !== s1, css: rule.css || undefined };
      return { matched: false };
    }
    switch (rule.operator) {
      case 'greaterThan':        return { matched: num > n1, css: rule.css || undefined };
      case 'lessThan':           return { matched: num < n1, css: rule.css || undefined };
      case 'greaterThanOrEqual': return { matched: num >= n1, css: rule.css || undefined };
      case 'lessThanOrEqual':    return { matched: num <= n1, css: rule.css || undefined };
      case 'equal':              return { matched: num === n1, css: rule.css || undefined };
      case 'notEqual':           return { matched: num !== n1, css: rule.css || undefined };
      case 'between':
        return { matched: !Number.isNaN(n2) && num >= n1 && num <= n2, css: rule.css || undefined };
      default:
        return { matched: false };
    }
  }

  private evaluateTextRule(rule: CfRule, cellValue: any): CfResult {
    const txt = rule.text != null ? rule.text : (rule.formula || '');
    const valueText = String(cellValue ?? '');
    let matched = false;
    if (rule.type === 'beginsWith') matched = valueText.startsWith(txt);
    else if (rule.type === 'endsWith') matched = valueText.endsWith(txt);
    else matched = valueText.indexOf(txt) !== -1;
    return { matched: rule.type === 'notContainsText' ? !matched : matched, css: rule.css || undefined };
  }

  private async evaluateTop10(rule: CfRule, cellValue: any): Promise<CfResult> {
    const numeric = await this.getRuleNumericValues(rule);
    if (!numeric.length) return { matched: false };
    const rankRaw = parseInt(rule.rank || '10', 10);
    const rank = Number.isNaN(rankRaw) ? 10 : rankRaw;
    const isPercent = rule.percent === '1';
    const isBottom = rule.bottom === '1' || rule.top === '0';
    let count = isPercent ? Math.ceil((numeric.length * rank) / 100) : rank;
    count = Math.max(1, Math.min(count, numeric.length));
    const sorted = [...numeric].sort((a, b) => (isBottom ? a - b : b - a));
    const threshold = sorted[count - 1];
    const num = Number(cellValue);
    if (Number.isNaN(num)) return { matched: false };
    return { matched: isBottom ? num <= threshold : num >= threshold, css: rule.css || undefined };
  }

  private async evaluateAboveAverage(rule: CfRule, cellValue: any): Promise<CfResult> {
    const numeric = await this.getRuleNumericValues(rule);
    if (!numeric.length) return { matched: false };
    const avg = numeric.reduce((sum, v) => sum + v, 0) / numeric.length;
    const num = Number(cellValue);
    if (Number.isNaN(num)) return { matched: false };
    const above = rule.aboveAverage !== '0';
    const equal = rule.equalAverage === '1';
    if (above) return { matched: equal ? num >= avg : num > avg, css: rule.css || undefined };
    return { matched: equal ? num <= avg : num < avg, css: rule.css || undefined };
  }

  private async evaluateColorScale(
    rule: CfRule,
    cellValue: any,
    row: number,
    col: number,
    anchor: CellRef
  ): Promise<CssProperties | null> {
    if (!rule.colorScale) return null;
    const num = Number(cellValue);
    if (Number.isNaN(num)) return null;
    const colors = rule.colorScale.colors || [];
    const cfvos = rule.colorScale.cfvos || [];
    if (colors.length < 2 || cfvos.length < 2) return null;
    const values = await this.resolveCfvoValues(cfvos, rule, row, col, anchor);
    if (values.some((v) => Number.isNaN(v))) return null;
    const resolvedColors = colors
      .map((c) => this.resolveThemeColor(c))
      .filter(Boolean) as string[];
    if (resolvedColors.length < 2) return null;
    const min = values[0];
    const max = values[values.length - 1];
    if (max === min) return { backgroundColor: resolvedColors[resolvedColors.length - 1] };
    let color = resolvedColors[resolvedColors.length - 1];
    if (values.length >= 3 && resolvedColors.length >= 3) {
      const mid = values[1];
      if (num <= mid) {
        const t = ConditionalFormatting.clamp((num - min) / (mid - min), 0, 1);
        color = ConditionalFormatting.interpolateColor(resolvedColors[0], resolvedColors[1], t) || resolvedColors[1];
      } else {
        const t = ConditionalFormatting.clamp((num - mid) / (max - mid), 0, 1);
        color = ConditionalFormatting.interpolateColor(resolvedColors[1], resolvedColors[2], t) || resolvedColors[2];
      }
    } else {
      const t = ConditionalFormatting.clamp((num - min) / (max - min), 0, 1);
      color = ConditionalFormatting.interpolateColor(resolvedColors[0], resolvedColors[1], t) || resolvedColors[1];
    }
    return { backgroundColor: color };
  }

  private async evaluateDataBar(
    rule: CfRule,
    cellValue: any,
    row: number,
    col: number,
    anchor: CellRef
  ): Promise<{ css: CssProperties; hideValue: boolean } | null> {
    if (!rule.dataBar) return null;
    const num = Number(cellValue);
    if (Number.isNaN(num)) return null;
    const cfvos = rule.dataBar.cfvos || [];
    if (cfvos.length < 2) return null;
    const values = await this.resolveCfvoValues(cfvos, rule, row, col, anchor);
    if (values.some((v) => Number.isNaN(v))) return null;
    const min = values[0];
    const max = values[values.length - 1];
    const color = this.resolveThemeColor(rule.dataBar.color) || '#638ec6';
    const percent =
      max === min ? 100 : ConditionalFormatting.clamp(((num - min) / (max - min)) * 100, 0, 100);
    return {
      css: {
        backgroundImage: `linear-gradient(90deg, ${color} ${percent}%, transparent ${percent}%)`,
        backgroundRepeat: 'no-repeat',
        backgroundSize: '100% 100%',
      },
      hideValue: !rule.dataBar.showValue,
    };
  }

  private async evaluateIconSet(
    rule: CfRule,
    cellValue: any,
    row: number,
    col: number,
    anchor: CellRef
  ): Promise<{ icon: { icon: string; color: string }; hideValue: boolean } | null> {
    if (!rule.iconSet) return null;
    const num = Number(cellValue);
    if (Number.isNaN(num)) return null;
    const cfvos = rule.iconSet.cfvos || [];
    if (!cfvos.length) return null;
    const values = await this.resolveCfvoValues(cfvos, rule, row, col, anchor);
    if (values.some((v) => Number.isNaN(v))) return null;
    const icons = ICON_SET_META[rule.iconSet.iconSet] || ICON_SET_META['3TrafficLights1'];
    let idx = 0;
    for (let i = values.length - 1; i >= 0; i -= 1) {
      const gte = cfvos[i]?.gte !== false;
      if (gte ? num >= values[i] : num > values[i]) {
        idx = i;
        break;
      }
    }
    const icon = icons[Math.min(idx, icons.length - 1)] || icons[0];
    return { icon, hideValue: !rule.iconSet.showValue };
  }

  // ---- Value resolution helpers ----

  private async resolveCfValue(
    raw: string | null,
    row: number,
    col: number,
    anchor: CellRef
  ): Promise<any> {
    const normalized = ConditionalFormatting.normalizeCfFormula(raw);
    if (!normalized) return null;
    const simple = SheetParser.evalSimpleFormula(normalized);
    if (simple != null) return simple;
    if (!this.formulaEngine) return simple;
    const rowOffset = row - anchor.row;
    const colOffset = col - anchor.col;
    const shifted = FormulaShifter.shiftFormulaRefs(normalized, rowOffset, colOffset);
    return this.formulaEngine.evaluateFormula(shifted, {
      resolveCell: async (r: string) => await this.evaluateCellByRef(r),
      sharedStrings: this.sharedStrings,
      zip: this.zip,
      sheetDoc: this.sheetDoc,
    });
  }

  private async resolveCfvoValue(
    cfvo: CfvoInfo | null,
    rule: CfRule,
    row: number,
    col: number,
    anchor: CellRef
  ): Promise<number | null> {
    if (!cfvo) return null;
    const type = cfvo.type || 'num';
    const numeric = await this.getRuleNumericValues(rule);
    const min = numeric.length ? Math.min(...numeric) : 0;
    const max = numeric.length ? Math.max(...numeric) : 0;
    if (type === 'min') return min;
    if (type === 'max') return max;
    if (type === 'percent') {
      const p = Number(cfvo.val || 0) / 100;
      return min + (max - min) * p;
    }
    if (type === 'percentile') {
      const p = Number(cfvo.val || 0) / 100;
      if (!numeric.length) return 0;
      const sorted = [...numeric].sort((a, b) => a - b);
      const idx = ConditionalFormatting.clamp(
        Math.ceil((sorted.length - 1) * p),
        0,
        sorted.length - 1
      );
      return sorted[idx];
    }
    if (type === 'formula') return this.resolveCfValue(cfvo.val, row, col, anchor);
    return Number(cfvo.val);
  }

  private async resolveCfvoValues(
    cfvos: CfvoInfo[],
    rule: CfRule,
    row: number,
    col: number,
    anchor: CellRef
  ): Promise<number[]> {
    const values: number[] = [];
    for (const cfvo of cfvos) {
      const v = await this.resolveCfvoValue(cfvo, rule, row, col, anchor);
      values.push(Number(v));
    }
    return values;
  }

  private async getRuleNumericValues(rule: CfRule): Promise<number[]> {
    if (!rule) return [];
    if (rule._numericValues) return rule._numericValues;
    const cells = await ConditionalFormatting.getRuleTargetCells(rule);
    const numeric: number[] = [];
    for (const cell of cells) {
      const a1 = CellReference.toA1(cell.row, cell.col);
      const val = await this.evaluateCellByRef(a1);
      const num = Number(val);
      if (!Number.isNaN(num)) numeric.push(num);
    }
    rule._numericValues = numeric;
    return numeric;
  }

  private resolveThemeColor(color: CfColorInfo | null): string | null {
    if (!color) return null;
    if (color.rgb) return ColorResolver.normalizeRgb(color.rgb) || null;
    if (color.theme != null) {
      const idx = parseInt(color.theme, 10);
      if (!Number.isNaN(idx) && this.styles.themeColors && this.styles.themeColors[idx]) {
        return this.styles.themeColors[idx];
      }
    }
    return null;
  }

  // ---- Static utility helpers ----

  private static getRuleTargetCells(rule: CfRule): Array<{ row: number; col: number; key: string }> {
    if (!rule || !rule.targets) return [];
    if (rule._targetCells) return rule._targetCells;
    const cells = Array.from(rule.targets.values());
    rule._targetCells = cells;
    return cells;
  }

  private static normalizeCfFormula(raw: string | null): string | null {
    if (raw == null) return null;
    const text = String(raw).trim();
    if (!text) return null;
    return text.replace(/^=/, '');
  }

  static isTruthyCfValue(value: any): boolean {
    if (value == null) return false;
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value !== 0;
    const text = String(value).trim();
    if (!text) return false;
    const upper = text.toUpperCase();
    if (upper === 'TRUE') return true;
    if (upper === 'FALSE') return false;
    const num = Number(text);
    if (!Number.isNaN(num)) return num !== 0;
    return true;
  }

  private static clamp(value: number, min: number, max: number): number {
    if (value < min) return min;
    if (value > max) return max;
    return value;
  }

  private static interpolateColor(hexA: string, hexB: string, t: number): string | null {
    if (!hexA || !hexB) return null;
    const parse = (hex: string): [number, number, number] => {
      const cleaned = hex.replace('#', '');
      return [
        parseInt(cleaned.slice(0, 2), 16),
        parseInt(cleaned.slice(2, 4), 16),
        parseInt(cleaned.slice(4, 6), 16),
      ];
    };
    const [r1, g1, b1] = parse(hexA);
    const [r2, g2, b2] = parse(hexB);
    const r = Math.round(r1 + (r2 - r1) * t);
    const g = Math.round(g1 + (g2 - g1) * t);
    const b = Math.round(b1 + (b2 - b1) * t);
    return `#${[r, g, b].map((v) => v.toString(16).padStart(2, '0')).join('')}`;
  }
}
