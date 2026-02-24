// StyleApplicator.ts - Generate CSS objects from style definitions

import type {
  FontInfo, CellXf, DxfStyle, FillInfo, BorderDef, BorderSide,
  CssProperties, CellStyleResult, ColorObject,
} from '../types';
import { ColorResolver } from './ColorResolver';
import { StyleParser } from './StyleParser';

/**
 * Generates CSS property objects from parsed Excel style data and applies them to DOM elements.
 */
export class StyleApplicator {
  /**
   * Convert a FontInfo to a CssProperties object.
   */
  static fontToCss(font: FontInfo | null | undefined, themeColors: Record<number, string>): CssProperties {
    const css: CssProperties = {};
    if (!font) return css;
    if (font.name) css['fontFamily'] = font.name;
    if (font.pt) css['fontSize'] = `${(font.pt * 96) / 72}px`;
    if (font.bold) css['fontWeight'] = 'bold';
    if (font.italic) css['fontStyle'] = 'italic';
    const decorations: string[] = [];
    if (font.underline) decorations.push('underline');
    if (font.strike) decorations.push('line-through');
    if (decorations.length) css['textDecoration'] = decorations.join(' ');
    const fontColor = ColorResolver.resolveColor(
      { rgb: font.rgb, theme: font.theme, indexed: font.indexed, tint: font.tint },
      themeColors
    );
    if (fontColor) css.color = fontColor;
    return css;
  }

  /**
   * Convert a DxfStyle to a CssProperties object.
   */
  static dxfToCss(dxf: DxfStyle, themeColors: Record<number, string>): CssProperties {
    const css: CssProperties = {};
    if (dxf.font) Object.assign(css, StyleApplicator.fontToCss(dxf.font, themeColors));
    if (dxf.fill) {
      const resolved = ColorResolver.resolveColor(dxf.fill, themeColors);
      if (resolved) css.backgroundColor = resolved;
    }
    if (dxf.border) Object.assign(css, StyleApplicator.borderDefToCss(dxf.border, themeColors));
    return css;
  }

  /**
   * Convert border style strings to CSS pixel widths.
   */
  static borderStyleToPx(style: string | null | undefined): string {
    if (!style) return '1px';
    const s = style.toLowerCase();
    if (s === 'thin' || s === 'hair') return '1px';
    if (s === 'medium') return '2px';
    if (s === 'thick') return '3px';
    return '1px';
  }

  /**
   * Convert a border definition to CSS border properties.
   */
  static borderDefToCss(borderDef: BorderDef | null | undefined, themeColors: Record<number, string>): CssProperties {
    if (!borderDef) return {};
    const css: CssProperties = {};
    const sides: Array<'top' | 'right' | 'bottom' | 'left'> = ['top', 'right', 'bottom', 'left'];
    sides.forEach((side) => {
      const sideObj = borderDef[side];
      if (!sideObj) return;
      const px = StyleApplicator.borderStyleToPx(sideObj.style);
      let color = '#000';
      const resolved = ColorResolver.resolveColor(sideObj as ColorObject, themeColors);
      if (resolved) color = resolved;
      const cssKey = `border${side.charAt(0).toUpperCase()}${side.slice(1)}`;
      css[cssKey] = `${px} solid ${color}`;
    });
    return css;
  }

  /**
   * Cache for base style objects keyed by style index.
   * Avoids recomputing the same CSS properties for every cell that shares a style.
   */
  private static baseStyleCache = new Map<number, CssProperties>();

  /**
   * Clear the base-style cache (call when loading a new workbook).
   */
  static clearStyleCache(): void {
    StyleApplicator.baseStyleCache.clear();
  }

  /**
   * Get the full cell style (single value or rich-text runs) for a cell.
   */
  static getCellStyle(
    cellNode: Element | null,
    cellXfs: CellXf[],
    fonts: FontInfo[],
    themeColors: Record<number, string>,
    fills: FillInfo[],
    borders: BorderDef[],
    cellStyleXfs: CellXf[],
    fallbackStyleIndex?: string | null
  ): CellStyleResult {
    let styleIndex: number | null = null;

    if (cellNode?.getAttribute) {
      const s = cellNode.getAttribute('s');
      if (s != null) {
        const idx = parseInt(s, 10);
        if (!Number.isNaN(idx)) styleIndex = idx;
      }
    }

    if (styleIndex == null && fallbackStyleIndex != null) {
      const idx = parseInt(fallbackStyleIndex, 10);
      if (!Number.isNaN(idx)) styleIndex = idx;
    }

    // Use cached base style when the same style index is reused across cells
    let baseStyle: CssProperties;
    if (styleIndex != null && StyleApplicator.baseStyleCache.has(styleIndex)) {
      baseStyle = StyleApplicator.baseStyleCache.get(styleIndex)!;
    } else {
      baseStyle = {};
      if (styleIndex != null && cellXfs) {
        const xf = cellXfs[styleIndex];
        if (xf) {
          if (xf.xfId != null && cellStyleXfs) {
            const baseIdx = parseInt(xf.xfId, 10);
            if (!Number.isNaN(baseIdx) && cellStyleXfs[baseIdx]) {
              StyleApplicator.applyXfStyle(baseStyle, cellStyleXfs[baseIdx], fonts, themeColors, fills, borders);
            }
          }
          StyleApplicator.applyXfStyle(baseStyle, xf, fonts, themeColors, fills, borders);
        }
      }
      if (styleIndex != null) {
        StyleApplicator.baseStyleCache.set(styleIndex, baseStyle);
      }
    }

    if (!cellNode) return { type: 'single', style: baseStyle };

    const inlineNode = cellNode.getElementsByTagName('is')[0];
    if (inlineNode) {
      const runs = Array.from(inlineNode.getElementsByTagName('r')).map((r) => {
        const rPr = r.getElementsByTagName('rPr')[0];
        let runStyle: CssProperties = { ...baseStyle };
        if (rPr) {
          const runFont = StyleParser.parseFontFromRunPr(rPr);
          runStyle = StyleApplicator.fontToCss(runFont, themeColors || {});
        }
        const textNode = r.getElementsByTagName('t')[0];
        return { text: textNode ? textNode.textContent || '' : '', style: runStyle };
      });
      return { type: 'runs', runs, baseline: baseStyle };
    }

    return { type: 'single', style: baseStyle };
  }

  /**
   * Convert a camelCase CSS property name to kebab-case.
   */
  private static toKebab(prop: string): string {
    return prop.replace(/[A-Z]/g, (m) => '-' + m.toLowerCase());
  }

  /**
   * Apply CSS properties object to an element's inline style.
   * Uses cssText for bulk application to avoid repeated style recalculations.
   */
  static applyCssToElement(element: HTMLElement, cssObject: CssProperties | null | undefined): void {
    if (!cssObject) return;
    const entries = Object.entries(cssObject);
    if (entries.length === 0) return;

    // Build a single cssText string and apply at once to minimise style recalcs
    const parts: string[] = [];
    if (element.style.cssText) parts.push(element.style.cssText);
    for (const [key, value] of entries) {
      if (value != null) parts.push(`${StyleApplicator.toKebab(key)}:${value}`);
    }
    element.style.cssText = parts.join(';');
  }

  // ---- Private helpers ----

  private static shouldApply(xf: CellXf | null | undefined, key: string): boolean {
    if (!xf) return false;
    const raw = (xf as any)[key];
    if (raw == null) return true;
    return raw === '1';
  }

  private static applyXfStyle(
    target: CssProperties,
    xf: CellXf,
    fonts: FontInfo[],
    themeColors: Record<number, string>,
    fills: FillInfo[],
    borders: BorderDef[]
  ): void {
    if (!xf) return;

    if (xf.fontId != null && StyleApplicator.shouldApply(xf, 'applyFont')) {
      const fontIdx = parseInt(xf.fontId, 10);
      if (!Number.isNaN(fontIdx) && fonts && fonts[fontIdx]) {
        Object.assign(target, StyleApplicator.fontToCss(fonts[fontIdx], themeColors || {}));
      }
    }

    if (xf.fillId != null && StyleApplicator.shouldApply(xf, 'applyFill') && fills) {
      const fid = parseInt(xf.fillId, 10);
      if (!Number.isNaN(fid) && fills[fid]) {
        const f = fills[fid];
        if (f && f.patternType !== 'none' && f.fg) {
          const resolved = ColorResolver.resolveColor(f.fg, themeColors);
          if (resolved) target.backgroundColor = resolved;
        }
      }
    }

    if (xf.borderId != null && StyleApplicator.shouldApply(xf, 'applyBorder') && borders) {
      const bid = parseInt(xf.borderId, 10);
      if (!Number.isNaN(bid) && borders[bid]) {
        const borderDef = borders[bid];
        const sides: Array<'top' | 'right' | 'bottom' | 'left'> = ['top', 'right', 'bottom', 'left'];
        sides.forEach((side) => {
          const sideObj = borderDef[side];
          if (!sideObj) return;
          const px = StyleApplicator.borderStyleToPx(sideObj.style);
          let color = '#000';
          const resolved = ColorResolver.resolveColor(sideObj as ColorObject, themeColors);
          if (resolved) color = resolved;
          const cssKey = `border${side.charAt(0).toUpperCase()}${side.slice(1)}`;
          target[cssKey] = `${px} solid ${color}`;
        });
      }
    }
  }
}
