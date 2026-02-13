// StyleParser.ts - Parses xl/styles.xml and xl/theme/theme1.xml

import type JSZip from 'jszip';
import type { FontInfo, CellXf, DxfStyle, FillInfo, BorderDef, BorderSide, StyleResult, ColorObject } from '../types';
import { XmlParser } from '../core/XmlParser';

/**
 * Parses all style components from a workbook ZIP.
 */
export class StyleParser {
  /**
   * Parse all styles (fonts, cellXfs, dxfs, fills, borders, themeColors).
   */
  static async parseStyles(zip: JSZip): Promise<StyleResult> {
    const fonts = await StyleParser.loadFonts(zip);
    const cellXfs = await StyleParser.loadCellXfs(zip);
    const cellStyleXfs = await StyleParser.loadCellStyleXfs(zip);
    const themeColors = await StyleParser.loadThemeColors(zip);
    const dxfs = await StyleParser.loadDxfs(zip);
    const fills = await StyleParser.loadFills(zip);
    const borders = await StyleParser.loadBorders(zip);
    return { fonts, cellXfs, cellStyleXfs, themeColors, dxfs, fills, borders };
  }

  static async loadFonts(zip: JSZip): Promise<FontInfo[]> {
    const stylesXml = zip.file('xl/styles.xml');
    if (!stylesXml) return [];
    const stylesText = await stylesXml.async('text');
    const stylesDoc = XmlParser.parseXml(stylesText);
    const fontNodes = Array.from(stylesDoc.getElementsByTagName('font'));
    return fontNodes.map((fn) => {
      const font: FontInfo = {};
      const sz = fn.getElementsByTagName('sz')[0];
      if (sz) font.pt = parseFloat(sz.getAttribute('val')!) || 11;
      const name = fn.getElementsByTagName('name')[0];
      if (name) font.name = name.getAttribute('val') || undefined;
      const color = fn.getElementsByTagName('color')[0];
      if (color) {
        font.rgb = color.getAttribute('rgb');
        font.theme = color.getAttribute('theme');
        font.indexed = color.getAttribute('indexed');
        font.tint = color.getAttribute('tint');
      }
      font.bold = !!fn.getElementsByTagName('b')[0];
      font.italic = !!fn.getElementsByTagName('i')[0];
      font.underline = !!fn.getElementsByTagName('u')[0];
      font.strike = !!fn.getElementsByTagName('strike')[0];
      return font;
    });
  }

  static async loadCellXfs(zip: JSZip): Promise<CellXf[]> {
    const stylesXml = zip.file('xl/styles.xml');
    if (!stylesXml) return [];
    const stylesText = await stylesXml.async('text');
    const stylesDoc = XmlParser.parseXml(stylesText);
    const cellXfsNode = stylesDoc.getElementsByTagName('cellXfs')[0];
    if (!cellXfsNode) return [];
    const xfNodes = Array.from(cellXfsNode.getElementsByTagName('xf'));
    return xfNodes.map((xf) => ({
      fontId: xf.getAttribute('fontId'),
      applyFont: xf.getAttribute('applyFont'),
      fillId: xf.getAttribute('fillId'),
      applyFill: xf.getAttribute('applyFill'),
      borderId: xf.getAttribute('borderId'),
      applyBorder: xf.getAttribute('applyBorder'),
      xfId: xf.getAttribute('xfId'),
    }));
  }

  static async loadCellStyleXfs(zip: JSZip): Promise<CellXf[]> {
    const stylesXml = zip.file('xl/styles.xml');
    if (!stylesXml) return [];
    const stylesText = await stylesXml.async('text');
    const stylesDoc = XmlParser.parseXml(stylesText);
    const cellStyleXfsNode = stylesDoc.getElementsByTagName('cellStyleXfs')[0];
    if (!cellStyleXfsNode) return [];
    const xfNodes = Array.from(cellStyleXfsNode.getElementsByTagName('xf'));
    return xfNodes.map((xf) => ({
      fontId: xf.getAttribute('fontId'),
      applyFont: xf.getAttribute('applyFont'),
      fillId: xf.getAttribute('fillId'),
      applyFill: xf.getAttribute('applyFill'),
      borderId: xf.getAttribute('borderId'),
      applyBorder: xf.getAttribute('applyBorder'),
    }));
  }

  static async loadDxfs(zip: JSZip): Promise<DxfStyle[]> {
    const stylesXml = zip.file('xl/styles.xml');
    if (!stylesXml) return [];
    const stylesText = await stylesXml.async('text');
    const stylesDoc = XmlParser.parseXml(stylesText);
    const dxfNodes = Array.from(stylesDoc.getElementsByTagName('dxf'));
    return dxfNodes.map((dxf) => {
      const fontNode = dxf.getElementsByTagName('font')[0];
      const fillNode = dxf.getElementsByTagName('fill')[0];
      const borderNode = dxf.getElementsByTagName('border')[0];
      const style: DxfStyle = {};

      if (fontNode) {
        const f: FontInfo = {};
        const sz = fontNode.getElementsByTagName('sz')[0];
        if (sz) f.pt = parseFloat(sz.getAttribute('val')!) || 11;
        const name = fontNode.getElementsByTagName('name')[0];
        if (name) f.name = name.getAttribute('val') || undefined;
        const color = fontNode.getElementsByTagName('color')[0];
        if (color) f.rgb = color.getAttribute('rgb');
        f.bold = !!fontNode.getElementsByTagName('b')[0];
        f.italic = !!fontNode.getElementsByTagName('i')[0];
        f.underline = !!fontNode.getElementsByTagName('u')[0];
        f.strike = !!fontNode.getElementsByTagName('strike')[0];
        style.font = f;
      }

      if (fillNode) {
        const patternFill = fillNode.getElementsByTagName('patternFill')[0];
        if (patternFill) {
          const fg = patternFill.getElementsByTagName('fgColor')[0];
          const bg = patternFill.getElementsByTagName('bgColor')[0];
          if (fg) {
            style.fill = {
              rgb: fg.getAttribute('rgb'),
              theme: fg.getAttribute('theme'),
              indexed: fg.getAttribute('indexed'),
              tint: fg.getAttribute('tint'),
            };
          } else if (bg) {
            style.fill = {
              rgb: bg.getAttribute('rgb'),
              theme: bg.getAttribute('theme'),
              indexed: bg.getAttribute('indexed'),
              tint: bg.getAttribute('tint'),
            };
          }
        }
      }

      if (borderNode) style.border = StyleParser.parseBorderNode(borderNode);
      return style;
    });
  }

  static async loadFills(zip: JSZip): Promise<FillInfo[]> {
    const stylesXml = zip.file('xl/styles.xml');
    if (!stylesXml) return [];
    const stylesText = await stylesXml.async('text');
    const stylesDoc = XmlParser.parseXml(stylesText);
    const fillNodes = Array.from(stylesDoc.getElementsByTagName('fill'));
    return fillNodes.map((fill) => {
      const patternFill = fill.getElementsByTagName('patternFill')[0];
      if (!patternFill) return {};
      const patternType = patternFill.getAttribute('patternType');
      const fg = patternFill.getElementsByTagName('fgColor')[0];
      const bg = patternFill.getElementsByTagName('bgColor')[0];
      const res: FillInfo = { patternType };
      if (fg) {
        res.fg = {
          rgb: fg.getAttribute('rgb'),
          theme: fg.getAttribute('theme'),
          indexed: fg.getAttribute('indexed'),
          tint: fg.getAttribute('tint'),
        };
      }
      if (bg) {
        res.bg = {
          rgb: bg.getAttribute('rgb'),
          theme: bg.getAttribute('theme'),
          indexed: bg.getAttribute('indexed'),
          tint: bg.getAttribute('tint'),
        };
      }
      return res;
    });
  }

  static async loadBorders(zip: JSZip): Promise<BorderDef[]> {
    const stylesXml = zip.file('xl/styles.xml');
    if (!stylesXml) return [];
    const stylesText = await stylesXml.async('text');
    const stylesDoc = XmlParser.parseXml(stylesText);
    const borderNodes = Array.from(stylesDoc.getElementsByTagName('border'));
    return borderNodes.map((b) => StyleParser.parseBorderNode(b));
  }

  static async loadThemeColors(zip: JSZip): Promise<Record<number, string>> {
    const themeXml = zip.file('xl/theme/theme1.xml');
    if (!themeXml) return {};
    const themeText = await themeXml.async('text');
    const themeDoc = XmlParser.parseXml(themeText);
    const colorMap: Record<number, string> = {};
    const clrScheme = Array.from(themeDoc.getElementsByTagName('*')).find(
      (n) => n.localName === 'clrScheme'
    );
    if (!clrScheme) return colorMap;

    const colorNames = [
      'lt1', 'dk1', 'lt2', 'dk2',
      'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
      'hyperlink', 'folHyperlink',
    ];

    colorNames.forEach((name, idx) => {
      const node = Array.from(clrScheme.childNodes).find(
        (n) => (n as Element).localName === name
      ) as Element | undefined;
      if (!node) return;
      const srgb = Array.from(node.childNodes).find(
        (n) => (n as Element).localName === 'srgbClr'
      ) as Element | undefined;
      if (srgb) {
        colorMap[idx] = `#${srgb.getAttribute('val')}`;
        return;
      }
      const sys = Array.from(node.childNodes).find(
        (n) => (n as Element).localName === 'sysClr'
      ) as Element | undefined;
      if (sys) {
        const lastClr = sys.getAttribute('lastClr');
        if (lastClr) colorMap[idx] = `#${lastClr}`;
      }
    });

    return colorMap;
  }

  /**
   * Parse a font from a run properties (rPr) element.
   */
  static parseFontFromRunPr(rPr: Element): FontInfo {
    const font: FontInfo = {};
    const sz = rPr.getElementsByTagName('sz')[0];
    if (sz) font.pt = parseFloat(sz.getAttribute('val')!) || 11;
    const name = rPr.getElementsByTagName('name')[0];
    if (name) font.name = name.getAttribute('val') || undefined;
    const color = rPr.getElementsByTagName('color')[0];
    if (color) {
      font.rgb = color.getAttribute('rgb');
      font.theme = color.getAttribute('theme');
      font.indexed = color.getAttribute('indexed');
      font.tint = color.getAttribute('tint');
    }
    font.bold = !!rPr.getElementsByTagName('b')[0];
    font.italic = !!rPr.getElementsByTagName('i')[0];
    font.underline = !!rPr.getElementsByTagName('u')[0];
    font.strike = !!rPr.getElementsByTagName('strike')[0];
    return font;
  }

  /**
   * Parse a `<border>` XML node into a BorderDef.
   */
  static parseBorderNode(borderNode: Element): BorderDef {
    const sides: BorderDef = {};
    (['left', 'right', 'top', 'bottom', 'diagonal'] as const).forEach((side) => {
      const node = borderNode.getElementsByTagName(side)[0];
      if (!node) return;
      const style = node.getAttribute('style');
      const colorNode = node.getElementsByTagName('color')[0];
      const sideObj: BorderSide = { style: style || null };
      if (colorNode) {
        sideObj.rgb = colorNode.getAttribute('rgb');
        sideObj.theme = colorNode.getAttribute('theme');
        sideObj.indexed = colorNode.getAttribute('indexed');
      }
      sides[side] = sideObj;
    });
    return sides;
  }
}
