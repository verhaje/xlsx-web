// styles.js - parse xl/styles.xml and theme, map to CSS
import { parseXml } from './parser.js';

// Standard Excel indexed color palette (BIFF8)
const INDEXED_COLORS = [
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF', // 0-7
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF', // 8-15
  '800000','008000','000080','808000','800080','008080','C0C0C0','808080', // 16-23
  '9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF', // 24-31
  '000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF', // 32-39
  '00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99', // 40-47
  '3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696', // 48-55
  '003366','339966','003300','333300','993300','993366','333399','333333', // 56-63
];

function applyTint(hex, tint) {
  if (!tint) return `#${hex}`;
  const t = parseFloat(tint);
  if (Number.isNaN(t)) return `#${hex}`;
  let r = parseInt(hex.substring(0, 2), 16);
  let g = parseInt(hex.substring(2, 4), 16);
  let b = parseInt(hex.substring(4, 6), 16);
  if (t < 0) {
    r = Math.round(r * (1 + t));
    g = Math.round(g * (1 + t));
    b = Math.round(b * (1 + t));
  } else {
    r = Math.round(r * (1 - t) + 255 * t);
    g = Math.round(g * (1 - t) + 255 * t);
    b = Math.round(b * (1 - t) + 255 * t);
  }
  const clamp = (v) => Math.max(0, Math.min(255, v));
  return `#${clamp(r).toString(16).padStart(2, '0')}${clamp(g).toString(16).padStart(2, '0')}${clamp(b).toString(16).padStart(2, '0')}`;
}

export function resolveColor(colorObj, themeColors) {
  if (!colorObj) return undefined;
  if (colorObj.rgb) {
    const base = normalizeRgb(colorObj.rgb);
    return colorObj.tint ? applyTint(base.replace('#', ''), colorObj.tint) : base;
  }
  if (colorObj.theme != null) {
    const idx = parseInt(colorObj.theme, 10);
    if (!Number.isNaN(idx) && themeColors && themeColors[idx]) {
      const base = themeColors[idx].replace('#', '');
      return colorObj.tint ? applyTint(base, colorObj.tint) : `#${base}`;
    }
  }
  if (colorObj.indexed != null) {
    const idx = parseInt(colorObj.indexed, 10);
    if (!Number.isNaN(idx) && idx >= 0 && idx < INDEXED_COLORS.length) {
      return `#${INDEXED_COLORS[idx]}`;
    }
  }
  return undefined;
}

export async function parseStyles(zip) {
  const fonts = await loadFonts(zip);
  const cellXfs = await loadCellXfs(zip);
  const cellStyleXfs = await loadCellStyleXfs(zip);
  const themeColors = await loadThemeColors(zip);
  const dxfs = await loadDxfs(zip);
  const fills = await loadFills(zip);
  const borders = await loadBorders(zip);
  return { fonts, cellXfs, cellStyleXfs, themeColors, dxfs, fills, borders };
}

export async function loadFonts(zip) {
  const stylesXml = zip.file('xl/styles.xml');
  if (!stylesXml) return [];
  const stylesText = await stylesXml.async('text');
  const stylesDoc = parseXml(stylesText);
  const fontNodes = Array.from(stylesDoc.getElementsByTagName('font'));
  return fontNodes.map((fn) => {
    const font = {};
    const sz = fn.getElementsByTagName('sz')[0]; if (sz) font.pt = parseFloat(sz.getAttribute('val')) || 11;
    const name = fn.getElementsByTagName('name')[0]; if (name) font.name = name.getAttribute('val');
    const color = fn.getElementsByTagName('color')[0]; if (color) {
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

export async function loadCellXfs(zip) {
  const stylesXml = zip.file('xl/styles.xml');
  if (!stylesXml) return [];
  const stylesText = await stylesXml.async('text');
  const stylesDoc = parseXml(stylesText);
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
    xfId: xf.getAttribute('xfId')
  }));
}

export async function loadCellStyleXfs(zip) {
  const stylesXml = zip.file('xl/styles.xml');
  if (!stylesXml) return [];
  const stylesText = await stylesXml.async('text');
  const stylesDoc = parseXml(stylesText);
  const cellStyleXfsNode = stylesDoc.getElementsByTagName('cellStyleXfs')[0];
  if (!cellStyleXfsNode) return [];
  const xfNodes = Array.from(cellStyleXfsNode.getElementsByTagName('xf'));
  return xfNodes.map((xf) => ({
    fontId: xf.getAttribute('fontId'),
    applyFont: xf.getAttribute('applyFont'),
    fillId: xf.getAttribute('fillId'),
    applyFill: xf.getAttribute('applyFill'),
    borderId: xf.getAttribute('borderId'),
    applyBorder: xf.getAttribute('applyBorder')
  }));
}

export async function loadDxfs(zip) {
  const stylesXml = zip.file('xl/styles.xml');
  if (!stylesXml) return [];
  const stylesText = await stylesXml.async('text');
  const stylesDoc = parseXml(stylesText);
  const dxfNodes = Array.from(stylesDoc.getElementsByTagName('dxf'));
  return dxfNodes.map((dxf) => {
    const fontNode = dxf.getElementsByTagName('font')[0];
    const fillNode = dxf.getElementsByTagName('fill')[0];
    const borderNode = dxf.getElementsByTagName('border')[0];
    const style = {};
    if (fontNode) {
      const f = {};
      const sz = fontNode.getElementsByTagName('sz')[0]; if (sz) f.pt = parseFloat(sz.getAttribute('val')) || 11;
      const name = fontNode.getElementsByTagName('name')[0]; if (name) f.name = name.getAttribute('val');
      const color = fontNode.getElementsByTagName('color')[0]; if (color) f.rgb = color.getAttribute('rgb');
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
        if (fg) style.fill = { rgb: fg.getAttribute('rgb'), theme: fg.getAttribute('theme'), indexed: fg.getAttribute('indexed'), tint: fg.getAttribute('tint') };
        else if (bg) style.fill = { rgb: bg.getAttribute('rgb'), theme: bg.getAttribute('theme'), indexed: bg.getAttribute('indexed'), tint: bg.getAttribute('tint') };
      }
    }
    if (borderNode) style.border = parseBorderNode(borderNode);
    return style;
  });
}

export async function loadFills(zip) {
  const stylesXml = zip.file('xl/styles.xml');
  if (!stylesXml) return [];
  const stylesText = await stylesXml.async('text');
  const stylesDoc = parseXml(stylesText);
  const fillNodes = Array.from(stylesDoc.getElementsByTagName('fill'));
  return fillNodes.map((fill) => {
    const patternFill = fill.getElementsByTagName('patternFill')[0];
    if (!patternFill) return {};
    const patternType = patternFill.getAttribute('patternType');
    const fg = patternFill.getElementsByTagName('fgColor')[0];
    const bg = patternFill.getElementsByTagName('bgColor')[0];
    const res = { patternType };
    if (fg) res.fg = { rgb: fg.getAttribute('rgb'), theme: fg.getAttribute('theme'), indexed: fg.getAttribute('indexed'), tint: fg.getAttribute('tint') };
    if (bg) res.bg = { rgb: bg.getAttribute('rgb'), theme: bg.getAttribute('theme'), indexed: bg.getAttribute('indexed'), tint: bg.getAttribute('tint') };
    return res;
  });
}

export async function loadBorders(zip) {
  const stylesXml = zip.file('xl/styles.xml');
  if (!stylesXml) return [];
  const stylesText = await stylesXml.async('text');
  const stylesDoc = parseXml(stylesText);
  const borderNodes = Array.from(stylesDoc.getElementsByTagName('border'));
  return borderNodes.map((b) => parseBorderNode(b));
}

export async function loadThemeColors(zip) {
  const themeXml = zip.file('xl/theme/theme1.xml');
  if (!themeXml) return {};
  const themeText = await themeXml.async('text');
  const themeDoc = parseXml(themeText);
  const colorMap = {};
  const clrScheme = Array.from(themeDoc.getElementsByTagName('*')).find((n) => n.localName === 'clrScheme');
  if (!clrScheme) return colorMap;
  const colorNames = ['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hyperlink','folHyperlink'];
  colorNames.forEach((name, idx) => {
    const node = Array.from(clrScheme.childNodes).find(n => n.localName === name);
    if (!node) return;
    const srgb = Array.from(node.childNodes).find(n => n.localName === 'srgbClr');
    if (srgb) { colorMap[idx] = `#${srgb.getAttribute('val')}`; return; }
    const sys = Array.from(node.childNodes).find(n => n.localName === 'sysClr');
    if (sys) {
      const lastClr = sys.getAttribute('lastClr');
      if (lastClr) colorMap[idx] = `#${lastClr}`;
    }
  });
  return colorMap;
}

export function normalizeRgb(raw) {
  if (!raw) return undefined;
  const s = raw.replace(/^0x/i, '').replace(/^#/, '');
  if (s.length === 8) return `#${s.slice(2)}`;
  if (s.length === 6) return `#${s}`;
  return undefined;
}

export function fontToCss(font, themeColors) {
  const css = {};
  if (!font) return css;
  if (font.name) css['fontFamily'] = font.name;
  if (font.pt) css['fontSize'] = `${(font.pt * 96) / 72}px`;
  if (font.bold) css['fontWeight'] = 'bold';
  if (font.italic) css['fontStyle'] = 'italic';
  const decorations = [];
  if (font.underline) decorations.push('underline');
  if (font.strike) decorations.push('line-through');
  if (decorations.length) css['textDecoration'] = decorations.join(' ');
  const fontColor = resolveColor({ rgb: font.rgb, theme: font.theme, indexed: font.indexed, tint: font.tint }, themeColors);
  if (fontColor) css.color = fontColor;
  return css;
}

export function dxfToCss(dxf, themeColors) {
  const css = {};
  if (dxf.font) Object.assign(css, fontToCss(dxf.font, themeColors));
  if (dxf.fill) {
    const resolved = resolveColor(dxf.fill, themeColors);
    if (resolved) css.backgroundColor = resolved;
  }
  if (dxf.border) Object.assign(css, borderDefToCss(dxf.border, themeColors));
  return css;
}

export function parseFontFromRunPr(rPr) {
  const font = {};
  const sz = rPr.getElementsByTagName('sz')[0]; if (sz) font.pt = parseFloat(sz.getAttribute('val')) || 11;
  const name = rPr.getElementsByTagName('name')[0]; if (name) font.name = name.getAttribute('val');
  const color = rPr.getElementsByTagName('color')[0]; if (color) { font.rgb = color.getAttribute('rgb'); font.theme = color.getAttribute('theme'); font.indexed = color.getAttribute('indexed'); font.tint = color.getAttribute('tint'); }
  font.bold = !!rPr.getElementsByTagName('b')[0];
  font.italic = !!rPr.getElementsByTagName('i')[0];
  font.underline = !!rPr.getElementsByTagName('u')[0];
  font.strike = !!rPr.getElementsByTagName('strike')[0];
  return font;
}

function borderStyleToPx(style) {
  if (!style) return '1px';
  const s = style.toLowerCase();
  if (s === 'thin' || s === 'hair') return '1px';
  if (s === 'medium') return '2px';
  if (s === 'thick') return '3px';
  return '1px';
}

function parseBorderNode(borderNode) {
  const sides = {};
  ['left', 'right', 'top', 'bottom', 'diagonal'].forEach((side) => {
    const node = borderNode.getElementsByTagName(side)[0];
    if (!node) return;
    const style = node.getAttribute('style');
    const colorNode = node.getElementsByTagName('color')[0];
    const sideObj = { style: style || null };
    if (colorNode) {
      sideObj.rgb = colorNode.getAttribute('rgb');
      sideObj.theme = colorNode.getAttribute('theme');
      sideObj.indexed = colorNode.getAttribute('indexed');
    }
    sides[side] = sideObj;
  });
  return sides;
}

function borderDefToCss(borderDef, themeColors) {
  if (!borderDef) return {};
  const css = {};
  const sides = ['top', 'right', 'bottom', 'left'];
  sides.forEach((side) => {
    const sideObj = borderDef[side];
    if (!sideObj) return;
    const px = borderStyleToPx(sideObj.style);
    let color = '#000';
    const resolved = resolveColor(sideObj, themeColors);
    if (resolved) color = resolved;
    const cssKey = `border${side.charAt(0).toUpperCase()}${side.slice(1)}`;
    css[cssKey] = `${px} solid ${color}`;
  });
  return css;
}

function shouldApply(xf, key) {
  if (!xf) return false;
  const raw = xf[key];
  if (raw == null) return true;
  return raw === '1';
}

function applyXfStyle(target, xf, fonts, themeColors, fills, borders) {
  if (!xf) return;

  if (xf.fontId != null && shouldApply(xf, 'applyFont')) {
    const fontIdx = parseInt(xf.fontId, 10);
    if (!Number.isNaN(fontIdx) && fonts && fonts[fontIdx]) {
      Object.assign(target, fontToCss(fonts[fontIdx], themeColors || {}));
    }
  }

  if (xf.fillId != null && shouldApply(xf, 'applyFill') && fills) {
    const fid = parseInt(xf.fillId, 10);
    if (!Number.isNaN(fid) && fills[fid]) {
      const f = fills[fid];
      if (f && f.patternType !== 'none' && f.fg) {
        const resolved = resolveColor(f.fg, themeColors);
        if (resolved) target.backgroundColor = resolved;
      }
    }
  }

  if (xf.borderId != null && shouldApply(xf, 'applyBorder') && borders) {
    const bid = parseInt(xf.borderId, 10);
    if (!Number.isNaN(bid) && borders[bid]) {
      const borderDef = borders[bid];
      const sides = ['top', 'right', 'bottom', 'left'];
      sides.forEach((side) => {
        const sideObj = borderDef[side];
        if (!sideObj) return;
        const px = borderStyleToPx(sideObj.style);
        let color = '#000';
        const resolved = resolveColor(sideObj, themeColors);
        if (resolved) color = resolved;
        const cssKey = `border${side.charAt(0).toUpperCase()}${side.slice(1)}`;
        target[cssKey] = `${px} solid ${color}`;
      });
    }
  }
}

export function getCellStyle(cellNode, cellXfs, fonts, themeColors, fills, borders, cellStyleXfs, fallbackStyleIndex) {
  let baseStyle = {};
  let styleIndex = null;

  if (cellNode && cellNode.getAttribute) {
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

  if (styleIndex != null && cellXfs) {
    const xf = cellXfs[styleIndex];
    if (xf) {
      if (xf.xfId != null && cellStyleXfs) {
        const baseIdx = parseInt(xf.xfId, 10);
        if (!Number.isNaN(baseIdx) && cellStyleXfs[baseIdx]) {
          applyXfStyle(baseStyle, cellStyleXfs[baseIdx], fonts, themeColors, fills, borders);
        }
      }
      applyXfStyle(baseStyle, xf, fonts, themeColors, fills, borders);
    }
  }

  if (!cellNode) return { type: 'single', style: baseStyle };

  const inlineNode = cellNode.getElementsByTagName('is')[0];
  if (inlineNode) {
    const runs = Array.from(inlineNode.getElementsByTagName('r')).map((r) => {
      const rPr = r.getElementsByTagName('rPr')[0];
      let runStyle = { ...baseStyle };
      if (rPr) {
        const runFont = parseFontFromRunPr(rPr);
        runStyle = fontToCss(runFont, themeColors || {});
      }
      const textNode = r.getElementsByTagName('t')[0];
      return { text: textNode ? textNode.textContent || '' : '', style: runStyle };
    });
    return { type: 'runs', runs, baseline: baseStyle };
  }
  return { type: 'single', style: baseStyle };
}

export function applyCssToElement(element, cssObject) {
  if (!cssObject) return;
  Object.entries(cssObject).forEach(([key, value]) => { if (value != null) element.style[key] = value; });
}
