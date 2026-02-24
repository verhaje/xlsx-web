// ColorResolver.ts - Color resolution and manipulation utilities

import type { ColorObject } from '../types';

/** Standard Excel indexed color palette (BIFF8) */
const INDEXED_COLORS: string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', // 0-7
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', // 8-15
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', // 16-23
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', // 24-31
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', // 32-39
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', // 40-47
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', // 48-55
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333', // 56-63
];

/**
 * Resolves and manipulates colors from Excel color objects.
 */
export class ColorResolver {
  /**
   * Normalize an RGB hex string to `#RRGGBB` format.
   */
  static normalizeRgb(raw: string | null | undefined): string | undefined {
    if (!raw) return undefined;
    const s = raw.replace(/^0x/i, '').replace(/^#/, '');
    if (s.length === 8) return `#${s.slice(2)}`;
    if (s.length === 6) return `#${s}`;
    return undefined;
  }

  /**
   * Apply a tint value to a hex color string.
   */
  static applyTint(hex: string, tint: string | null | undefined): string {
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

    const clamp = (v: number): number => Math.max(0, Math.min(255, v));
    return `#${clamp(r).toString(16).padStart(2, '0')}${clamp(g).toString(16).padStart(2, '0')}${clamp(b).toString(16).padStart(2, '0')}`;
  }

  /**
   * Resolve a ColorObject to a CSS hex color string.
   */
  static resolveColor(
    colorObj: ColorObject | null | undefined,
    themeColors: Record<number, string> | null | undefined
  ): string | undefined {
    if (!colorObj) return undefined;

    if (colorObj.rgb) {
      const base = ColorResolver.normalizeRgb(colorObj.rgb);
      if (!base) return undefined;
      return colorObj.tint ? ColorResolver.applyTint(base.replace('#', ''), colorObj.tint) : base;
    }

    if (colorObj.theme != null) {
      const idx = parseInt(colorObj.theme, 10);
      if (!Number.isNaN(idx) && themeColors && themeColors[idx]) {
        const base = themeColors[idx].replace('#', '');
        return colorObj.tint ? ColorResolver.applyTint(base, colorObj.tint) : `#${base}`;
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
}
