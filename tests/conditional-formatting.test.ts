import { describe, it, expect } from 'vitest';
import { StyleApplicator } from '../src/ts/styles/StyleApplicator';
import type { DxfStyle } from '../src/ts/types';

describe('Conditional Formatting Styles', () => {
  it('DXF to CSS with border', () => {
    const dxf: DxfStyle = {
      font: { bold: true, rgb: 'FFFF0000' },
      fill: { rgb: 'FF00FF00' },
      border: { top: { style: 'thin', rgb: 'FF0000FF' } },
    };
    const css = StyleApplicator.dxfToCss(dxf, {});
    expect(css).toEqual({
      fontWeight: 'bold',
      color: '#FF0000',
      backgroundColor: '#00FF00',
      borderTop: '1px solid #0000FF',
    });
  });

  it('DXF to CSS with theme colors', () => {
    const themeColors: Record<number, string> = { 2: '#123456', 5: '#abcdef' };
    const dxf: DxfStyle = {
      font: { theme: '2' },
      fill: { theme: '5' },
      border: { bottom: { style: 'medium', theme: '2' } },
    };
    const css = StyleApplicator.dxfToCss(dxf, themeColors);
    expect(css).toEqual({
      color: '#123456',
      backgroundColor: '#abcdef',
      borderBottom: '2px solid #123456',
    });
  });
});
