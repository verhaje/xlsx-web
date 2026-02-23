// ChartRenderer.ts - Renders Excel charts using Canvas 2D API

import type { ChartAnchor, ChartData, ChartSeries, ChartType } from '../types';

/**
 * Default color palette (matches ChartParser).
 */
const DEFAULT_COLORS = [
  '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5',
  '#70AD47', '#264478', '#9B57A0', '#636363', '#EB5757',
  '#00B0F0', '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0',
  '#9966FF', '#FF9F40', '#C9CBCF',
];

/** Layout padding/margins for the chart area */
const PADDING = { top: 40, right: 20, bottom: 60, left: 60 };
const LEGEND_HEIGHT = 24;

/**
 * Renders chart overlays positioned over the sheet table, similar to ImageRenderer.
 */
export class ChartRenderer {
  /**
   * Render chart overlays into the table container.
   */
  static renderCharts(tableContainer: HTMLElement, charts: ChartAnchor[]): void {
    if (!charts.length) return;

    const chartContainer = document.createElement('div');
    chartContainer.className = 'chart-overlay';
    chartContainer.style.position = 'absolute';
    chartContainer.style.top = '0';
    chartContainer.style.left = '0';
    chartContainer.style.pointerEvents = 'none';
    chartContainer.style.zIndex = '20';

    charts.forEach((chart) => {
      if (!chart.chartData) return;

      const wrapper = document.createElement('div');
      wrapper.className = 'chart-wrapper';
      wrapper.style.position = 'absolute';
      wrapper.style.pointerEvents = 'auto';
      wrapper.style.background = '#ffffff';
      wrapper.style.border = '1px solid #d1d5db';
      wrapper.style.borderRadius = '6px';
      wrapper.style.boxShadow = '0 2px 8px rgba(0,0,0,0.1)';
      wrapper.style.overflow = 'hidden';

      // Position using cell coordinates (simplified: 64px per col, 20px per row)
      const colWidth = 64;
      const rowHeight = 20;
      const left = (chart.from.col - 1) * colWidth;
      const top = (chart.from.row - 1) * rowHeight;
      const width = Math.max((chart.to.col - chart.from.col + 1) * colWidth, 300);
      const height = Math.max((chart.to.row - chart.from.row + 1) * rowHeight, 200);

      wrapper.style.left = `${left}px`;
      wrapper.style.top = `${top}px`;
      wrapper.style.width = `${width}px`;
      wrapper.style.height = `${height}px`;

      const canvas = document.createElement('canvas');
      const dpr = window.devicePixelRatio || 1;
      canvas.width = width * dpr;
      canvas.height = height * dpr;
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;

      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.scale(dpr, dpr);
        ChartRenderer.drawChart(ctx, chart.chartData, width, height);
      }

      wrapper.appendChild(canvas);
      chartContainer.appendChild(wrapper);
    });

    tableContainer.style.position = 'relative';
    tableContainer.appendChild(chartContainer);
  }

  /**
   * Draw a chart onto a canvas context.
   */
  static drawChart(ctx: CanvasRenderingContext2D, data: ChartData, width: number, height: number): void {
    // Background
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);

    // Title
    if (data.title) {
      ctx.fillStyle = '#1f2937';
      ctx.font = 'bold 14px Inter, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(data.title, width / 2, 24, width - 20);
    }

    if (data.series.length === 0) {
      ctx.fillStyle = '#9ca3af';
      ctx.font = '12px Inter, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText('No data', width / 2, height / 2);
      return;
    }

    switch (data.type) {
      case 'bar':
        ChartRenderer.drawBarChart(ctx, data, width, height, 'horizontal');
        break;
      case 'col':
        ChartRenderer.drawBarChart(ctx, data, width, height, 'vertical');
        break;
      case 'line':
        ChartRenderer.drawLineChart(ctx, data, width, height);
        break;
      case 'pie':
        ChartRenderer.drawPieChart(ctx, data, width, height, false);
        break;
      case 'doughnut':
        ChartRenderer.drawPieChart(ctx, data, width, height, true);
        break;
      case 'area':
        ChartRenderer.drawAreaChart(ctx, data, width, height);
        break;
      case 'scatter':
        ChartRenderer.drawScatterChart(ctx, data, width, height);
        break;
      case 'radar':
        ChartRenderer.drawRadarChart(ctx, data, width, height);
        break;
      default:
        // Fallback: try bar chart
        ChartRenderer.drawBarChart(ctx, data, width, height, 'vertical');
        break;
    }

    // Legend
    ChartRenderer.drawLegend(ctx, data.series, width, height);
  }

  // ---- Bar / Column Chart ----

  private static drawBarChart(
    ctx: CanvasRenderingContext2D, data: ChartData,
    width: number, height: number, direction: 'horizontal' | 'vertical'
  ): void {
    const series = data.series;
    const categories = ChartRenderer.getCategories(series);
    const allValues = series.flatMap(s => s.points.map(p => p.value));
    const maxVal = Math.max(...allValues, 0);
    const minVal = Math.min(...allValues, 0);
    const range = maxVal - minVal || 1;

    const plotLeft = PADDING.left;
    const plotTop = PADDING.top;
    const plotWidth = width - PADDING.left - PADDING.right;
    const plotHeight = height - PADDING.top - PADDING.bottom - LEGEND_HEIGHT;

    if (plotWidth <= 0 || plotHeight <= 0) return;

    if (direction === 'vertical') {
      // Vertical bars (column chart)
      const groupWidth = plotWidth / categories.length;
      const barWidth = Math.max(groupWidth / (series.length + 1), 2);
      const barGap = (groupWidth - barWidth * series.length) / (series.length + 1);

      // Y axis
      ChartRenderer.drawYAxis(ctx, plotLeft, plotTop, plotHeight, minVal, maxVal, 5);

      // X axis labels
      ctx.fillStyle = '#6b7280';
      ctx.font = '10px Inter, sans-serif';
      ctx.textAlign = 'center';
      categories.forEach((cat, i) => {
        const x = plotLeft + i * groupWidth + groupWidth / 2;
        const label = cat.length > 10 ? cat.substring(0, 9) + '…' : cat;
        ctx.fillText(label, x, plotTop + plotHeight + 16);
      });

      // Bars
      const zeroY = plotTop + plotHeight - ((0 - minVal) / range) * plotHeight;
      series.forEach((s, si) => {
        ctx.fillStyle = s.color || DEFAULT_COLORS[si % DEFAULT_COLORS.length];
        s.points.forEach((p, pi) => {
          if (pi >= categories.length) return;
          const x = plotLeft + pi * groupWidth + barGap + si * (barWidth + barGap);
          const valHeight = (p.value / range) * plotHeight;
          const y = p.value >= 0 ? zeroY - valHeight : zeroY;
          const h = Math.abs(valHeight);
          ChartRenderer.roundRect(ctx, x, y, barWidth, h, 2);
        });
      });

      // Gridline at zero if needed
      if (minVal < 0) {
        ctx.strokeStyle = '#d1d5db';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(plotLeft, zeroY);
        ctx.lineTo(plotLeft + plotWidth, zeroY);
        ctx.stroke();
      }
    } else {
      // Horizontal bars
      const groupHeight = plotHeight / categories.length;
      const barHeight = Math.max(groupHeight / (series.length + 1), 2);
      const barGap = (groupHeight - barHeight * series.length) / (series.length + 1);

      // X axis
      ChartRenderer.drawXAxis(ctx, plotLeft, plotTop + plotHeight, plotWidth, minVal, maxVal, 5);

      // Y axis labels
      ctx.fillStyle = '#6b7280';
      ctx.font = '10px Inter, sans-serif';
      ctx.textAlign = 'right';
      categories.forEach((cat, i) => {
        const y = plotTop + i * groupHeight + groupHeight / 2;
        const label = cat.length > 10 ? cat.substring(0, 9) + '…' : cat;
        ctx.fillText(label, plotLeft - 6, y + 3);
      });

      // Bars
      series.forEach((s, si) => {
        ctx.fillStyle = s.color || DEFAULT_COLORS[si % DEFAULT_COLORS.length];
        s.points.forEach((p, pi) => {
          if (pi >= categories.length) return;
          const y = plotTop + pi * groupHeight + barGap + si * (barHeight + barGap);
          const barW = ((p.value - minVal) / range) * plotWidth;
          ChartRenderer.roundRect(ctx, plotLeft, y, barW, barHeight, 2);
        });
      });
    }
  }

  // ---- Line Chart ----

  private static drawLineChart(
    ctx: CanvasRenderingContext2D, data: ChartData,
    width: number, height: number
  ): void {
    const series = data.series;
    const categories = ChartRenderer.getCategories(series);
    const allValues = series.flatMap(s => s.points.map(p => p.value));
    const maxVal = Math.max(...allValues, 0);
    const minVal = Math.min(...allValues, 0);
    const range = maxVal - minVal || 1;

    const plotLeft = PADDING.left;
    const plotTop = PADDING.top;
    const plotWidth = width - PADDING.left - PADDING.right;
    const plotHeight = height - PADDING.top - PADDING.bottom - LEGEND_HEIGHT;

    if (plotWidth <= 0 || plotHeight <= 0) return;

    // Axes
    ChartRenderer.drawYAxis(ctx, plotLeft, plotTop, plotHeight, minVal, maxVal, 5);

    // Grid lines
    ChartRenderer.drawGridLines(ctx, plotLeft, plotTop, plotWidth, plotHeight, 5);

    // X labels
    ctx.fillStyle = '#6b7280';
    ctx.font = '10px Inter, sans-serif';
    ctx.textAlign = 'center';
    const step = categories.length > 1 ? plotWidth / (categories.length - 1) : 0;
    categories.forEach((cat, i) => {
      const x = plotLeft + i * step;
      const label = cat.length > 8 ? cat.substring(0, 7) + '…' : cat;
      ctx.fillText(label, x, plotTop + plotHeight + 16);
    });

    // Lines
    series.forEach((s, si) => {
      const color = s.color || DEFAULT_COLORS[si % DEFAULT_COLORS.length];
      ctx.strokeStyle = color;
      ctx.lineWidth = 2;
      ctx.lineJoin = 'round';
      ctx.beginPath();

      s.points.forEach((p, pi) => {
        const x = plotLeft + (categories.length > 1 ? pi * step : plotWidth / 2);
        const y = plotTop + plotHeight - ((p.value - minVal) / range) * plotHeight;
        if (pi === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.stroke();

      // Data points
      ctx.fillStyle = color;
      s.points.forEach((p, pi) => {
        const x = plotLeft + (categories.length > 1 ? pi * step : plotWidth / 2);
        const y = plotTop + plotHeight - ((p.value - minVal) / range) * plotHeight;
        ctx.beginPath();
        ctx.arc(x, y, 3, 0, Math.PI * 2);
        ctx.fill();
      });
    });
  }

  // ---- Area Chart ----

  private static drawAreaChart(
    ctx: CanvasRenderingContext2D, data: ChartData,
    width: number, height: number
  ): void {
    const series = data.series;
    const categories = ChartRenderer.getCategories(series);
    const allValues = series.flatMap(s => s.points.map(p => p.value));
    const maxVal = Math.max(...allValues, 0);
    const minVal = Math.min(...allValues, 0);
    const range = maxVal - minVal || 1;

    const plotLeft = PADDING.left;
    const plotTop = PADDING.top;
    const plotWidth = width - PADDING.left - PADDING.right;
    const plotHeight = height - PADDING.top - PADDING.bottom - LEGEND_HEIGHT;

    if (plotWidth <= 0 || plotHeight <= 0) return;

    ChartRenderer.drawYAxis(ctx, plotLeft, plotTop, plotHeight, minVal, maxVal, 5);
    ChartRenderer.drawGridLines(ctx, plotLeft, plotTop, plotWidth, plotHeight, 5);

    const step = categories.length > 1 ? plotWidth / (categories.length - 1) : 0;
    const baselineY = plotTop + plotHeight;

    // X labels
    ctx.fillStyle = '#6b7280';
    ctx.font = '10px Inter, sans-serif';
    ctx.textAlign = 'center';
    categories.forEach((cat, i) => {
      const x = plotLeft + i * step;
      const label = cat.length > 8 ? cat.substring(0, 7) + '…' : cat;
      ctx.fillText(label, x, baselineY + 16);
    });

    // Draw areas back to front
    for (let si = series.length - 1; si >= 0; si -= 1) {
      const s = series[si];
      const color = s.color || DEFAULT_COLORS[si % DEFAULT_COLORS.length];

      ctx.fillStyle = ChartRenderer.hexToRgba(color, 0.3);
      ctx.strokeStyle = color;
      ctx.lineWidth = 2;

      ctx.beginPath();
      ctx.moveTo(plotLeft, baselineY);

      s.points.forEach((p, pi) => {
        const x = plotLeft + (categories.length > 1 ? pi * step : plotWidth / 2);
        const y = plotTop + plotHeight - ((p.value - minVal) / range) * plotHeight;
        ctx.lineTo(x, y);
      });

      const lastX = plotLeft + (categories.length > 1 ? (s.points.length - 1) * step : plotWidth / 2);
      ctx.lineTo(lastX, baselineY);
      ctx.closePath();
      ctx.fill();

      // Stroke the line on top
      ctx.beginPath();
      s.points.forEach((p, pi) => {
        const x = plotLeft + (categories.length > 1 ? pi * step : plotWidth / 2);
        const y = plotTop + plotHeight - ((p.value - minVal) / range) * plotHeight;
        if (pi === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.stroke();
    }
  }

  // ---- Pie / Doughnut Chart ----

  private static drawPieChart(
    ctx: CanvasRenderingContext2D, data: ChartData,
    width: number, height: number, isDoughnut: boolean
  ): void {
    const series = data.series[0]; // Pie charts typically have one series
    if (!series || series.points.length === 0) return;

    const centerX = width / 2;
    const centerY = (PADDING.top + height - LEGEND_HEIGHT) / 2;
    const radius = Math.min(
      (width - PADDING.left - PADDING.right) / 2,
      (height - PADDING.top - PADDING.bottom - LEGEND_HEIGHT) / 2
    ) * 0.85;

    if (radius <= 0) return;

    const total = series.points.reduce((acc, p) => acc + Math.abs(p.value), 0);
    if (total === 0) return;

    let startAngle = -Math.PI / 2;

    series.points.forEach((p, i) => {
      const sliceAngle = (Math.abs(p.value) / total) * Math.PI * 2;
      const color = DEFAULT_COLORS[i % DEFAULT_COLORS.length];

      ctx.fillStyle = color;
      ctx.beginPath();
      ctx.moveTo(centerX, centerY);
      ctx.arc(centerX, centerY, radius, startAngle, startAngle + sliceAngle);
      ctx.closePath();
      ctx.fill();

      // Slice border
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 2;
      ctx.stroke();

      // Label
      if (sliceAngle > 0.2) {
        const midAngle = startAngle + sliceAngle / 2;
        const labelR = radius * 0.65;
        const lx = centerX + Math.cos(midAngle) * labelR;
        const ly = centerY + Math.sin(midAngle) * labelR;
        const pct = ((Math.abs(p.value) / total) * 100).toFixed(0);

        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 11px Inter, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(`${pct}%`, lx, ly);
      }

      startAngle += sliceAngle;
    });

    // Doughnut hole
    if (isDoughnut) {
      ctx.fillStyle = '#ffffff';
      ctx.beginPath();
      ctx.arc(centerX, centerY, radius * 0.5, 0, Math.PI * 2);
      ctx.fill();
    }
  }

  // ---- Scatter Chart ----

  private static drawScatterChart(
    ctx: CanvasRenderingContext2D, data: ChartData,
    width: number, height: number
  ): void {
    const series = data.series;
    const allX: number[] = [];
    const allY: number[] = [];

    series.forEach(s => {
      s.points.forEach(p => {
        allX.push(parseFloat(p.category) || 0);
        allY.push(p.value);
      });
    });

    if (allX.length === 0) return;

    const minX = Math.min(...allX);
    const maxX = Math.max(...allX);
    const minY = Math.min(...allY, 0);
    const maxY = Math.max(...allY, 0);
    const rangeX = maxX - minX || 1;
    const rangeY = maxY - minY || 1;

    const plotLeft = PADDING.left;
    const plotTop = PADDING.top;
    const plotWidth = width - PADDING.left - PADDING.right;
    const plotHeight = height - PADDING.top - PADDING.bottom - LEGEND_HEIGHT;

    if (plotWidth <= 0 || plotHeight <= 0) return;

    ChartRenderer.drawYAxis(ctx, plotLeft, plotTop, plotHeight, minY, maxY, 5);
    ChartRenderer.drawXAxis(ctx, plotLeft, plotTop + plotHeight, plotWidth, minX, maxX, 5);
    ChartRenderer.drawGridLines(ctx, plotLeft, plotTop, plotWidth, plotHeight, 5);

    series.forEach((s, si) => {
      const color = s.color || DEFAULT_COLORS[si % DEFAULT_COLORS.length];
      ctx.fillStyle = color;

      s.points.forEach(p => {
        const xVal = parseFloat(p.category) || 0;
        const px = plotLeft + ((xVal - minX) / rangeX) * plotWidth;
        const py = plotTop + plotHeight - ((p.value - minY) / rangeY) * plotHeight;

        ctx.beginPath();
        ctx.arc(px, py, 4, 0, Math.PI * 2);
        ctx.fill();
      });
    });
  }

  // ---- Radar Chart ----

  private static drawRadarChart(
    ctx: CanvasRenderingContext2D, data: ChartData,
    width: number, height: number
  ): void {
    const series = data.series;
    const categories = ChartRenderer.getCategories(series);
    if (categories.length < 3) return;

    const centerX = width / 2;
    const centerY = (PADDING.top + height - LEGEND_HEIGHT) / 2;
    const radius = Math.min(
      (width - PADDING.left - PADDING.right) / 2,
      (height - PADDING.top - PADDING.bottom - LEGEND_HEIGHT) / 2
    ) * 0.75;

    if (radius <= 0) return;

    const allValues = series.flatMap(s => s.points.map(p => p.value));
    const maxVal = Math.max(...allValues, 1);
    const n = categories.length;
    const angleStep = (Math.PI * 2) / n;

    // Draw grid circles
    ctx.strokeStyle = '#e5e7eb';
    ctx.lineWidth = 1;
    for (let level = 1; level <= 4; level += 1) {
      const r = (level / 4) * radius;
      ctx.beginPath();
      for (let i = 0; i <= n; i += 1) {
        const angle = -Math.PI / 2 + i * angleStep;
        const x = centerX + Math.cos(angle) * r;
        const y = centerY + Math.sin(angle) * r;
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }
      ctx.stroke();
    }

    // Axis lines and labels
    ctx.strokeStyle = '#d1d5db';
    ctx.fillStyle = '#6b7280';
    ctx.font = '10px Inter, sans-serif';
    ctx.textAlign = 'center';
    for (let i = 0; i < n; i += 1) {
      const angle = -Math.PI / 2 + i * angleStep;
      const x = centerX + Math.cos(angle) * radius;
      const y = centerY + Math.sin(angle) * radius;

      ctx.beginPath();
      ctx.moveTo(centerX, centerY);
      ctx.lineTo(x, y);
      ctx.stroke();

      const labelR = radius + 14;
      const lx = centerX + Math.cos(angle) * labelR;
      const ly = centerY + Math.sin(angle) * labelR;
      ctx.fillText(categories[i].substring(0, 8), lx, ly + 3);
    }

    // Series polygons
    series.forEach((s, si) => {
      const color = s.color || DEFAULT_COLORS[si % DEFAULT_COLORS.length];

      ctx.fillStyle = ChartRenderer.hexToRgba(color, 0.2);
      ctx.strokeStyle = color;
      ctx.lineWidth = 2;

      ctx.beginPath();
      s.points.forEach((p, pi) => {
        if (pi >= n) return;
        const angle = -Math.PI / 2 + pi * angleStep;
        const r = (p.value / maxVal) * radius;
        const x = centerX + Math.cos(angle) * r;
        const y = centerY + Math.sin(angle) * r;
        if (pi === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.closePath();
      ctx.fill();
      ctx.stroke();

      // Data points
      ctx.fillStyle = color;
      s.points.forEach((p, pi) => {
        if (pi >= n) return;
        const angle = -Math.PI / 2 + pi * angleStep;
        const r = (p.value / maxVal) * radius;
        const x = centerX + Math.cos(angle) * r;
        const y = centerY + Math.sin(angle) * r;
        ctx.beginPath();
        ctx.arc(x, y, 3, 0, Math.PI * 2);
        ctx.fill();
      });
    });
  }

  // ---- Legend ----

  private static drawLegend(
    ctx: CanvasRenderingContext2D, series: ChartSeries[],
    width: number, height: number
  ): void {
    if (series.length <= 1 && series[0]?.points.length <= 1) return;

    const isPie = series.length === 1 && series[0].points.length > 1;
    const items = isPie
      ? series[0].points.map((p, i) => ({
          name: p.category,
          color: DEFAULT_COLORS[i % DEFAULT_COLORS.length],
        }))
      : series.map((s, i) => ({
          name: s.name,
          color: s.color || DEFAULT_COLORS[i % DEFAULT_COLORS.length],
        }));

    const legendY = height - LEGEND_HEIGHT + 4;
    const entryPadding = 12;
    const swatchSize = 10;

    ctx.font = '10px Inter, sans-serif';

    // Measure total width
    let totalWidth = 0;
    items.forEach(item => {
      totalWidth += swatchSize + 4 + ctx.measureText(item.name).width + entryPadding;
    });
    totalWidth -= entryPadding;

    let x = Math.max(PADDING.left, (width - totalWidth) / 2);

    items.forEach(item => {
      // Swatch
      ctx.fillStyle = item.color;
      ctx.fillRect(x, legendY, swatchSize, swatchSize);

      // Label
      ctx.fillStyle = '#6b7280';
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      ctx.fillText(item.name, x + swatchSize + 4, legendY + swatchSize / 2, width - x - 10);

      x += swatchSize + 4 + ctx.measureText(item.name).width + entryPadding;
    });
  }

  // ---- Axis / Grid helpers ----

  private static drawYAxis(
    ctx: CanvasRenderingContext2D,
    x: number, y: number, height: number,
    minVal: number, maxVal: number, ticks: number
  ): void {
    ctx.strokeStyle = '#d1d5db';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x, y + height);
    ctx.stroke();

    ctx.fillStyle = '#6b7280';
    ctx.font = '10px Inter, sans-serif';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';

    const range = maxVal - minVal || 1;
    for (let i = 0; i <= ticks; i += 1) {
      const val = minVal + (range * i) / ticks;
      const ty = y + height - (height * i) / ticks;
      ctx.fillText(ChartRenderer.formatAxisValue(val), x - 6, ty);
    }
  }

  private static drawXAxis(
    ctx: CanvasRenderingContext2D,
    x: number, y: number, width: number,
    minVal: number, maxVal: number, ticks: number
  ): void {
    ctx.strokeStyle = '#d1d5db';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + width, y);
    ctx.stroke();

    ctx.fillStyle = '#6b7280';
    ctx.font = '10px Inter, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';

    const range = maxVal - minVal || 1;
    for (let i = 0; i <= ticks; i += 1) {
      const val = minVal + (range * i) / ticks;
      const tx = x + (width * i) / ticks;
      ctx.fillText(ChartRenderer.formatAxisValue(val), tx, y + 4);
    }
  }

  private static drawGridLines(
    ctx: CanvasRenderingContext2D,
    x: number, y: number, width: number, height: number, ticks: number
  ): void {
    ctx.strokeStyle = '#f3f4f6';
    ctx.lineWidth = 1;
    for (let i = 1; i <= ticks; i += 1) {
      const ty = y + height - (height * i) / ticks;
      ctx.beginPath();
      ctx.moveTo(x, ty);
      ctx.lineTo(x + width, ty);
      ctx.stroke();
    }
  }

  // ---- Utility helpers ----

  private static getCategories(series: ChartSeries[]): string[] {
    if (series.length === 0) return [];
    // Use the series with the most points
    let maxLen = 0;
    let best = series[0];
    for (const s of series) {
      if (s.points.length > maxLen) {
        maxLen = s.points.length;
        best = s;
      }
    }
    return best.points.map(p => p.category);
  }

  private static formatAxisValue(val: number): string {
    if (Math.abs(val) >= 1000000) return (val / 1000000).toFixed(1) + 'M';
    if (Math.abs(val) >= 1000) return (val / 1000).toFixed(1) + 'K';
    if (Number.isInteger(val)) return val.toString();
    return val.toFixed(1);
  }

  private static hexToRgba(hex: string, alpha: number): string {
    const h = hex.replace('#', '');
    const r = parseInt(h.substring(0, 2), 16);
    const g = parseInt(h.substring(2, 4), 16);
    const b = parseInt(h.substring(4, 6), 16);
    if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return `rgba(100,100,100,${alpha})`;
    return `rgba(${r},${g},${b},${alpha})`;
  }

  private static roundRect(
    ctx: CanvasRenderingContext2D,
    x: number, y: number, w: number, h: number, r: number
  ): void {
    if (h < 1) return;
    r = Math.min(r, Math.abs(h) / 2, Math.abs(w) / 2);
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + w - r, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + r);
    ctx.lineTo(x + w, y + h - r);
    ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    ctx.lineTo(x + r, y + h);
    ctx.quadraticCurveTo(x, y + h, x, y + h - r);
    ctx.lineTo(x, y + r);
    ctx.quadraticCurveTo(x, y, x + r, y);
    ctx.closePath();
    ctx.fill();
  }
}
