/**
 * The chart catalogue of the Insert Chart dialog and the ribbon's quick chart menus (categories
 * and subtypes), and inserting a chart for the selected data.
 */
import type { ChartSeries, ChartType, SeriesKind, SheetChart } from '../core/model/drawing';
import { chartFromRange, guessLayout, readChartData, resolveRef, seriesForLayout, trimRange } from '../core/chartData';
import type { SpreadsheetDocument } from '../core/document';
import { filterRegion } from '../core/filter';
import { SheetLayout } from '../core/layout';
import type { Sheet } from '../core/model/sheet';
import { makeRange, type Range } from '../core/refs';
import { chartSvg } from '../grid/chartRender';
import type { IconName } from './icons';
import type { UIContextValue } from './context';
import { commit } from './commands';

export interface ChartSubtype {
  id: string;
  label: string;
  type: ChartType;
  icon: IconName;
  /** Settings of the subtype (markers, scatter lines, stock layout). */
  patch?: Partial<SheetChart>;
  /** Combo charts: how the series are drawn. */
  combo?: 'column-line' | 'column-line-secondary' | 'area-column' | 'custom';
}

export interface ChartCategory {
  id: string;
  label: string;
  icon: IconName;
  subtypes: ChartSubtype[];
}

export const CHART_CATEGORIES: ChartCategory[] = [
  {
    id: 'column',
    label: 'Column',
    icon: 'bar_chart',
    subtypes: [
      { id: 'column-clustered', label: 'Clustered Column', type: 'column', icon: 'bar_chart' },
      { id: 'column-stacked', label: 'Stacked Column', type: 'stacked-column', icon: 'stacked_bar_chart' },
      { id: 'column-percent', label: '100% Stacked Column', type: 'percent-column', icon: 'leaderboard' },
    ],
  },
  {
    id: 'line',
    label: 'Line',
    icon: 'show_chart',
    subtypes: [
      { id: 'line', label: 'Line', type: 'line', icon: 'show_chart' },
      { id: 'line-stacked', label: 'Stacked Line', type: 'stacked-line', icon: 'stacked_line_chart' },
      { id: 'line-percent', label: '100% Stacked Line', type: 'percent-line', icon: 'ssid_chart' },
      { id: 'line-markers', label: 'Line with Markers', type: 'line', icon: 'timeline', patch: { markers: true } },
      { id: 'line-stacked-markers', label: 'Stacked Line with Markers', type: 'stacked-line', icon: 'multiline_chart', patch: { markers: true } },
      { id: 'line-percent-markers', label: '100% Stacked Line with Markers', type: 'percent-line', icon: 'insights', patch: { markers: true } },
    ],
  },
  {
    id: 'pie',
    label: 'Pie',
    icon: 'pie_chart',
    subtypes: [
      { id: 'pie', label: 'Pie', type: 'pie', icon: 'pie_chart' },
      { id: 'doughnut', label: 'Doughnut', type: 'doughnut', icon: 'donut_large' },
    ],
  },
  {
    id: 'bar',
    label: 'Bar',
    icon: 'align_horizontal_left',
    subtypes: [
      { id: 'bar-clustered', label: 'Clustered Bar', type: 'bar', icon: 'align_horizontal_left' },
      { id: 'bar-stacked', label: 'Stacked Bar', type: 'stacked-bar', icon: 'horizontal_split' },
      { id: 'bar-percent', label: '100% Stacked Bar', type: 'percent-bar', icon: 'view_headline' },
    ],
  },
  {
    id: 'area',
    label: 'Area',
    icon: 'area_chart',
    subtypes: [
      { id: 'area', label: 'Area', type: 'area', icon: 'area_chart' },
      { id: 'area-stacked', label: 'Stacked Area', type: 'stacked-area', icon: 'stacked_line_chart' },
      { id: 'area-percent', label: '100% Stacked Area', type: 'percent-area', icon: 'ssid_chart' },
    ],
  },
  {
    id: 'scatter',
    label: 'X Y (Scatter)',
    icon: 'scatter_plot',
    subtypes: [
      { id: 'scatter', label: 'Scatter', type: 'scatter', icon: 'scatter_plot', patch: { scatterStyle: 'markers' } },
      { id: 'scatter-smooth-markers', label: 'Scatter with Smooth Lines and Markers', type: 'scatter', icon: 'insights', patch: { scatterStyle: 'smooth-markers' } },
      { id: 'scatter-smooth', label: 'Scatter with Smooth Lines', type: 'scatter', icon: 'ssid_chart', patch: { scatterStyle: 'smooth' } },
      { id: 'scatter-lines-markers', label: 'Scatter with Straight Lines and Markers', type: 'scatter', icon: 'timeline', patch: { scatterStyle: 'lines-markers' } },
      { id: 'scatter-lines', label: 'Scatter with Straight Lines', type: 'scatter', icon: 'show_chart', patch: { scatterStyle: 'lines' } },
      { id: 'bubble', label: 'Bubble', type: 'bubble', icon: 'bubble_chart' },
    ],
  },
  {
    id: 'stock',
    label: 'Stock',
    icon: 'candlestick_chart',
    subtypes: [
      { id: 'stock-hlc', label: 'High-Low-Close', type: 'stock', icon: 'candlestick_chart', patch: { stockKind: 'hlc' } },
      { id: 'stock-ohlc', label: 'Open-High-Low-Close', type: 'stock', icon: 'candlestick_chart', patch: { stockKind: 'ohlc' } },
      { id: 'stock-vhlc', label: 'Volume-High-Low-Close', type: 'stock', icon: 'candlestick_chart', patch: { stockKind: 'vhlc' } },
      { id: 'stock-vohlc', label: 'Volume-Open-High-Low-Close', type: 'stock', icon: 'candlestick_chart', patch: { stockKind: 'vohlc' } },
    ],
  },
  {
    id: 'radar',
    label: 'Radar',
    icon: 'track_changes',
    subtypes: [
      { id: 'radar', label: 'Radar', type: 'radar', icon: 'track_changes' },
      { id: 'radar-markers', label: 'Radar with Markers', type: 'radar', icon: 'radar', patch: { markers: true } },
      { id: 'radar-filled', label: 'Filled Radar', type: 'filled-radar', icon: 'change_history' },
    ],
  },
  { id: 'sunburst', label: 'Sunburst', icon: 'donut_small', subtypes: [{ id: 'sunburst', label: 'Sunburst', type: 'sunburst', icon: 'donut_small' }] },
  { id: 'treemap', label: 'Treemap', icon: 'view_quilt', subtypes: [{ id: 'treemap', label: 'Treemap', type: 'treemap', icon: 'view_quilt' }] },
  { id: 'funnel', label: 'Funnel', icon: 'filter_alt', subtypes: [{ id: 'funnel', label: 'Funnel', type: 'funnel', icon: 'filter_alt' }] },
  { id: 'waterfall', label: 'Waterfall', icon: 'waterfall_chart', subtypes: [{ id: 'waterfall', label: 'Waterfall', type: 'waterfall', icon: 'waterfall_chart' }] },
  {
    id: 'statistic',
    label: 'Statistic',
    icon: 'equalizer',
    subtypes: [
      { id: 'histogram', label: 'Histogram', type: 'histogram', icon: 'equalizer' },
      { id: 'pareto', label: 'Pareto', type: 'pareto', icon: 'auto_graph' },
      { id: 'box-whisker', label: 'Box and Whisker', type: 'box-whisker', icon: 'vertical_align_center' },
    ],
  },
  {
    id: 'combo',
    label: 'Combo',
    icon: 'insert_chart',
    subtypes: [
      { id: 'combo-column-line', label: 'Clustered Column - Line', type: 'combo', icon: 'insert_chart', combo: 'column-line' },
      { id: 'combo-column-line-secondary', label: 'Clustered Column - Line on Secondary Axis', type: 'combo', icon: 'addchart', combo: 'column-line-secondary' },
      { id: 'combo-area-column', label: 'Stacked Area - Clustered Column', type: 'combo', icon: 'area_chart', combo: 'area-column' },
      { id: 'combo-custom', label: 'Custom Combination', type: 'combo', icon: 'tune', combo: 'custom' },
    ],
  },
];

export const ALL_SUBTYPES: ChartSubtype[] = CHART_CATEGORIES.flatMap((c) => c.subtypes);

export function subtypeById(id: string): ChartSubtype | undefined {
  return ALL_SUBTYPES.find((s) => s.id === id);
}

/** The catalogue entry matching a chart's settings. */
export function subtypeOf(chart: Pick<SheetChart, 'type' | 'markers' | 'scatterStyle' | 'stockKind' | 'series'>): ChartSubtype {
  const candidates = ALL_SUBTYPES.filter((s) => s.type === chart.type);
  if (chart.type === 'combo') {
    const n = chart.series.length;
    const kinds = chart.series.map((s, i) => s.kind ?? (i === n - 1 && n > 1 ? 'line' : 'column'));
    for (const c of candidates) {
      if (c.combo === 'custom' || !c.combo) continue;
      const want = comboKinds(c.combo, n);
      if (want.every((k, i) => k.kind === kinds[i] && !!k.secondary === !!chart.series[i].secondaryAxis)) return c;
    }
    return candidates.find((c) => c.combo === 'custom')!;
  }
  return (
    candidates.find((s) => {
      const p = s.patch ?? {};
      if (chart.type === 'scatter') return (chart.scatterStyle ?? 'markers') === p.scatterStyle;
      if (chart.type === 'stock') return (chart.stockKind ?? 'hlc') === p.stockKind;
      return !!p.markers === !!chart.markers;
    }) ?? candidates[0] ?? ALL_SUBTYPES[0]
  );
}

/** Series kinds of the combo presets. */
export function comboKinds(preset: NonNullable<ChartSubtype['combo']>, n: number): { kind: SeriesKind; secondary?: boolean }[] {
  return Array.from({ length: n }, (_, i) => {
    const last = i === n - 1 && n > 1;
    switch (preset) {
      case 'column-line-secondary':
        return { kind: last ? 'line' : 'column', secondary: last };
      case 'area-column':
        return { kind: i < Math.ceil(n / 2) ? 'stacked-area' : 'column' };
      default:
        return { kind: last ? 'line' : 'column' };
    }
  });
}

/** Chart settings for a subtype (the series keep their references). */
export function applySubtype(chart: Omit<SheetChart, 'id' | 'z' | 'from' | 'to'>, sub: ChartSubtype, custom?: { kind: SeriesKind; secondary?: boolean }[]): Omit<SheetChart, 'id' | 'z' | 'from' | 'to'> {
  const reset: Partial<SheetChart> = { markers: undefined, scatterStyle: undefined, stockKind: undefined };
  let series: ChartSeries[] = chart.series.map((s) => ({ ...s, kind: undefined, markers: undefined, smooth: undefined, secondaryAxis: sub.type === 'combo' ? undefined : s.secondaryAxis }));
  if (sub.type === 'combo') {
    const kinds = custom ?? comboKinds(sub.combo === 'custom' || !sub.combo ? 'column-line' : sub.combo, series.length);
    series = series.map((s, i) => ({ ...s, kind: kinds[i]?.kind ?? 'column', secondaryAxis: kinds[i]?.secondary || undefined }));
  }
  return { ...chart, ...reset, ...sub.patch, type: sub.type, series };
}

/** The data a new chart is built from: the selection, or the data region around the active cell. */
export function chartSource(ctx: UIContextValue): ReturnType<typeof makeRange> | null {
  const { doc } = ctx;
  const sel = doc.selection;
  const range = doc.range;
  if (range.r0 !== range.r1 || range.c0 !== range.c1) return trimRange(doc.sheet, range);
  const table = doc.tableAt(sel.active.r, sel.active.c);
  if (table) return table.ref;
  return filterRegion(doc.sheet, makeRange(sel.active.r, sel.active.c), sel.active);
}

/** Inserts a chart of the subtype for the selected data; returns its id. */
export function insertChart(ctx: UIContextValue, sub: ChartSubtype, custom?: { kind: SeriesKind; secondary?: boolean }[]): string | null {
  commit(ctx);
  const { doc } = ctx;
  const blocked = doc.blockedBy('objects');
  if (blocked) {
    ctx.notify(blocked);
    return null;
  }
  const source = chartSource(ctx);
  if (!source) {
    ctx.notify('Select the data for the chart first.');
    return null;
  }
  const sheet = doc.sheet;
  let def = applySubtype(chartFromRange(doc, sheet, source, sub.type), sub, custom);
  if (!def.series.length) {
    ctx.notify('The selection has no data to chart.');
    return null;
  }
  const single = def.series.length === 1 && def.series[0].nameRef;
  def = { ...def, title: single ? doc.displayText(...cellOf(def.series[0].nameRef!), sheet) || 'Chart Title' : 'Chart Title' };
  const n = sheet.charts.length + 1;
  def.name = `Chart ${n}`;
  // Desktop: to the right of the data; phones: under it, as wide as the screen allows.
  const layout = SheetLayout.of(sheet);
  let rect: { x: number; y: number; width: number; height: number };
  if (ctx.phone) {
    const width = Math.min(480, Math.max(240, window.innerWidth - 60));
    rect = { x: layout.colX(source.c0), y: layout.rowY(Math.min(source.r1 + 2, 1048575)), width, height: Math.round(width * 0.625) };
  } else {
    rect = { x: layout.colX(Math.min(source.c1 + 1, 16383)) + 20, y: layout.rowY(source.r0) + 20, width: 480, height: 300 };
  }
  const chart = doc.addChart(def, rect);
  if (!chart) return null;
  ctx.selectObject(chart.id);
  ctx.showTab('chart-design');
  if (ctx.phone) ctx.grid?.scrollIntoView(Math.min(source.r1 + 2, 1048575), source.c0);
  return chart.id;
}

function cellOf(ref: string): [number, number] {
  const m = /\$?([A-Z]+)\$?(\d+)$/i.exec(ref);
  if (!m) return [0, 0];
  let c = 0;
  for (const ch of m[1].toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64;
  return [Number(m[2]) - 1, c - 1];
}

// -----------------------------------------------------------------------------------------------
// The chart dialogs (Insert Chart, Change Chart Type, Select Data)

/** Series kinds offered for each series of a custom combo chart. */
export const COMBO_SERIES_KINDS: { kind: SeriesKind; label: string }[] = [
  { kind: 'column', label: 'Clustered Column' },
  { kind: 'stacked-column', label: 'Stacked Column' },
  { kind: 'line', label: 'Line' },
  { kind: 'area', label: 'Area' },
  { kind: 'stacked-area', label: 'Stacked Area' },
];

/** Chart types whose data has one meaningful orientation (no second preview). */
export const ONE_WAY_CHARTS = new Set<SheetChart['type']>(['sunburst', 'treemap', 'histogram', 'box-whisker', 'pareto', 'waterfall', 'funnel', 'stock', 'scatter', 'bubble']);

/** The picture of a chart definition, for the dialogs' previews. */
export function chartPreviewSvg(doc: SpreadsheetDocument, chart: Omit<SheetChart, 'id' | 'z' | 'from' | 'to'>, w: number, h: number): string {
  const full: SheetChart = { ...chart, id: 'preview', z: 0, from: { r: 0, c: 0, dx: 0, dy: 0 }, to: { r: 0, c: 0, dx: 0, dy: 0 } };
  return chartSvg(full, readChartData(doc, full), w, h);
}

/** A chart for a range with the series forced into rows or columns. */
export function forcedChart(doc: SpreadsheetDocument, sheet: Sheet, range: Range, type: SheetChart['type'], seriesIn: 'rows' | 'columns'): Omit<SheetChart, 'id' | 'z' | 'from' | 'to'> {
  const base = chartFromRange(doc, sheet, range, type);
  const area = trimRange(sheet, range);
  const layout = guessLayout(doc, sheet, area, type, seriesIn);
  const { categories, series } = seriesForLayout(sheet, area, layout, type);
  return { ...base, seriesIn, headers: layout.headers, categories, series };
}

/** The name of a chart series as the dialogs list it. */
export function chartSeriesName(doc: SpreadsheetDocument, s: ChartSeries, i: number): string {
  if (s.nameRef) {
    const hit = resolveRef(doc, s.nameRef);
    if (hit) return doc.displayText(hit.range.r0, hit.range.c0, hit.sheet) || `Series ${i + 1}`;
  }
  return s.name ?? `Series ${i + 1}`;
}

/** Series of a chart's range read the other way (Switch Row/Column). */
export function switchedSeries(doc: SpreadsheetDocument, chart: SheetChart, seriesIn: 'rows' | 'columns'): Partial<SheetChart> {
  const hit = resolveRef(doc, chart.range);
  if (!hit) return {};
  const def = forcedChart(doc, hit.sheet, hit.range, chart.type, seriesIn);
  return { seriesIn, headers: def.headers, categories: def.categories, series: def.series };
}
