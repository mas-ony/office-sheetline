/**
 * Charts and shapes in DrawingML: reading chart parts (classic charts and Office 2016 "chartex"),
 * shapes and connectors from drawing anchors into the editor's model, and writing them back.
 *
 * Charts read from a file keep their original part while their settings are unchanged (so
 * formatting the editor does not model survives); edited and new charts are written from the
 * model, with cached values so viewers that do not recalculate still show them.
 */
import { CHART_PALETTE, type ChartSeries, type ChartType, type ErrorBars, type LineDash, type ScatterStyle, type SeriesKind, type SheetChart, type SheetShape, type Side, type Trendline } from '../model/drawing';
import type { AnchorPoint, Sheet } from '../model/sheet';
import type { Theme, ThemeSlot } from '../model/color';
import type { Workbook } from '../model/workbook';
import { SheetLayout } from '../layout';
import { parseRange, splitSheetRef } from '../refs';
import { formatValue } from '../format/display';
import { attr, el, elements, escapeAttr, escapeText, numAttr, textContent, txt, type XmlElement, type XmlNode } from '../xml/xml';
import { EMU_PER_PX } from './util';

export const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
export const REL_CHARTEX = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
export const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const CT_CHARTEX = 'application/vnd.ms-office.chartex+xml';

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const NS_XE = 'urn:schemas-xlsx-editor:extensions';
/** Extension of chart parts holding editor settings that the chart schema has no place for. */
const SETTINGS_URI = '{5D0E8E86-7C2A-4C5B-9A61-2B8F3C1E7D40}';

/** Series colours of charts made in the editor (written explicitly, so Excel shows the same). */
const EDITOR_COLORS = CHART_PALETTE.map((c) => c.slice(1));

const local = (name: string) => name.slice(name.indexOf(':') + 1);

function kid(e: XmlElement | undefined, name: string): XmlElement | undefined {
  if (!e) return undefined;
  for (const c of e.children) if (c.type === 'element' && local(c.name) === name) return c;
  return undefined;
}

function kids(e: XmlElement | undefined, name: string): XmlElement[] {
  return elements(e).filter((c) => local(c.name) === name);
}

function deep(e: XmlElement | undefined, name: string, out: XmlElement[] = []): XmlElement[] {
  if (!e) return out;
  for (const c of e.children) {
    if (c.type !== 'element') continue;
    if (local(c.name) === name) out.push(c);
    deep(c, name, out);
  }
  return out;
}

// -------------------------------------------------------------------------------------------------
// DrawingML colours

const SCHEME_ALIASES: Record<string, ThemeSlot> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
const PRESET_COLORS: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080',
  darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500', purple: '800080', navy: '000080', teal: '008080', maroon: '800000', olive: '808000',
  silver: 'C0C0C0', lime: '00FF00', aqua: '00FFFF', cyan: '00FFFF', fuchsia: 'FF00FF', magenta: 'FF00FF', brown: 'A52A2A',
};

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (!s) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}

const hex2 = (n: number) => Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, '0').toUpperCase();

/** '#RRGGBB' of a DrawingML colour element (a:srgbClr, a:schemeClr...) with its modifiers applied. */
export function colorOf(c: XmlElement | undefined, theme: Theme): string | undefined {
  if (!c) return undefined;
  let hex: string | undefined;
  switch (local(c.name)) {
    case 'srgbClr':
      hex = attr(c, 'val');
      break;
    case 'schemeClr': {
      const v = attr(c, 'val') ?? '';
      hex = theme.colors[(SCHEME_ALIASES[v] ?? v) as ThemeSlot];
      break;
    }
    case 'sysClr':
      hex = attr(c, 'lastClr') ?? (attr(c, 'val') === 'window' ? 'FFFFFF' : '000000');
      break;
    case 'prstClr':
      hex = PRESET_COLORS[attr(c, 'val') ?? ''];
      break;
    case 'scrgbClr':
      hex = ['r', 'g', 'b'].map((k) => hex2(((numAttr(c, k) ?? 0) / 100000) * 255)).join('');
      break;
    default:
      return undefined;
  }
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return undefined;
  let r = parseInt(hex.slice(0, 2), 16);
  let g = parseInt(hex.slice(2, 4), 16);
  let b = parseInt(hex.slice(4, 6), 16);
  for (const m of elements(c)) {
    const v = (numAttr(m, 'val') ?? 100000) / 100000;
    switch (local(m.name)) {
      case 'lumMod':
      case 'lumOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        [r, g, b] = hslToRgb(h, s, Math.max(0, Math.min(1, local(m.name) === 'lumMod' ? l * v : l + v)));
        break;
      }
      case 'shade':
        r *= v;
        g *= v;
        b *= v;
        break;
      case 'tint':
        r = r * v + 255 * (1 - v);
        g = g * v + 255 * (1 - v);
        b = b * v + 255 * (1 - v);
        break;
      default:
        break;
    }
  }
  return `#${hex2(r)}${hex2(g)}${hex2(b)}`;
}

/** Fill colour of shape properties: a colour, null for "no fill", undefined when not set there. */
function fillOf(spPr: XmlElement | undefined, theme: Theme): string | null | undefined {
  if (!spPr) return undefined;
  if (kid(spPr, 'noFill')) return null;
  const solid = kid(spPr, 'solidFill');
  if (solid) return colorOf(elements(solid)[0], theme);
  const grad = kid(spPr, 'gradFill');
  if (grad) return colorOf(elements(deep(grad, 'gs')[0])[0], theme);
  const patt = kid(spPr, 'pattFill');
  if (patt) return colorOf(elements(kid(patt, 'fgClr'))[0], theme);
  return undefined;
}

function dashOf(ln: XmlElement | undefined): LineDash | undefined {
  const v = attr(kid(ln, 'prstDash'), 'val');
  if (!v || v === 'solid') return undefined;
  if (/dot$/i.test(v) && !/dash/i.test(v)) return 'dotted';
  if (/dashdot/i.test(v)) return 'dash-dot';
  return 'dashed';
}

const PRST_DASH: Record<LineDash, string> = { solid: 'solid', dashed: 'dash', dotted: 'sysDot', 'dash-dot': 'dashDot' };

function accents(theme: Theme): string[] {
  return (['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as ThemeSlot[]).map((k) => `#${theme.colors[k]}`);
}

// -------------------------------------------------------------------------------------------------
// Reading classic charts

type Settings = Partial<Pick<SheetChart, 'range' | 'seriesIn' | 'headers' | 'titlePosition' | 'subtitle' | 'legendTitle' | 'type' | 'style' | 'stockKind'>>;

function richParagraphs(e: XmlElement | undefined): string[] {
  if (!e) return [];
  return deep(e, 'p').map((p) => deep(p, 't').map((t) => textContent(t)).join(''));
}

function refOf(e: XmlElement | undefined): string | undefined {
  const ref = kid(e, 'strRef') ?? kid(e, 'numRef') ?? kid(e, 'multiLvlStrRef');
  const f = kid(ref, 'f');
  return f ? textContent(f).trim() || undefined : undefined;
}

const TRENDS: Record<string, Trendline> = { linear: 'linear', exp: 'exponential', log: 'logarithmic', poly: 'polynomial', power: 'power', movingAvg: 'moving-average' };
const SIDE_OF: Record<string, Side> = { b: 'bottom', t: 'top', l: 'left', r: 'right', tr: 'right' };

const GROUP_KIND: Record<string, string> = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area', pieChart: 'pie', pie3DChart: 'pie',
  ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter', bubbleChart: 'bubble', radarChart: 'radar', stockChart: 'stock',
  surfaceChart: 'line', surface3DChart: 'line',
};

function groupType(g: XmlElement): ChartType {
  const kind = GROUP_KIND[local(g.name)];
  const grouping = attr(kid(g, 'grouping'), 'val');
  const prefix = grouping === 'stacked' ? 'stacked-' : grouping === 'percentStacked' ? 'percent-' : '';
  switch (kind) {
    case 'bar':
      return `${prefix}${attr(kid(g, 'barDir'), 'val') === 'bar' ? 'bar' : 'column'}` as ChartType;
    case 'area':
      return `${prefix}area` as ChartType;
    case 'line':
      return `${prefix}line` as ChartType;
    case 'radar':
      return attr(kid(g, 'radarStyle'), 'val') === 'filled' ? 'filled-radar' : 'radar';
    default:
      return kind as ChartType;
  }
}

/** How a chart group draws in a combo chart, or null for groups combos cannot hold. */
function comboKind(t: ChartType): SeriesKind | null {
  switch (t) {
    case 'column':
    case 'bar':
      return 'column';
    case 'stacked-column':
    case 'stacked-bar':
    case 'percent-column':
    case 'percent-bar':
      return 'stacked-column';
    case 'line':
    case 'stacked-line':
    case 'percent-line':
      return 'line';
    case 'area':
      return 'area';
    case 'stacked-area':
    case 'percent-area':
      return 'stacked-area';
    default:
      return null;
  }
}

const ERR_TYPES: Record<string, ErrorBars> = { stdErr: 'stderr', percentage: 'percent', stdDev: 'stddev' };

/** Chart settings from a classic chart part, or null when it has no plot the editor can show. */
export function readChartPart(root: XmlElement, theme: Theme, name: string): Omit<SheetChart, 'id' | 'z' | 'from' | 'to'> | null {
  const chartEl = kid(root, 'chart');
  const plotArea = kid(chartEl, 'plotArea');
  const groups = elements(plotArea).filter((g) => GROUP_KIND[local(g.name)]);
  if (!groups.length) return null;
  const stock = groups.some((g) => GROUP_KIND[local(g.name)] === 'stock');
  // Volume columns come first in volume stock charts.
  if (stock) groups.sort((a, b) => (GROUP_KIND[local(a.name)] === 'stock' ? 1 : 0) - (GROUP_KIND[local(b.name)] === 'stock' ? 1 : 0));
  const combo = !stock && groups.length > 1 && groups.every((g) => comboKind(groupType(g)) !== null) && new Set(groups.map((g) => comboKind(groupType(g)))).size > 1;
  const type: ChartType = stock ? 'stock' : combo ? 'combo' : groupType(groups[0]);
  const lineLike = (t: ChartType) => t.endsWith('line') || t === 'scatter' || t === 'radar' || t === 'stock';
  const axes = elements(plotArea).filter((a) => /Ax$/.test(local(a.name)));
  const axisById = new Map(axes.map((a) => [attr(kid(a, 'axId'), 'val'), a]));
  const firstAxIds = kids(groups[0], 'axId').map((a) => attr(a, 'val'));
  const palette = accents(theme);
  const series: ChartSeries[] = [];
  let categories: string | undefined;
  let valueLabels = false;
  let pieColors: string[] | undefined;
  let errorBars: ErrorBars | undefined;
  let scatterStyle: ScatterStyle | undefined;
  let stockPrices = 0;
  groups.forEach((g, gi) => {
    const gType = groupType(g);
    const axIds = kids(g, 'axId').map((a) => attr(a, 'val'));
    const secondary = gi > 0 && axIds.some((id) => !firstAxIds.includes(id));
    if (labelsShown(kid(g, 'dLbls'))) valueLabels = true;
    for (const ser of kids(g, 'ser')) {
      const spPr = kid(ser, 'spPr');
      const ln = kid(spPr, 'ln');
      const color = lineLike(gType) ? fillOf(ln, theme) ?? fillOf(spPr, theme) : fillOf(spPr, theme) ?? fillOf(ln, theme);
      const tx = kid(ser, 'tx');
      const xy = gType === 'scatter' || gType === 'bubble';
      const values = refOf(kid(ser, xy ? 'yVal' : 'val'));
      if (!values) continue;
      const s: ChartSeries = {
        values,
        nameRef: refOf(tx),
        name: refOf(tx) ? undefined : textContent(kid(tx, 'v') ?? { type: 'text', text: '' }) || undefined,
        color: color ?? palette[series.length % palette.length],
      };
      if (xy) s.xValues = refOf(kid(ser, 'xVal'));
      if (gType === 'bubble') s.sizes = refOf(kid(ser, 'bubbleSize'));
      if (!xy && !categories) categories = refOf(kid(ser, 'cat'));
      const w = numAttr(ln, 'w');
      if (w && lineLike(gType)) s.lineWidth = Math.round((w / EMU_PER_PX) * 2) / 2;
      const dash = dashOf(ln);
      if (dash) s.lineStyle = dash;
      const markerOn = attr(kid(kid(ser, 'marker'), 'symbol'), 'val') !== 'none';
      if (lineLike(gType) && gType !== 'stock') s.markers = markerOn;
      if (attr(kid(ser, 'smooth'), 'val') === '1') s.smooth = true;
      if (gType === 'scatter' && !scatterStyle) {
        const lines = !kid(ln, 'noFill');
        scatterStyle = !lines ? 'markers' : s.smooth ? (markerOn ? 'smooth-markers' : 'smooth') : markerOn ? 'lines-markers' : 'lines';
      }
      const err = attr(kid(kid(ser, 'errBars'), 'errValType'), 'val');
      if (err && ERR_TYPES[err] && !errorBars) errorBars = ERR_TYPES[err];
      const trend = attr(kid(kid(ser, 'trendline'), 'trendlineType'), 'val');
      if (trend && TRENDS[trend]) s.trendline = TRENDS[trend];
      if (secondary && !stock) s.secondaryAxis = true;
      if (combo) s.kind = comboKind(gType) ?? 'column';
      if (gType === 'stock') stockPrices++;
      if (labelsShown(kid(ser, 'dLbls'))) valueLabels = true;
      if ((gType === 'pie' || gType === 'doughnut') && !pieColors) {
        const pts = kids(ser, 'dPt');
        if (pts.length) {
          pieColors = palette.slice();
          for (const p of pts) {
            const idx = numAttr(kid(p, 'idx'), 'val') ?? 0;
            const c = fillOf(kid(p, 'spPr'), theme);
            if (c) pieColors[idx] = c;
          }
          for (let i = 0; i < pieColors.length; i++) pieColors[i] ??= palette[i % palette.length];
        }
      }
      series.push(s);
    }
  });
  if (!series.length) return null;
  const valAx = firstAxIds.map((id) => axisById.get(id)).find((a) => a && local(a.name) === 'valAx' && (type !== 'scatter' && type !== 'bubble' ? true : ['l', 'r'].includes(attr(kid(a, 'axPos'), 'val') ?? '')));
  const catAx = firstAxIds.map((id) => axisById.get(id)).find((a) => a && a !== valAx);
  const secAx = axes.find((a) => local(a.name) === 'valAx' && a !== valAx && !firstAxIds.includes(attr(kid(a, 'axId'), 'val')));
  const titleParas = richParagraphs(kid(kid(kid(chartEl, 'title'), 'tx'), 'rich'));
  const legend = kid(chartEl, 'legend');
  const settings = readSettings(root);
  const scaling = kid(valAx, 'scaling');
  const vertical = (() => {
    const r = series[0].values ? parseRange(splitSheetRef(series[0].values).ref.replace(/\$/g, '')) : null;
    return r ? r.r1 > r.r0 || r.c1 === r.c0 : true;
  })();
  const volume = stock && groups.some((g) => GROUP_KIND[local(g.name)] === 'bar');
  // Chart area: white with a light grey line is the default look.
  const areaPr = kid(root, 'spPr');
  const areaFill = fillOf(areaPr, theme);
  const areaLn = kid(areaPr, 'ln');
  let border: SheetChart['border'];
  if (kid(areaLn, 'noFill')) border = null;
  else {
    const c = fillOf(areaLn, theme);
    const w = numAttr(areaLn, 'w');
    const width = w ? Math.max(0.5, Math.round((w / EMU_PER_PX) * 2) / 2) : 1;
    if (c && (c.toUpperCase() !== '#D9D9D9' || width !== 1 || dashOf(areaLn))) border = { color: c, width, dash: dashOf(areaLn) };
  }
  return {
    name,
    type: settings.type ?? type,
    range: settings.range ?? '',
    seriesIn: settings.seriesIn ?? (vertical ? 'columns' : 'rows'),
    headers: settings.headers ?? !!series[0].nameRef,
    categories,
    series,
    title: titleParas[0] || undefined,
    subtitle: settings.subtitle ?? (titleParas.slice(1).join(' ') || undefined),
    titlePosition: settings.titlePosition,
    legendTitle: settings.legendTitle,
    legendPosition: legend ? SIDE_OF[attr(kid(legend, 'legendPos'), 'val') ?? 'r'] ?? 'right' : undefined,
    showLegend: !!legend,
    gridlines: !!kid(valAx, 'majorGridlines'),
    vGridlines: !!kid(catAx, 'majorGridlines') || undefined,
    hideXAxis: attr(kid(catAx, 'delete'), 'val') === '1' || undefined,
    hideYAxis: attr(kid(valAx, 'delete'), 'val') === '1' || undefined,
    valueLabels,
    xAxisTitle: richParagraphs(kid(kid(kid(catAx, 'title'), 'tx'), 'rich')).join(' ') || undefined,
    yAxisTitle: richParagraphs(kid(kid(kid(valAx, 'title'), 'tx'), 'rich')).join(' ') || undefined,
    secondaryAxisTitle: richParagraphs(kid(kid(kid(secAx, 'title'), 'tx'), 'rich')).join(' ') || undefined,
    yMin: numAttr(kid(scaling, 'min'), 'val'),
    yMax: numAttr(kid(scaling, 'max'), 'val'),
    palette: pieColors ?? palette,
    errorBars,
    scatterStyle,
    stockKind: stock ? settings.stockKind ?? (volume ? (stockPrices >= 4 ? 'vohlc' : 'vhlc') : stockPrices >= 4 ? 'ohlc' : 'hlc') : undefined,
    style: settings.style,
    fill: areaFill === undefined || areaFill?.toUpperCase() === '#FFFFFF' ? undefined : areaFill,
    border,
  };
}

function labelsShown(dLbls: XmlElement | undefined): boolean {
  return attr(kid(dLbls, 'showVal'), 'val') === '1' || attr(kid(dLbls, 'showPercent'), 'val') === '1';
}

function readSettings(root: XmlElement): Settings {
  for (const ext of deep(kid(root, 'extLst'), 'ext')) {
    if (attr(ext, 'uri') !== SETTINGS_URI) continue;
    try {
      const parsed = JSON.parse(textContent(ext)) as Settings;
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch {
      return {};
    }
  }
  return {};
}

// -------------------------------------------------------------------------------------------------
// Reading chartex (histogram, pareto, box and whisker, waterfall, funnel, treemap)

export function readChartexPart(root: XmlElement, theme: Theme, name: string): Omit<SheetChart, 'id' | 'z' | 'from' | 'to'> | null {
  const data = new Map<string, { cat?: string; val?: string }>();
  for (const d of kids(kid(root, 'chartData'), 'data')) {
    const entry: { cat?: string; val?: string } = {};
    for (const dim of elements(d)) {
      const f = kid(dim, 'f');
      const ref = f ? textContent(f).trim() : undefined;
      const kind = attr(dim, 'type');
      if (local(dim.name) === 'strDim' && kind === 'cat') entry.cat = ref;
      if (local(dim.name) === 'numDim' && (kind === 'val' || kind === 'size')) entry.val = ref;
    }
    data.set(attr(d, 'id') ?? '0', entry);
  }
  const chartEl = kid(root, 'chart');
  const all = deep(kid(chartEl, 'plotArea'), 'series');
  const layouts = all.map((s) => attr(s, 'layoutId'));
  let type: ChartType;
  switch (layouts[0]) {
    case 'clusteredColumn':
      type = layouts.includes('paretoLine') ? 'pareto' : 'histogram';
      break;
    case 'boxWhisker':
      type = 'box-whisker';
      break;
    case 'waterfall':
      type = 'waterfall';
      break;
    case 'funnel':
      type = 'funnel';
      break;
    case 'treemap':
      type = 'treemap';
      break;
    case 'sunburst':
      type = 'sunburst';
      break;
    default:
      return null;
  }
  const palette = accents(theme);
  let categories: string | undefined;
  const series: ChartSeries[] = [];
  for (const s of all) {
    if (attr(s, 'layoutId') === 'paretoLine') continue;
    const d = data.get(attr(kid(s, 'dataId'), 'val') ?? '0');
    if (!d?.val) continue;
    categories ??= d.cat;
    const txData = kid(kid(s, 'tx'), 'txData');
    const nameRef = kid(txData, 'f') ? textContent(kid(txData, 'f')!).trim() : undefined;
    series.push({
      values: d.val,
      nameRef,
      name: nameRef ? undefined : textContent(kid(txData, 'v') ?? { type: 'text', text: '' }) || undefined,
      color: fillOf(kid(s, 'spPr'), theme) ?? palette[series.length % palette.length],
    });
  }
  if (!series.length) return null;
  const title = kid(chartEl, 'title');
  const titleText = title ? deep(title, 'v').map((v) => textContent(v)).join('') || deep(title, 't').map((t) => textContent(t)).join('') : '';
  const legend = kid(chartEl, 'legend');
  return {
    name,
    type,
    range: '',
    seriesIn: 'columns',
    headers: !!series[0].nameRef,
    categories,
    series,
    title: titleText || undefined,
    legendPosition: legend ? SIDE_OF[attr(legend, 'pos') ?? 'b'] ?? 'bottom' : undefined,
    showLegend: !!legend,
    gridlines: deep(kid(chartEl, 'plotArea'), 'majorGridlines').length > 0,
    valueLabels: deep(kid(chartEl, 'plotArea'), 'visibility').some((v) => attr(v, 'value') === '1'),
    palette,
  };
}

// -------------------------------------------------------------------------------------------------
// Reading shapes and connectors

/** A shape or connector with a preset geometry, or null for anything else (groups, freeforms...). */
export function readShape(node: XmlElement, theme: Theme): Omit<SheetShape, 'id' | 'z' | 'from' | 'to'> | null {
  const kind = local(node.name);
  if (kind !== 'sp' && kind !== 'cxnSp') return null;
  const spPr = kid(node, 'spPr');
  const preset = attr(kid(spPr, 'prstGeom'), 'prst');
  if (!preset) return null;
  const nv = kid(node, kind === 'sp' ? 'nvSpPr' : 'nvCxnSpPr');
  const cNvPr = kid(nv, 'cNvPr');
  const style = kid(node, 'style');
  const styleColor = (ref: string) => colorOf(elements(kid(style, ref))[0], theme);
  const isLine = kind === 'cxnSp' || preset === 'line' || /Connector/.test(preset);
  let fill = fillOf(spPr, theme);
  if (fill === undefined) fill = isLine ? null : styleColor('fillRef') ?? null;
  const ln = kid(spPr, 'ln');
  let line: SheetShape['line'];
  if (kid(ln, 'noFill')) line = null;
  else {
    const color = fillOf(ln, theme) ?? styleColor('lnRef') ?? (isLine ? '#4472C4' : undefined);
    const w = numAttr(ln, 'w');
    line = color ? { color, width: w ? Math.max(0.5, Math.round((w / EMU_PER_PX) * 2) / 2) : style ? 1.5 : 1, dash: dashOf(ln) } : null;
  }
  const body = kid(node, 'txBody');
  const paras = body ? deep(body, 'p').map((p) => deep(p, 't').map((t) => textContent(t)).join('')) : [];
  const text = paras.join('\n').replace(/\n+$/, '');
  const runs = body ? deep(body, 'rPr') : [];
  const runColor = runs.map((r) => colorOf(elements(kid(r, 'solidFill'))[0], theme)).find(Boolean);
  const runSize = runs.map((r) => numAttr(r, 'sz')).find((v) => v !== undefined);
  const xfrm = kid(spPr, 'xfrm');
  const rot = numAttr(xfrm, 'rot');
  const end = (n: string) => {
    const t = attr(kid(ln, n), 'type');
    return !!t && t !== 'none';
  };
  return {
    name: attr(cNvPr, 'name') ?? 'Shape',
    preset,
    fill,
    line,
    fontColor: runColor ?? colorOf(elements(kid(style, 'fontRef'))[0], theme) ?? '#000000',
    text: text || undefined,
    arrowStart: end('headEnd') || undefined,
    arrowEnd: end('tailEnd') || undefined,
    rotation: rot ? rot / 60000 : undefined,
    flipH: attr(xfrm, 'flipH') === '1' || undefined,
    flipV: attr(xfrm, 'flipV') === '1' || undefined,
    textBox: attr(kid(nv, 'cNvSpPr'), 'txBox') === '1' || undefined,
    bold: runs.some((r) => attr(r, 'b') === '1') || undefined,
    italic: runs.some((r) => attr(r, 'i') === '1') || undefined,
    fontSize: runSize ? runSize / 100 : undefined,
    altTitle: attr(cNvPr, 'title') || undefined,
    descr: attr(cNvPr, 'descr') || undefined,
    hidden: attr(cNvPr, 'hidden') === '1' || undefined,
  };
}

/** The graphic frame of an anchor (also inside mc:AlternateContent) and the chart it references. */
export function chartReference(anchor: XmlElement): { rid: string; chartex: boolean } | null {
  for (const frame of deep(anchor, 'graphicFrame')) {
    const data = deep(frame, 'graphicData')[0];
    const uri = attr(data, 'uri') ?? '';
    const ref = elements(data)[0];
    const rid = attr(ref, 'r:id');
    if (!rid) continue;
    if (uri === NS_C) return { rid, chartex: false };
    if (/chartex/i.test(uri) || local(ref?.name ?? '') === 'chart') return { rid, chartex: /chartex/i.test(uri) };
  }
  return null;
}

/** The content element of an anchor: the shape, picture, frame or group after from/to/ext. */
export function anchorContent(anchor: XmlElement): XmlElement | undefined {
  return elements(anchor).find((e) => !['from', 'to', 'ext', 'pos', 'clientData'].includes(local(e.name)));
}

// -------------------------------------------------------------------------------------------------
// Values of references at save time (cached values written into chart parts)

interface RefData {
  texts: string[];
  numbers: (number | null)[];
  numFmt: string;
}

export function readRefData(wb: Workbook, ref: string | undefined): RefData {
  const empty = { texts: [], numbers: [], numFmt: 'General' };
  if (!ref) return empty;
  const { sheet: name, ref: local } = splitSheetRef(ref);
  const sheet = name !== undefined ? wb.sheetByName(name) : undefined;
  const range = parseRange(local.replace(/\$/g, '').toUpperCase());
  if (!sheet || !range) return empty;
  const ext = sheet.dataExtent();
  const r1 = Math.min(range.r1, Math.max(range.r0, ext.rows));
  const c1 = Math.min(range.c1, Math.max(range.c0, ext.cols));
  const texts: string[] = [];
  const numbers: (number | null)[] = [];
  let numFmt = '';
  for (let r = range.r0; r <= r1; r++) {
    for (let c = range.c0; c <= c1; c++) {
      const v = sheet.getCell(r, c)?.v ?? null;
      const fmt = wb.styles.get(sheet.styleAt(r, c)).numFmt || 'General';
      if (!numFmt) numFmt = fmt;
      numbers.push(typeof v === 'number' && Number.isFinite(v) ? v : null);
      texts.push(v === null ? '' : typeof v === 'object' ? String((v as { code?: string }).code ?? '') : formatValue(v, fmt, wb.date1904).text);
      if (texts.length >= 10000) return { texts, numbers, numFmt: numFmt || 'General' };
    }
  }
  return { texts, numbers, numFmt: numFmt || 'General' };
}

// -------------------------------------------------------------------------------------------------
// Writing classic charts

const x = (tag: string, attrs: Record<string, string | number | undefined> = {}, inner = ''): string => {
  const a = Object.entries(attrs)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => ` ${k}="${escapeAttr(String(v))}"`)
    .join('');
  return inner ? `<${tag}${a}>${inner}</${tag}>` : `<${tag}${a}/>`;
};
const val = (tag: string, v: string | number) => x(tag, { val: v });

function richText(paras: { text: string; size?: number; bold?: boolean }[]): string {
  return x(
    'c:rich',
    {},
    `<a:bodyPr/><a:lstStyle/>${paras
      .map((p) => `<a:p><a:pPr><a:defRPr${p.size ? ` sz="${p.size}"` : ''}${p.bold === false ? ' b="0"' : ''}/></a:pPr><a:r><a:rPr lang="en-US"${p.size ? ` sz="${p.size}"` : ''}${p.bold === false ? ' b="0"' : ''}/><a:t>${escapeText(p.text)}</a:t></a:r></a:p>`)
      .join('')}`,
  );
}

function titleXml(text: string | undefined, sub?: string): string {
  const paras = [];
  if (text) paras.push({ text, size: 1400 });
  if (sub) paras.push({ text: sub, size: 1100, bold: false });
  if (!paras.length) return '';
  return x('c:title', {}, `${x('c:tx', {}, richText(paras))}${val('c:overlay', 0)}`);
}

function strCache(texts: string[]): string {
  return x('c:strCache', {}, `${val('c:ptCount', texts.length)}${texts.map((t, i) => (t === '' ? '' : x('c:pt', { idx: i }, `<c:v>${escapeText(t)}</c:v>`))).join('')}`);
}

function numCache(nums: (number | null)[], fmt: string): string {
  return x('c:numCache', {}, `<c:formatCode>${escapeText(fmt || 'General')}</c:formatCode>${val('c:ptCount', nums.length)}${nums.map((n, i) => (n === null ? '' : x('c:pt', { idx: i }, `<c:v>${n}</c:v>`))).join('')}`);
}

function strRef(wb: Workbook, ref: string): string {
  return x('c:strRef', {}, `<c:f>${escapeText(ref)}</c:f>${strCache(readRefData(wb, ref).texts)}`);
}

function numRef(wb: Workbook, ref: string): string {
  const d = readRefData(wb, ref);
  return x('c:numRef', {}, `<c:f>${escapeText(ref)}</c:f>${numCache(d.numbers, d.numFmt)}`);
}

function solid(hex: string): string {
  return `<a:solidFill><a:srgbClr val="${hex.replace('#', '').toUpperCase()}"/></a:solidFill>`;
}

function seriesColor(chart: SheetChart, s: ChartSeries, i: number): string {
  return (s.color ?? `#${(chart.palette?.[i] ?? EDITOR_COLORS[i % EDITOR_COLORS.length]).replace('#', '')}`).replace('#', '').toUpperCase();
}

const TREND_TYPE: Record<Trendline, string> = { none: '', linear: 'linear', exponential: 'exp', logarithmic: 'log', polynomial: 'poly', power: 'power', 'moving-average': 'movingAvg' };

function trendlineXml(s: ChartSeries, color: string): string {
  const t = s.trendline && s.trendline !== 'none' ? TREND_TYPE[s.trendline] : '';
  if (!t) return '';
  const extra = t === 'poly' ? val('c:order', 2) : t === 'movingAvg' ? val('c:period', 2) : '';
  return x('c:trendline', {}, `${x('c:spPr', {}, `<a:ln w="19050" cap="rnd">${solid(color)}<a:prstDash val="sysDot"/></a:ln>`)}${val('c:trendlineType', t)}${extra}${val('c:dispRSqr', 0)}${val('c:dispEq', 0)}`);
}

function dLblsXml(show: boolean, percent = false): string {
  if (!show) return '';
  return x('c:dLbls', {}, `${val('c:showLegendKey', 0)}${val('c:showVal', percent ? 0 : 1)}${val('c:showCatName', 0)}${val('c:showSerName', 0)}${val('c:showPercent', percent ? 1 : 0)}${val('c:showBubbleSize', 0)}`);
}

function lnXml(color: string, widthPx: number | undefined, dash: LineDash | undefined, none = false): string {
  if (none) return '<a:ln><a:noFill/></a:ln>';
  const w = Math.round((widthPx ?? 2.25) * EMU_PER_PX);
  return `<a:ln w="${w}" cap="rnd">${solid(color)}${dash && dash !== 'solid' ? `<a:prstDash val="${PRST_DASH[dash]}"/>` : ''}<a:round/></a:ln>`;
}

const ERR_VAL_TYPE: Record<Exclude<ErrorBars, 'none'>, string> = { stderr: 'stdErr', percent: 'percentage', stddev: 'stdDev' };

function errBarsXml(kind: ErrorBars | undefined, dir?: 'x' | 'y'): string {
  if (!kind || kind === 'none') return '';
  const amount = kind === 'percent' ? val('c:val', 5) : kind === 'stddev' ? val('c:val', 1) : '';
  return x('c:errBars', {}, `${dir ? val('c:errDir', dir) : ''}${val('c:errBarType', 'both')}${val('c:errValType', ERR_VAL_TYPE[kind])}${val('c:noEndCap', 0)}${amount}${x('c:spPr', {}, `<a:ln w="9525">${solid('595959')}</a:ln>`)}`);
}

function scatterLook(chart: SheetChart, s: ChartSeries): { lines: boolean; markers: boolean; smooth: boolean } {
  const style = chart.scatterStyle ?? (s.lineWidth ? 'lines-markers' : 'markers');
  return { lines: style !== 'markers', markers: style === 'markers' || style.endsWith('-markers'), smooth: style.startsWith('smooth') || !!s.smooth };
}

function serXml(wb: Workbook, chart: SheetChart, s: ChartSeries, i: number, kind: string): string {
  const color = seriesColor(chart, s, i);
  const tx = s.nameRef ? x('c:tx', {}, strRef(wb, s.nameRef)) : x('c:tx', {}, `<c:v>${escapeText(s.name ?? `Series ${i + 1}`)}</c:v>`);
  const head = `${val('c:idx', i)}${val('c:order', i)}${tx}`;
  const cat = chart.categories ? x('c:cat', {}, strRef(wb, chart.categories)) : '';
  const values = x('c:val', {}, numRef(wb, s.values));
  const marker = (on: boolean) => (on ? x('c:marker', {}, `${val('c:symbol', 'circle')}${val('c:size', 5)}${x('c:spPr', {}, `${solid(color)}<a:ln w="9525">${solid(color)}</a:ln>`)}`) : x('c:marker', {}, val('c:symbol', 'none')));
  const err = chart.type.startsWith('percent-') ? '' : errBarsXml(chart.errorBars);
  switch (kind) {
    case 'bar':
      return x('c:ser', {}, `${head}${x('c:spPr', {}, solid(color))}${val('c:invertIfNegative', 0)}${trendlineXml(s, color)}${err}${cat}${values}`);
    case 'line':
      return x('c:ser', {}, `${head}${x('c:spPr', {}, lnXml(color, s.lineWidth, s.lineStyle))}${marker(s.markers ?? !!chart.markers)}${trendlineXml(s, color)}${err}${cat}${values}${val('c:smooth', s.smooth ? 1 : 0)}`);
    case 'area':
      return x('c:ser', {}, `${head}${x('c:spPr', {}, solid(color))}${trendlineXml(s, color)}${err}${cat}${values}`);
    case 'radar':
      return x('c:ser', {}, `${head}${x('c:spPr', {}, chart.type === 'filled-radar' ? `${solid(color)}${lnXml(color, s.lineWidth, s.lineStyle)}` : lnXml(color, s.lineWidth, s.lineStyle))}${marker(chart.type !== 'filled-radar' && (s.markers ?? !!chart.markers))}${cat}${values}`);
    case 'pie': {
      const n = chart.categories ? readRefData(wb, chart.categories).texts.length : readRefData(wb, s.values).numbers.length;
      const colors = chart.palette?.length ? chart.palette : EDITOR_COLORS.map((c) => `#${c}`);
      const pts = Array.from({ length: n }, (_, k) => x('c:dPt', {}, `${val('c:idx', k)}${val('c:bubble3D', 0)}${x('c:spPr', {}, `${solid(colors[k % colors.length])}<a:ln w="19050">${solid('FFFFFF')}</a:ln>`)}`)).join('');
      return x('c:ser', {}, `${head}${pts}${cat}${values}`);
    }
    case 'scatter': {
      const xv = s.xValues ? x('c:xVal', {}, numRef(wb, s.xValues)) : '';
      const look = scatterLook(chart, s);
      return x('c:ser', {}, `${head}${x('c:spPr', {}, lnXml(color, s.lineWidth, s.lineStyle, !look.lines))}${marker(look.markers)}${trendlineXml(s, color)}${errBarsXml(chart.errorBars, 'y')}${xv}${x('c:yVal', {}, numRef(wb, s.values))}${val('c:smooth', look.smooth ? 1 : 0)}`);
    }
    case 'bubble': {
      const xv = s.xValues ? x('c:xVal', {}, numRef(wb, s.xValues)) : '';
      const size = s.sizes ? x('c:bubbleSize', {}, numRef(wb, s.sizes)) : '';
      return x('c:ser', {}, `${head}${x('c:spPr', {}, `${solid(color)}<a:ln><a:noFill/></a:ln>`)}${val('c:invertIfNegative', 0)}${xv}${x('c:yVal', {}, numRef(wb, s.values))}${size}${val('c:bubble3D', 0)}`);
    }
    case 'stock':
      return x('c:ser', {}, `${head}${x('c:spPr', {}, '<a:ln w="19050"><a:noFill/></a:ln>')}${x('c:marker', {}, val('c:symbol', 'none'))}${cat}${values}${val('c:smooth', 0)}`);
    default:
      return '';
  }
}

function axisTitle(text: string | undefined, ink: string): string {
  return text ? x('c:title', {}, `${x('c:tx', {}, richText([{ text, size: 1000, bold: false }]))}${val('c:overlay', 0)}${textPr(ink, 1000)}`) : '';
}

function textPr(ink: string, size = 900): string {
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}"><a:solidFill><a:srgbClr val="${ink}"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

interface AxisOpts {
  title?: string;
  gridlines?: boolean;
  min?: number;
  max?: number;
  percent?: boolean;
  crossesMax?: boolean;
  deleted?: boolean;
  ink: string;
  grid: string;
}

function catAxXml(id: number, cross: number, pos: string, opts: AxisOpts): string {
  const grid = opts.gridlines ? x('c:majorGridlines', {}, x('c:spPr', {}, `<a:ln w="9525">${solid(opts.grid)}</a:ln>`)) : '';
  return x(
    'c:catAx',
    {},
    `${val('c:axId', id)}${x('c:scaling', {}, val('c:orientation', 'minMax'))}${val('c:delete', opts.deleted ? 1 : 0)}${val('c:axPos', pos)}${grid}${axisTitle(opts.deleted ? undefined : opts.title, opts.ink)}${x('c:numFmt', { formatCode: 'General', sourceLinked: 1 })}${val('c:majorTickMark', 'none')}${val('c:minorTickMark', 'none')}${val('c:tickLblPos', 'nextTo')}${x('c:spPr', {}, `<a:ln w="9525">${solid(opts.grid)}</a:ln>`)}${textPr(opts.ink)}${val('c:crossAx', cross)}${val('c:crosses', 'autoZero')}${val('c:auto', 1)}${val('c:lblAlgn', 'ctr')}${val('c:lblOffset', 100)}${val('c:noMultiLvlLbl', 0)}`,
  );
}

function valAxXml(id: number, cross: number, pos: string, opts: AxisOpts): string {
  const scaling = `${val('c:orientation', 'minMax')}${opts.max !== undefined ? val('c:max', opts.max) : ''}${opts.min !== undefined ? val('c:min', opts.min) : ''}`;
  return x(
    'c:valAx',
    {},
    `${val('c:axId', id)}${x('c:scaling', {}, scaling)}${val('c:delete', opts.deleted ? 1 : 0)}${val('c:axPos', pos)}${opts.gridlines ? x('c:majorGridlines', {}, x('c:spPr', {}, `<a:ln w="9525">${solid(opts.grid)}</a:ln>`)) : ''}${axisTitle(opts.title, opts.ink)}${x('c:numFmt', { formatCode: opts.percent ? '0%' : 'General', sourceLinked: opts.percent ? 0 : 1 })}${val('c:majorTickMark', 'none')}${val('c:minorTickMark', 'none')}${val('c:tickLblPos', 'nextTo')}${x('c:spPr', {}, '<a:ln><a:noFill/></a:ln>')}${textPr(opts.ink)}${val('c:crossAx', cross)}${opts.crossesMax ? val('c:crosses', 'max') : val('c:crosses', 'autoZero')}${val('c:crossBetween', 'between')}`,
  );
}

const LEGEND_POS: Record<Side, string> = { bottom: 'b', top: 't', left: 'l', right: 'r' };

function isDark(hex: string): boolean {
  const n = parseInt(hex.replace('#', ''), 16);
  return Number.isFinite(n) && (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255 < 0.45;
}

interface ChartGroup {
  kind: string;
  grouping?: string;
  dir?: 'bar' | 'col';
  secondary: boolean;
  list: { s: ChartSeries; i: number }[];
}

/** Series grouped the way chart parts hold them: one group per chart kind, stacking and axis. */
function chartGroups(chart: SheetChart): ChartGroup[] {
  const type = chart.type;
  const groups: ChartGroup[] = [];
  const add = (g: Omit<ChartGroup, 'list'>, e: { s: ChartSeries; i: number }) => {
    let found = groups.find((o) => o.kind === g.kind && o.grouping === g.grouping && o.dir === g.dir && o.secondary === g.secondary);
    if (!found) groups.push((found = { ...g, list: [] }));
    found.list.push(e);
  };
  const entries = chart.series.map((s, i) => ({ s, i }));
  if (type === 'combo') {
    for (const e of entries) {
      const k: SeriesKind = e.s.kind ?? (e.i === entries.length - 1 && entries.length > 1 ? 'line' : 'column');
      const secondary = !!e.s.secondaryAxis;
      if (k === 'line') add({ kind: 'line', grouping: 'standard', secondary }, e);
      else if (k === 'area') add({ kind: 'area', grouping: 'standard', secondary }, e);
      else if (k === 'stacked-area') add({ kind: 'area', grouping: 'stacked', secondary }, e);
      else add({ kind: 'bar', dir: 'col', grouping: k === 'stacked-column' ? 'stacked' : 'clustered', secondary }, e);
    }
  } else if (type === 'stock') {
    const sk = chart.stockKind ?? (entries.length >= 5 ? 'vohlc' : entries.length === 4 ? 'ohlc' : 'hlc');
    const volume = sk === 'vhlc' || sk === 'vohlc';
    if (volume && entries.length) add({ kind: 'bar', dir: 'col', grouping: 'clustered', secondary: false }, entries[0]);
    const from = volume ? 1 : 0;
    for (const e of entries.slice(from, from + (sk === 'ohlc' || sk === 'vohlc' ? 4 : 3))) add({ kind: 'stock', secondary: volume }, e);
  } else {
    const percent = type.startsWith('percent-');
    const stacked = type.startsWith('stacked-');
    const grouping = percent ? 'percentStacked' : stacked ? 'stacked' : undefined;
    const kind = type.endsWith('column') || type.endsWith('bar') ? 'bar' : type.endsWith('area') ? 'area' : type.endsWith('line') ? 'line' : type.endsWith('radar') ? 'radar' : type === 'doughnut' ? 'pie' : type;
    for (const e of entries) {
      const secondary = !!e.s.secondaryAxis && !grouping && !['pie', 'radar'].includes(kind);
      add({ kind, grouping: kind === 'bar' ? grouping ?? 'clustered' : kind === 'line' || kind === 'area' ? grouping ?? 'standard' : grouping, dir: kind === 'bar' ? (type.endsWith('bar') ? 'bar' : 'col') : undefined, secondary }, e);
    }
  }
  return groups;
}

/** A classic chart part (c:chartSpace) for a chart of the editor. */
export function writeChartPart(wb: Workbook, chart: SheetChart): string {
  const type = chart.type;
  const dark = !!chart.fill && isDark(chart.fill);
  const ink = dark ? 'F2F2F2' : '595959';
  const gridInk = dark ? '7F7F7F' : 'D9D9D9';
  const labels = dLblsXml(chart.valueLabels, false);
  const groups = chartGroups(chart);
  const groupXml = (g: ChartGroup, axA: number, axB: number): string => {
    const ax = `${val('c:axId', axA)}${val('c:axId', axB)}`;
    const sers = g.list.map(({ s, i }) => serXml(wb, chart, s, i, g.kind)).join('');
    switch (g.kind) {
      case 'bar':
        return x('c:barChart', {}, `${val('c:barDir', g.dir ?? 'col')}${val('c:grouping', g.grouping ?? 'clustered')}${val('c:varyColors', 0)}${sers}${labels}${val('c:gapWidth', 219)}${g.grouping === 'clustered' ? val('c:overlap', -27) : val('c:overlap', 100)}${ax}`);
      case 'line':
        return x('c:lineChart', {}, `${val('c:grouping', g.grouping ?? 'standard')}${val('c:varyColors', 0)}${sers}${labels}${val('c:marker', 1)}${ax}`);
      case 'area':
        return x('c:areaChart', {}, `${val('c:grouping', g.grouping ?? 'standard')}${val('c:varyColors', 0)}${sers}${labels}${ax}`);
      case 'radar':
        return x('c:radarChart', {}, `${val('c:radarStyle', type === 'filled-radar' ? 'filled' : 'marker')}${val('c:varyColors', 0)}${sers}${labels}${ax}`);
      case 'scatter': {
        const smooth = g.list.some(({ s }) => scatterLook(chart, s).smooth);
        return x('c:scatterChart', {}, `${val('c:scatterStyle', smooth ? 'smoothMarker' : 'lineMarker')}${val('c:varyColors', 0)}${sers}${labels}${ax}`);
      }
      case 'bubble':
        return x('c:bubbleChart', {}, `${val('c:varyColors', 0)}${sers}${labels}${val('c:bubble3D', 0)}${val('c:bubbleScale', 100)}${val('c:showNegBubbles', 0)}${ax}`);
      case 'stock': {
        const updown = g.list.length >= 4 ? x('c:upDownBars', {}, `${val('c:gapWidth', 150)}${x('c:upBars', {}, x('c:spPr', {}, `${solid('FFFFFF')}<a:ln w="9525">${solid('404040')}</a:ln>`))}${x('c:downBars', {}, x('c:spPr', {}, `${solid('404040')}<a:ln w="9525">${solid('404040')}</a:ln>`))}`) : '';
        return x('c:stockChart', {}, `${sers}${labels}${x('c:hiLowLines', {}, x('c:spPr', {}, `<a:ln w="9525">${solid('595959')}</a:ln>`))}${updown}${ax}`);
      }
      default:
        return '';
    }
  };
  let plot: string;
  if (groups[0]?.kind === 'pie') {
    const primary = groups[0].list;
    const body = `${val('c:varyColors', 1)}${primary.slice(0, type === 'doughnut' ? primary.length : 1).map(({ s, i }) => serXml(wb, chart, s, i, 'pie')).join('')}${dLblsXml(chart.valueLabels, true)}${val('c:firstSliceAng', 0)}`;
    plot = type === 'doughnut' ? x('c:doughnutChart', {}, `${body}${val('c:holeSize', 50)}`) : x('c:pieChart', {}, body);
  } else {
    let primary = groups.filter((g) => !g.secondary);
    let secondary = groups.filter((g) => g.secondary);
    if (!primary.length) {
      primary = secondary;
      secondary = [];
    }
    const horizontal = type.endsWith('bar');
    const xy = type === 'scatter' || type === 'bubble';
    const percent = type.startsWith('percent-');
    plot = primary.map((g) => groupXml(g, 500000001, 500000002)).join('');
    plot += secondary.map((g) => groupXml(g, 500000003, 500000004)).join('');
    const base = { ink, grid: gridInk };
    const catAx = xy
      ? valAxXml(500000001, 500000002, 'b', { ...base, title: chart.xAxisTitle, gridlines: !!chart.vGridlines, deleted: chart.hideXAxis })
      : catAxXml(500000001, 500000002, horizontal ? 'l' : 'b', { ...base, title: chart.xAxisTitle, gridlines: !!chart.vGridlines, deleted: chart.hideXAxis });
    const valAx = valAxXml(500000002, 500000001, horizontal ? 'b' : 'l', { ...base, title: chart.yAxisTitle, gridlines: chart.gridlines, min: chart.yMin, max: chart.yMax, percent, deleted: chart.hideYAxis });
    plot += catAx + valAx;
    if (secondary.length) {
      plot += xy ? valAxXml(500000003, 500000004, 'b', { ...base, deleted: true }) : catAxXml(500000003, 500000004, 'b', { ...base, deleted: true });
      plot += valAxXml(500000004, 500000003, 'r', { ...base, title: chart.secondaryAxisTitle, crossesMax: true });
    }
  }
  const title = titleXml(chart.title, chart.subtitle);
  const legend = chart.showLegend ? x('c:legend', {}, `${val('c:legendPos', LEGEND_POS[chart.legendPosition ?? 'bottom'])}${val('c:overlay', 0)}${textPr(ink)}`) : '';
  const settings: Settings = { range: chart.range || undefined, seriesIn: chart.seriesIn, headers: chart.headers, titlePosition: chart.titlePosition, subtitle: chart.subtitle, legendTitle: chart.legendTitle, type: chart.type, style: chart.style, stockKind: chart.stockKind };
  const ext = x('c:extLst', {}, x('c:ext', { uri: SETTINGS_URI, 'xmlns:xe': NS_XE }, `<xe:chartSettings>${escapeText(JSON.stringify(settings))}</xe:chartSettings>`));
  const areaFill = chart.fill === undefined ? solid('FFFFFF') : chart.fill === null ? '<a:noFill/>' : solid(chart.fill);
  const areaLine = chart.border === undefined ? `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${solid('D9D9D9')}<a:round/></a:ln>` : chart.border === null ? '<a:ln><a:noFill/></a:ln>' : lnXml(chart.border.color, chart.border.width, chart.border.dash);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${x(
    'c:chartSpace',
    { 'xmlns:c': NS_C, 'xmlns:a': NS_A, 'xmlns:r': NS_R },
    `${val('c:date1904', wb.date1904 ? 1 : 0)}${val('c:roundedCorners', 0)}${x(
      'c:chart',
      {},
      `${title}${val('c:autoTitleDeleted', title ? 0 : 1)}${x('c:plotArea', {}, `<c:layout/>${plot}`)}${legend}${val('c:plotVisOnly', 1)}${val('c:dispBlanksAs', 'gap')}`,
    )}${x('c:spPr', {}, `${areaFill}${areaLine}`)}${textPr(ink, 1000)}${ext}`,
  )}`;
}

// -------------------------------------------------------------------------------------------------
// Writing chartex parts

function guid(): string {
  const h = () => Math.floor(Math.random() * 0x10000).toString(16).toUpperCase().padStart(4, '0');
  return `{${h()}${h()}-${h()}-4${h().slice(1)}-${h()}-${h()}${h()}${h()}}`;
}

function cxStr(wb: Workbook, ref: string, type = 'cat'): string {
  const { sheet: name, ref: local } = splitSheetRef(ref);
  const range = parseRange(local.replace(/\$/g, '').toUpperCase());
  const lvl = (texts: string[]) => x('cx:lvl', { ptCount: texts.length }, texts.map((t, i) => (t === '' ? '' : x('cx:pt', { idx: i }, escapeText(t)))).join(''));
  if (range && range.c1 > range.c0 && name !== undefined) {
    // Hierarchical labels: one level per column, the innermost (rightmost) first.
    const columns: string[][] = [];
    for (let c = range.c1; c >= range.c0; c--) columns.push(readRefData(wb, `'${name.replace(/'/g, "''")}'!${refOfColumn(range, c)}`).texts);
    return x('cx:strDim', { type }, `<cx:f>${escapeText(ref)}</cx:f>${columns.map(lvl).join('')}`);
  }
  const d = readRefData(wb, ref);
  return x('cx:strDim', { type }, `<cx:f>${escapeText(ref)}</cx:f>${lvl(d.texts)}`);
}

function refOfColumn(range: { r0: number; r1: number }, c: number): string {
  let name = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return `$${name}$${range.r0 + 1}:$${name}$${range.r1 + 1}`;
}

function cxNum(wb: Workbook, ref: string, type = 'val'): string {
  const d = readRefData(wb, ref);
  return x('cx:numDim', { type }, `<cx:f>${escapeText(ref)}</cx:f>${x('cx:lvl', { ptCount: d.numbers.length, formatCode: d.numFmt || 'General' }, d.numbers.map((n, i) => (n === null ? '' : x('cx:pt', { idx: i }, String(n)))).join(''))}`);
}

export function writeChartexPart(wb: Workbook, chart: SheetChart): string {
  const type = chart.type;
  const series = type === 'box-whisker' ? chart.series : chart.series.slice(0, 1);
  const data = series
    .map((s, i) => x('cx:data', { id: i }, `${chart.categories && type !== 'histogram' ? cxStr(wb, chart.categories) : ''}${cxNum(wb, s.values, type === 'treemap' || type === 'sunburst' ? 'size' : 'val')}`))
    .join('');
  const tiled = type === 'treemap' || type === 'sunburst';
  const labels = chart.valueLabels || type === 'funnel' || tiled ? x('cx:dataLabels', { pos: type === 'treemap' ? 'inEnd' : type === 'funnel' || type === 'sunburst' ? 'ctr' : 'outEnd' }, x('cx:visibility', { seriesName: 0, categoryName: tiled ? 1 : 0, value: tiled && !chart.valueLabels ? 0 : 1 })) : '';
  const layout = (s: ChartSeries, i: number): string => {
    const color = seriesColor(chart, s, i);
    const name = s.nameRef ? `<cx:f>${escapeText(s.nameRef)}</cx:f><cx:v>${escapeText(readRefData(wb, s.nameRef).texts.join(' '))}</cx:v>` : `<cx:v>${escapeText(s.name ?? `Series ${i + 1}`)}</cx:v>`;
    const tx = x('cx:tx', {}, x('cx:txData', {}, name));
    const spPr = tiled ? '' : x('cx:spPr', {}, solid(color));
    let pr = '';
    let axes = '';
    let id: string = type;
    switch (type) {
      case 'histogram':
        id = 'clusteredColumn';
        pr = x('cx:layoutPr', {}, x('cx:binning', { intervalClosed: 'r' }));
        break;
      case 'pareto':
        id = 'clusteredColumn';
        pr = x('cx:layoutPr', {}, x('cx:aggregation'));
        axes = `${x('cx:axisId', { val: 0 })}${x('cx:axisId', { val: 1 })}`;
        break;
      case 'box-whisker':
        id = 'boxWhisker';
        pr = x('cx:layoutPr', {}, `${x('cx:visibility', { meanLine: 0, meanMarker: 1, nonoutliers: 0, outliers: 1 })}${x('cx:statistics', { quartileMethod: 'exclusive' })}`);
        break;
      case 'waterfall': {
        const cats = chart.categories ? readRefData(wb, chart.categories).texts : [];
        const totals = cats.map((c, k) => (k > 0 && /^(total|subtotal|sum|net|balance|grand total)\b/i.test(c.trim()) ? k : -1)).filter((k) => k >= 0);
        pr = x('cx:layoutPr', {}, totals.length ? x('cx:subtotals', {}, totals.map((k) => x('cx:idx', { val: k })).join('')) : '');
        break;
      }
      case 'treemap':
        pr = x('cx:layoutPr', {}, x('cx:parentLabelLayout', { val: 'overlapping' }));
        break;
      default:
        break;
    }
    return x('cx:series', { layoutId: id, uniqueId: guid() }, `${tx}${spPr}${labels}${x('cx:dataId', { val: i })}${pr}${axes}`);
  };
  let plotSeries = series.map(layout).join('');
  if (type === 'pareto') plotSeries += x('cx:series', { layoutId: 'paretoLine', ownerIdx: 0, uniqueId: guid() }, x('cx:axisId', { val: 2 }));
  const grid = chart.gridlines ? x('cx:majorGridlines') : '';
  let axes = '';
  switch (type) {
    case 'histogram':
      axes = `${x('cx:axis', { id: 0 }, `${x('cx:catScaling', { gapWidth: 0 })}${x('cx:tickLabels')}`)}${x('cx:axis', { id: 1 }, `${x('cx:valScaling')}${grid}${x('cx:tickLabels')}`)}`;
      break;
    case 'pareto':
      axes = `${x('cx:axis', { id: 0 }, `${x('cx:catScaling', { gapWidth: 0 })}${x('cx:tickLabels')}`)}${x('cx:axis', { id: 1 }, `${x('cx:valScaling')}${grid}${x('cx:tickLabels')}`)}${x('cx:axis', { id: 2 }, `${x('cx:valScaling', { max: 1, min: 0 })}${x('cx:units', { unit: 'percentage' })}${x('cx:tickLabels')}`)}`;
      break;
    case 'box-whisker':
      axes = `${x('cx:axis', { id: 0 }, `${x('cx:catScaling', { gapWidth: 1 })}${x('cx:tickLabels')}`)}${x('cx:axis', { id: 1 }, `${x('cx:valScaling')}${grid}${x('cx:tickLabels')}`)}`;
      break;
    case 'waterfall':
      axes = `${x('cx:axis', { id: 0 }, `${x('cx:catScaling', { gapWidth: 0.5 })}${x('cx:tickLabels')}`)}${x('cx:axis', { id: 1 }, `${x('cx:valScaling')}${grid}${x('cx:tickLabels')}`)}`;
      break;
    case 'funnel':
      axes = x('cx:axis', { id: 0 }, `${x('cx:catScaling', { gapWidth: 0.06 })}${x('cx:tickLabels')}`);
      break;
    default:
      break;
  }
  const title = chart.title ? x('cx:title', { pos: 't', align: 'ctr', overlay: 0 }, x('cx:tx', {}, x('cx:txData', {}, `<cx:v>${escapeText(chart.title)}</cx:v>`))) : '';
  const legend = chart.showLegend ? x('cx:legend', { pos: LEGEND_POS[chart.legendPosition ?? 'bottom'], align: 'ctr', overlay: 0 }) : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${x(
    'cx:chartSpace',
    { 'xmlns:a': NS_A, 'xmlns:r': NS_R, 'xmlns:cx': NS_CX },
    `${x('cx:chartData', {}, data)}${x('cx:chart', {}, `${title}${x('cx:plotArea', {}, `${x('cx:plotAreaRegion', {}, plotSeries)}${axes}`)}${legend}`)}`,
  )}`;
}

// -------------------------------------------------------------------------------------------------
// Drawing anchors

function pointXml(tag: string, p: AnchorPoint): XmlElement {
  return el(tag, {}, [
    el('xdr:col', {}, [txt(String(p.c))]),
    el('xdr:colOff', {}, [txt(String(Math.round(p.dx)))]),
    el('xdr:row', {}, [txt(String(p.r))]),
    el('xdr:rowOff', {}, [txt(String(Math.round(p.dy)))]),
  ]);
}

/** An anchor of the file moved to new cells (its content kept). */
export function repositionAnchor(anchor: XmlElement, from: AnchorPoint, to: AnchorPoint): XmlElement {
  const content = elements(anchor).filter((e) => !['from', 'to', 'ext', 'pos', 'clientData'].includes(local(e.name)));
  const clientData = elements(anchor).find((e) => local(e.name) === 'clientData') ?? el('xdr:clientData');
  const editAs = local(anchor.name) === 'twoCellAnchor' ? attr(anchor, 'editAs') : local(anchor.name) === 'oneCellAnchor' ? 'oneCell' : 'absolute';
  return el('xdr:twoCellAnchor', { editAs }, [pointXml('xdr:from', from), pointXml('xdr:to', to), ...content, clientData]);
}

function xfrmEl(sheet: Sheet, from: AnchorPoint, to: AnchorPoint, extra: Record<string, string | number | undefined> = {}, name = 'a:xfrm'): XmlElement {
  const r = SheetLayout.of(sheet).anchorRect(from, to);
  return el(name, extra, [el('a:off', { x: Math.round(r.x * EMU_PER_PX), y: Math.round(r.y * EMU_PER_PX) }), el('a:ext', { cx: Math.round(r.w * EMU_PER_PX), cy: Math.round(r.h * EMU_PER_PX) })]);
}

/** Drawing anchor of a chart written by the editor (chartex charts carry a fallback for old Excel). */
export function chartAnchor(sheet: Sheet, chart: SheetChart, rid: string, id: number, chartex: boolean): XmlElement {
  const nv = el('xdr:nvGraphicFramePr', {}, [el('xdr:cNvPr', { id, name: chart.name || `Chart ${id}`, descr: chart.descr || undefined, title: chart.altTitle || undefined, hidden: chart.hidden ? 1 : undefined }), el('xdr:cNvGraphicFramePr')]);
  const xfrm = el('xdr:xfrm', {}, [el('a:off', { x: 0, y: 0 }), el('a:ext', { cx: 0, cy: 0 })]);
  if (!chartex) {
    const frame = el('xdr:graphicFrame', { macro: '' }, [
      nv,
      xfrm,
      el('a:graphic', {}, [el('a:graphicData', { uri: NS_C }, [el('c:chart', { 'xmlns:c': NS_C, 'xmlns:r': NS_R, 'r:id': rid })])]),
    ]);
    return el('xdr:twoCellAnchor', {}, [pointXml('xdr:from', chart.from), pointXml('xdr:to', chart.to), frame, el('xdr:clientData')]);
  }
  const requires = chart.type === 'funnel' ? 'cx2' : 'cx1';
  const requiresNs = chart.type === 'funnel' ? 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex' : 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex';
  const frame = el('xdr:graphicFrame', { macro: '' }, [
    nv,
    xfrm,
    el('a:graphic', {}, [el('a:graphicData', { uri: NS_CX }, [el('cx:chart', { 'xmlns:cx': NS_CX, 'xmlns:r': NS_R, 'r:id': rid })])]),
  ]);
  const fallback = el('xdr:sp', { macro: '', textlink: '' }, [
    el('xdr:nvSpPr', {}, [el('xdr:cNvPr', { id, name: chart.name || `Chart ${id}` }), el('xdr:cNvSpPr', {}, [el('a:spLocks', { noTextEdit: 1 })])]),
    el('xdr:spPr', {}, [
      xfrmEl(sheet, chart.from, chart.to),
      el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')]),
      el('a:solidFill', {}, [el('a:prstClr', { val: 'white' })]),
      el('a:ln', { w: 1 }, [el('a:solidFill', {}, [el('a:prstClr', { val: 'green' })])]),
    ]),
    el('xdr:txBody', {}, [
      el('a:bodyPr', { vertOverflow: 'clip', horzOverflow: 'clip' }),
      el('a:lstStyle'),
      el('a:p', {}, [el('a:r', {}, [el('a:rPr', { lang: 'en-US', sz: 1100 }), el('a:t', {}, [txt('This chart isn’t available in your version of Excel.')])])]),
    ]),
  ]);
  const alt = el('mc:AlternateContent', { 'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006' }, [
    el('mc:Choice', { [`xmlns:${requires}`]: requiresNs, Requires: requires }, [frame]),
    el('mc:Fallback', {}, [fallback]),
  ]);
  return el('xdr:twoCellAnchor', {}, [pointXml('xdr:from', chart.from), pointXml('xdr:to', chart.to), alt, el('xdr:clientData')]);
}

const LINE_PRESETS = new Set(['line', 'straightConnector1', 'bentConnector3', 'curvedConnector3']);

/** Drawing anchor of a shape (or a connector for line presets). */
export function shapeAnchor(sheet: Sheet, shape: SheetShape, id: number): XmlElement {
  const isLine = LINE_PRESETS.has(shape.preset);
  const xfrmAttrs: Record<string, string | number | undefined> = {
    rot: shape.rotation ? Math.round(shape.rotation * 60000) : undefined,
    flipH: shape.flipH ? 1 : undefined,
    flipV: shape.flipV ? 1 : undefined,
  };
  const hex = (c: string) => c.replace('#', '').toUpperCase();
  const fill = shape.fill && !isLine ? el('a:solidFill', {}, [el('a:srgbClr', { val: hex(shape.fill) })]) : el('a:noFill');
  const ln = shape.line
    ? el('a:ln', { w: Math.round(shape.line.width * EMU_PER_PX) }, [
        el('a:solidFill', {}, [el('a:srgbClr', { val: hex(shape.line.color) })]),
        shape.line.dash && shape.line.dash !== 'solid' ? el('a:prstDash', { val: PRST_DASH[shape.line.dash] }) : null,
        shape.arrowStart ? el('a:headEnd', { type: 'triangle' }) : null,
        shape.arrowEnd ? el('a:tailEnd', { type: 'triangle' }) : null,
      ])
    : el('a:ln', {}, [el('a:noFill')]);
  const spPr = el('xdr:spPr', {}, [xfrmEl(sheet, shape.from, shape.to, xfrmAttrs), el('a:prstGeom', { prst: shape.preset }, [el('a:avLst')]), fill, ln]);
  let content: XmlElement;
  if (isLine) {
    content = el('xdr:cxnSp', { macro: '' }, [el('xdr:nvCxnSpPr', {}, [el('xdr:cNvPr', { id, name: shape.name || `Connector ${id}`, descr: shape.descr || undefined, title: shape.altTitle || undefined, hidden: shape.hidden ? 1 : undefined }), el('xdr:cNvCxnSpPr')]), spPr]);
  } else {
    const size = Math.round((shape.fontSize ?? 11) * 100);
    const rPr = () => el('a:rPr', { lang: 'en-US', sz: size, b: shape.bold ? 1 : undefined, i: shape.italic ? 1 : undefined }, [el('a:solidFill', {}, [el('a:srgbClr', { val: hex(shape.fontColor || '#000000') })])]);
    const paras: XmlNode[] = (shape.text ?? '').split('\n').map((line) =>
      el('a:p', {}, [el('a:pPr', { algn: shape.textBox ? undefined : 'ctr' }), line ? el('a:r', {}, [rPr(), el('a:t', {}, [txt(line)])]) : el('a:endParaRPr', { lang: 'en-US', sz: size })]),
    );
    content = el('xdr:sp', { macro: '', textlink: '' }, [
      el('xdr:nvSpPr', {}, [el('xdr:cNvPr', { id, name: shape.name || `${shape.textBox ? 'TextBox' : 'Shape'} ${id}`, descr: shape.descr || undefined, title: shape.altTitle || undefined, hidden: shape.hidden ? 1 : undefined }), el('xdr:cNvSpPr', { txBox: shape.textBox ? 1 : undefined })]),
      spPr,
      el('xdr:txBody', {}, [el('a:bodyPr', { vertOverflow: 'clip', horzOverflow: 'clip', wrap: 'square', rtlCol: 0, anchor: shape.textBox ? 't' : 'ctr' }), el('a:lstStyle'), ...paras]),
    ]);
  }
  return el('xdr:twoCellAnchor', {}, [pointXml('xdr:from', shape.from), pointXml('xdr:to', shape.to), content, el('xdr:clientData')]);
}

/** Anchor points of a drawing anchor (one-cell and absolute anchors get a computed end). */
export function anchorPoints(sheet: Sheet, anchor: XmlElement): { from: AnchorPoint; to: AnchorPoint } {
  const point = (e: XmlElement | undefined): AnchorPoint => ({
    c: Number(textContent(kid(e, 'col') ?? { type: 'text', text: '0' })) || 0,
    r: Number(textContent(kid(e, 'row') ?? { type: 'text', text: '0' })) || 0,
    dx: Number(textContent(kid(e, 'colOff') ?? { type: 'text', text: '0' })) || 0,
    dy: Number(textContent(kid(e, 'rowOff') ?? { type: 'text', text: '0' })) || 0,
  });
  const layout = SheetLayout.of(sheet);
  const kind = local(anchor.name);
  if (kind === 'twoCellAnchor') return { from: point(kid(anchor, 'from')), to: point(kid(anchor, 'to')) };
  const ext = kid(anchor, 'ext');
  const w = (numAttr(ext, 'cx') ?? 0) / EMU_PER_PX;
  const h = (numAttr(ext, 'cy') ?? 0) / EMU_PER_PX;
  if (kind === 'absoluteAnchor') {
    const pos = kid(anchor, 'pos');
    const x0 = (numAttr(pos, 'x') ?? 0) / EMU_PER_PX;
    const y0 = (numAttr(pos, 'y') ?? 0) / EMU_PER_PX;
    return { from: layout.anchorAt(x0, y0), to: layout.anchorAt(x0 + w, y0 + h) };
  }
  const from = point(kid(anchor, 'from'));
  const x0 = layout.colX(from.c) + from.dx / EMU_PER_PX;
  const y0 = layout.rowY(from.r) + from.dy / EMU_PER_PX;
  return { from, to: layout.anchorAt(x0 + w, y0 + h) };
}

