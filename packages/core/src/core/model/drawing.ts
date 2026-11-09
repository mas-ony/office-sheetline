/**
 * Floating objects that the editor creates and edits: charts and shapes. Pictures live in
 * `Sheet.images`. All three share a z-order (`z`) and a two-cell anchor.
 */
import type { AnchorPoint } from './sheet';
import type { XmlElement } from '../xml/xml';

export const CHART_TYPES = [
  'line', 'stacked-line', 'percent-line', 'area', 'stacked-area', 'percent-area', 'bar', 'stacked-bar', 'percent-bar', 'column',
  'stacked-column', 'percent-column', 'pie', 'doughnut', 'scatter', 'bubble', 'radar', 'filled-radar', 'histogram', 'pareto',
  'box-whisker', 'waterfall', 'funnel', 'treemap', 'sunburst', 'stock', 'combo',
] as const;

export type ChartType = (typeof CHART_TYPES)[number];

/** Chart types stored as Office 2016 "chartex" parts. */
export const CHARTEX_TYPES = new Set<ChartType>(['histogram', 'pareto', 'box-whisker', 'waterfall', 'funnel', 'treemap', 'sunburst']);

/** Lines and markers of scatter series. */
export type ScatterStyle = 'markers' | 'lines-markers' | 'lines' | 'smooth-markers' | 'smooth';
/** Stock chart layouts: (volume,) (open,) high, low, close series in that order. */
export type StockKind = 'hlc' | 'ohlc' | 'vhlc' | 'vohlc';
/** How a series of a combo chart is drawn. */
export type SeriesKind = 'column' | 'stacked-column' | 'line' | 'area' | 'stacked-area';
export type ErrorBars = 'none' | 'stderr' | 'percent' | 'stddev';

export type LineDash = 'solid' | 'dashed' | 'dotted' | 'dash-dot';
export type Trendline = 'none' | 'linear' | 'exponential' | 'logarithmic' | 'polynomial' | 'power' | 'moving-average';
export type Side = 'top' | 'bottom' | 'left' | 'right';

export interface ChartSeries {
  /** Literal series name (used when there is no name reference). */
  name?: string;
  /** Reference of the cell holding the name, e.g. Sheet1!$B$1. */
  nameRef?: string;
  /** Reference of the values, e.g. Sheet1!$B$2:$B$5. */
  values: string;
  /** X values of scatter and bubble series. */
  xValues?: string;
  /** Bubble sizes. */
  sizes?: string;
  /** '#RRGGBB'; undefined uses the palette. */
  color?: string;
  /** Line width in pixels (line, radar and scatter charts). */
  lineWidth?: number;
  lineStyle?: LineDash;
  trendline?: Trendline;
  markers?: boolean;
  /** Smoothed line (line and scatter series). */
  smooth?: boolean;
  secondaryAxis?: boolean;
  /** Combo charts: how the series is drawn. */
  kind?: SeriesKind;
}

export interface SheetChart {
  id: string;
  name: string;
  type: ChartType;
  from: AnchorPoint;
  to: AnchorPoint;
  /** Data range the chart was built from (absolute, sheet-qualified). */
  range: string;
  /** 'rows': every row of the range is a series; 'columns': every column is. */
  seriesIn: 'rows' | 'columns';
  /** The first row/column of the range holds names and labels. */
  headers: boolean;
  /** Category labels (absolute, sheet-qualified reference). */
  categories?: string;
  series: ChartSeries[];
  title?: string;
  titlePosition?: Side;
  subtitle?: string;
  legendTitle?: string;
  legendPosition?: Side;
  showLegend: boolean;
  gridlines: boolean;
  valueLabels: boolean;
  xAxisTitle?: string;
  yAxisTitle?: string;
  secondaryAxisTitle?: string;
  yMin?: number;
  yMax?: number;
  /** Series and point colours (Change Colors; theme accents of charts read from a file). */
  palette?: string[];
  /** Markers on every line and radar series ("Line with Markers"). */
  markers?: boolean;
  scatterStyle?: ScatterStyle;
  stockKind?: StockKind;
  errorBars?: ErrorBars;
  /** Vertical (category) gridlines. */
  vGridlines?: boolean;
  hideXAxis?: boolean;
  hideYAxis?: boolean;
  /** Chart Styles gallery entry (1 is the default look). */
  style?: number;
  /** Chart area fill: undefined is white, null no fill. */
  fill?: string | null;
  /** Chart area border: undefined is the default thin grey line, null none. */
  border?: ShapeLine | null;
  /** Alternative text. */
  altTitle?: string;
  descr?: string;
  /** Hidden in the Selection Pane. */
  hidden?: boolean;
  z: number;
  /**
   * The chart came from the file: its drawing anchor and chart part. While the chart settings are
   * unchanged the original part is written back (keeping formatting the editor does not model).
   */
  source?: { anchor: XmlElement; chartPath: string; chartex?: boolean };
  /** Settings changed since import: the chart part is regenerated on save. */
  dirty?: boolean;
}

export interface ShapeLine {
  color: string;
  /** Pixels. */
  width: number;
  dash?: LineDash;
}

export interface SheetShape {
  id: string;
  name: string;
  /** DrawingML preset geometry (rect, roundRect, ellipse, rightArrow...). */
  preset: string;
  from: AnchorPoint;
  to: AnchorPoint;
  /** '#RRGGBB', or null for no fill. */
  fill: string | null;
  line: ShapeLine | null;
  fontColor: string;
  text?: string;
  /** Arrowheads of lines and connectors. */
  arrowStart?: boolean;
  arrowEnd?: boolean;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** Text box: no fill or line by default, text aligned left and top. */
  textBox?: boolean;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  altTitle?: string;
  descr?: string;
  /** Hidden in the Selection Pane. */
  hidden?: boolean;
  z: number;
  /** Original anchor from the file, written back while the shape is unchanged. */
  source?: { anchor: XmlElement };
  dirty?: boolean;
}

export const DEFAULT_SHAPE_FILL = '#4F81BD';
export const DEFAULT_SHAPE_LINE = '#385D8A';

/** Office 2013 theme accents: the series colours of new charts. */
export const CHART_PALETTE = ['#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000', '#4472C4', '#70AD47', '#255E91', '#9E480E', '#636363', '#997300', '#264478', '#43682B'];

let drawingCounter = 1;

export function newDrawingId(prefix: string): string {
  return `${prefix}${Date.now().toString(36)}${(drawingCounter++).toString(36)}`;
}

export function cloneChart(c: SheetChart): SheetChart {
  return { ...c, from: { ...c.from }, to: { ...c.to }, series: c.series.map((s) => ({ ...s })), palette: c.palette?.slice(), border: c.border ? { ...c.border } : c.border };
}

export function cloneShape(s: SheetShape): SheetShape {
  return { ...s, from: { ...s.from }, to: { ...s.to }, line: s.line ? { ...s.line } : null };
}
