/**
 * Contextual tabs: Table Design (active cell in a table), Sparkline (active cell holds one), Chart
 * Design, Shape Format and Picture Format (object selected).
 */
import type { MenuItem } from '../menuItems';
import { TINT, type Command, type Control, type Ctx, type Tab } from './types';
import { P } from './pageLayoutTab';
import { resolveColor, type ColorSpec } from '../../core/model/color';
import type { SheetChart, ErrorBars, Side, Trendline } from '../../core/model/drawing';
import type { SparklineGroup, SparklineColors } from '../../core/model/sparkline';
import { defaultSparklineColors } from '../../core/model/sparkline';
import { guessLayout, resolveRef, seriesForLayout } from '../../core/chartData';
import type { GalleryItem as ShapeGalleryItem } from '../../grid/shapes';
import { SHAPE_STYLES, selected } from '../objects';
import { CHART_PALETTE } from '../../core/model/drawing';

// -----------------------------------------------------------------------------------------------
// Table Design

const table = (ctx: Ctx) => (ctx.state.tableId !== null ? ctx.doc.tableById(ctx.state.tableId) : undefined);

function tableFlag(id: string, label: string, get: (t: NonNullable<ReturnType<typeof table>>['table']) => boolean, key: 'headerRow' | 'totalsRow' | 'autoFilter' | 'showRowStripes' | 'showColumnStripes' | 'showFirstColumn' | 'showLastColumn' | 'autoExpand'): Command {
  return {
    id,
    label,
    checked: (ctx) => {
      const t = table(ctx)?.table;
      return t ? get(t) : false;
    },
    run: (ctx) => {
      const t = table(ctx)?.table;
      if (t) ctx.doc.setTableOptions(t.id, { [key]: !get(t) });
    },
  };
}

/** The table at the selection (Table Design). */
export function currentTable(ctx: Ctx) {
  return table(ctx)?.table;
}

/** Renames a table from the Table Name field; false (with a message) when the name is refused. */
export function renameTableTo(ctx: Ctx, id: number, name: string): boolean {
  const error = ctx.doc.renameTable(id, name.trim());
  if (error) ctx.notify(error);
  return !error;
}

export const T = {
  resize: { id: 'resize-table', label: 'Resize Table', icon: 'aspect_ratio', run: (ctx) => ctx.state.tableId !== null && ctx.openDialog({ kind: 'resize-table', tableId: ctx.state.tableId }) } as Command,
  convert: {
    id: 'convert-to-range',
    label: 'Convert to Range',
    icon: 'table_rows',
    run: async (ctx) => {
      const t = table(ctx)?.table;
      if (t && (await ctx.confirm('Do you want to convert the table to a normal range?', { title: 'Convert to Range', ok: 'Yes', cancel: 'No' }))) ctx.doc.convertTableToRange(t.id);
    },
  } as Command,
  removeDuplicates: {
    id: 'table-remove-duplicates',
    label: 'Remove Duplicates',
    icon: 'remove_duplicates',
    run: (ctx) => {
      const t = table(ctx)?.table;
      if (!t) return;
      ctx.doc.select(t.ref);
      ctx.openDialog({ kind: 'remove-duplicates' });
    },
  } as Command,
  autoExpand: tableFlag('table-auto-expand', 'Allow Auto Expand', (t) => t.autoExpand !== false, 'autoExpand'),
  header: tableFlag('table-header-row', 'Header Row', (t) => t.headerRow, 'headerRow'),
  totals: tableFlag('table-total-row', 'Total Row', (t) => t.totalsRow, 'totalsRow'),
  banded: tableFlag('table-banded-rows', 'Banded Rows', (t) => t.style.showRowStripes, 'showRowStripes'),
  first: tableFlag('table-first-column', 'First Column', (t) => t.style.showFirstColumn, 'showFirstColumn'),
  last: tableFlag('table-last-column', 'Last Column', (t) => t.style.showLastColumn, 'showLastColumn'),
  bandedCols: tableFlag('table-banded-columns', 'Banded Columns', (t) => t.style.showColumnStripes, 'showColumnStripes'),
  filterButton: tableFlag('table-filter-button', 'Filter Button', (t) => t.autoFilter, 'autoFilter'),
  styles: {
    id: 'table-styles',
    label: 'Quick Styles',
    tip: 'Table Styles',
    icon: 'format_table',
    tint: TINT.blue,
    panel: {
      kind: 'table-styles',
      showNone: true,
      pick: (ctx, name) => {
        const t = table(ctx)?.table;
        if (t) ctx.doc.setTableStyle(t.id, name);
      },
    },
  } as Command,
};

export const tableDesignTab: Tab = {
  id: 'table-design',
  label: 'Table Design',
  contextual: true,
  visible: (ctx) => ctx.state.tableId !== null,
  groups: () => [
    {
      id: 'table-properties',
      label: 'Properties',
      icon: 'table_chart',
      controls: [{ type: 'stack', items: [{ type: 'view', id: 'table-name', view: { kind: 'table-name' } }, { type: 'small', cmd: T.resize }] }],
    },
    {
      id: 'table-tools',
      label: 'Tools',
      icon: 'settings',
      controls: [{ type: 'stack', items: [{ type: 'small', cmd: T.convert }, { type: 'small', cmd: T.removeDuplicates }, { type: 'check', cmd: T.autoExpand }] }],
    },
    {
      id: 'table-style-options',
      label: 'Table Style Options',
      icon: 'tune',
      controls: [
        { type: 'stack', items: [{ type: 'check', cmd: T.header }, { type: 'check', cmd: T.totals }, { type: 'check', cmd: T.banded }] },
        { type: 'stack', items: [{ type: 'check', cmd: T.first }, { type: 'check', cmd: T.last }, { type: 'check', cmd: T.bandedCols }] },
        { type: 'stack', items: [{ type: 'check', cmd: T.filterButton }] },
      ],
    },
    { id: 'table-styles', label: 'Table Styles', icon: 'format_table', controls: [{ type: 'large', cmd: T.styles }] },
  ],
};

// -----------------------------------------------------------------------------------------------
// Chart Design

const chartOf = (ctx: Ctx): SheetChart | null => {
  const s = selected(ctx);
  return s?.kind === 'chart' ? s.obj : null;
};

function updateChart(ctx: Ctx, patch: Partial<SheetChart> | ((c: SheetChart) => Partial<SheetChart>)): void {
  const c = chartOf(ctx);
  if (!c) return;
  ctx.doc.updateChart(c.id, typeof patch === 'function' ? patch(c) : patch);
}

const LEGENDS: { side: Side | null; label: string }[] = [
  { side: null, label: 'None' },
  { side: 'right', label: 'Right' },
  { side: 'top', label: 'Top' },
  { side: 'left', label: 'Left' },
  { side: 'bottom', label: 'Bottom' },
];

const ERROR_BARS: { kind: ErrorBars; label: string }[] = [
  { kind: 'none', label: 'None' },
  { kind: 'stderr', label: 'Standard Error' },
  { kind: 'percent', label: 'Percentage' },
  { kind: 'stddev', label: 'Standard Deviation' },
];

const TRENDS: { kind: Trendline; label: string }[] = [
  { kind: 'none', label: 'None' },
  { kind: 'linear', label: 'Linear' },
  { kind: 'exponential', label: 'Exponential' },
  { kind: 'logarithmic', label: 'Logarithmic' },
  { kind: 'polynomial', label: 'Polynomial' },
  { kind: 'power', label: 'Power' },
  { kind: 'moving-average', label: 'Moving Average' },
];

function elementMenu(ctx: Ctx): MenuItem[] {
  const c = chartOf(ctx);
  if (!c) return [];
  const pie = c.type === 'pie' || c.type === 'doughnut' || c.type === 'treemap' || c.type === 'sunburst' || c.type === 'funnel';
  return [
    {
      title: 'Axes',
      icon: 'linear_scale',
      disabled: pie,
      submenu: [
        { title: 'Primary Horizontal', checked: !c.hideXAxis, onClick: () => updateChart(ctx, { hideXAxis: !c.hideXAxis || undefined }) },
        { title: 'Primary Vertical', checked: !c.hideYAxis, onClick: () => updateChart(ctx, { hideYAxis: !c.hideYAxis || undefined }) },
      ],
    },
    {
      title: 'Axis Titles',
      icon: 'title',
      disabled: pie,
      submenu: [
        { title: 'Primary Horizontal', checked: !!c.xAxisTitle, onClick: () => updateChart(ctx, { xAxisTitle: c.xAxisTitle ? undefined : 'Axis Title' }) },
        { title: 'Primary Vertical', checked: !!c.yAxisTitle, onClick: () => updateChart(ctx, { yAxisTitle: c.yAxisTitle ? undefined : 'Axis Title' }) },
      ],
    },
    {
      title: 'Chart Title',
      icon: 'title',
      submenu: [
        { title: 'None', checked: !c.title, onClick: () => updateChart(ctx, { title: undefined }) },
        { title: 'Above Chart', checked: !!c.title && (c.titlePosition ?? 'top') === 'top', onClick: () => updateChart(ctx, { title: c.title || 'Chart Title', titlePosition: 'top' }) },
        { title: 'Below Chart', checked: !!c.title && c.titlePosition === 'bottom', onClick: () => updateChart(ctx, { title: c.title || 'Chart Title', titlePosition: 'bottom' }) },
      ],
    },
    {
      title: 'Data Labels',
      icon: 'numbers',
      submenu: [
        { title: 'None', checked: !c.valueLabels, onClick: () => updateChart(ctx, { valueLabels: false }) },
        { title: 'Show', checked: c.valueLabels, onClick: () => updateChart(ctx, { valueLabels: true }) },
      ],
    },
    {
      title: 'Error Bars',
      icon: 'vertical_align_center',
      disabled: pie,
      submenu: ERROR_BARS.map((e) => ({ title: e.label, checked: (c.errorBars ?? 'none') === e.kind, onClick: () => updateChart(ctx, { errorBars: e.kind === 'none' ? undefined : e.kind }) })),
    },
    {
      title: 'Gridlines',
      icon: 'grid_on',
      disabled: pie,
      submenu: [
        { title: 'Primary Major Horizontal', checked: c.gridlines, onClick: () => updateChart(ctx, { gridlines: !c.gridlines }) },
        { title: 'Primary Major Vertical', checked: !!c.vGridlines, onClick: () => updateChart(ctx, { vGridlines: !c.vGridlines || undefined }) },
      ],
    },
    {
      title: 'Legend',
      icon: 'list',
      submenu: LEGENDS.map((l) => ({
        title: l.label,
        checked: l.side ? c.showLegend && (c.legendPosition ?? 'bottom') === l.side : !c.showLegend,
        onClick: () => updateChart(ctx, l.side ? { showLegend: true, legendPosition: l.side } : { showLegend: false }),
      })),
    },
    {
      title: 'Trendline',
      icon: 'trending_up',
      disabled: pie,
      submenu: TRENDS.map((t) => ({
        title: t.label,
        checked: c.series.every((s) => (s.trendline ?? 'none') === t.kind),
        onClick: () => updateChart(ctx, (ch) => ({ series: ch.series.map((s) => ({ ...s, trendline: t.kind === 'none' ? undefined : t.kind })) })),
      })),
    },
  ];
}

const QUICK_LAYOUTS: { label: string; patch: Partial<SheetChart> }[] = [
  { label: 'Layout 1', patch: { title: 'Chart Title', showLegend: true, legendPosition: 'right', xAxisTitle: undefined, yAxisTitle: undefined, valueLabels: false, gridlines: true } },
  { label: 'Layout 2', patch: { title: 'Chart Title', showLegend: true, legendPosition: 'top', valueLabels: true, gridlines: false, hideYAxis: true, xAxisTitle: undefined, yAxisTitle: undefined } },
  { label: 'Layout 3', patch: { title: 'Chart Title', showLegend: true, legendPosition: 'bottom', valueLabels: false, gridlines: true, hideYAxis: undefined } },
  { label: 'Layout 4', patch: { title: undefined, showLegend: true, legendPosition: 'bottom', valueLabels: true, gridlines: false } },
  { label: 'Layout 5', patch: { title: 'Chart Title', showLegend: false, yAxisTitle: 'Axis Title', gridlines: true } },
  { label: 'Layout 6', patch: { title: 'Chart Title', showLegend: true, legendPosition: 'right', yAxisTitle: 'Axis Title', gridlines: true } },
  { label: 'Layout 7', patch: { title: undefined, showLegend: true, legendPosition: 'right', xAxisTitle: 'Axis Title', yAxisTitle: 'Axis Title', vGridlines: true } },
  { label: 'Layout 8', patch: { title: 'Chart Title', showLegend: false, xAxisTitle: 'Axis Title', yAxisTitle: 'Axis Title', gridlines: false } },
  { label: 'Layout 9', patch: { title: 'Chart Title', showLegend: true, legendPosition: 'right', xAxisTitle: 'Axis Title', yAxisTitle: 'Axis Title', gridlines: true } },
];

export const CHART_STYLES: { label: string; patch: Partial<SheetChart> }[] = [
  { label: 'Style 1', patch: { style: 1, fill: undefined, border: undefined, gridlines: true, valueLabels: false, hideYAxis: undefined } },
  { label: 'Style 2', patch: { style: 2, fill: undefined, border: undefined, gridlines: false, valueLabels: true, hideYAxis: true } },
  { label: 'Style 3', patch: { style: 3, fill: '#F2F2F2', border: null, gridlines: true, valueLabels: false, hideYAxis: undefined } },
  { label: 'Style 4', patch: { style: 4, fill: undefined, border: null, gridlines: false, valueLabels: false, hideYAxis: undefined } },
  { label: 'Style 5', patch: { style: 5, fill: '#D9D9D9', border: undefined, gridlines: true, valueLabels: false, hideYAxis: undefined } },
  { label: 'Style 6', patch: { style: 6, fill: '#FFFFFF', border: { color: '#595959', width: 2 }, gridlines: true, valueLabels: false, hideYAxis: undefined } },
  { label: 'Style 7', patch: { style: 7, fill: '#404040', border: null, gridlines: true, valueLabels: false, hideYAxis: undefined } },
  { label: 'Style 8', patch: { style: 8, fill: '#262626', border: null, gridlines: false, valueLabels: true, hideYAxis: true } },
];

function mix(hex: string, t: number): string {
  const n = parseInt(hex.replace('#', ''), 16);
  const ch = (v: number) => Math.round(t >= 0 ? v + (255 - v) * t : v * (1 + t)).toString(16).padStart(2, '0');
  return `#${ch((n >> 16) & 255)}${ch((n >> 8) & 255)}${ch(n & 255)}`.toUpperCase();
}

/** Change Colors: colourful palettes from the theme accents, monochromatic ones from one accent. */
export function colorPalettes(ctx: Ctx): { label: string; colors: string[] }[] {
  const t = ctx.doc.wb.theme.colors;
  const a = [t.accent1, t.accent2, t.accent3, t.accent4, t.accent5, t.accent6].map((x) => `#${x}`);
  const out = [
    { label: 'Colorful Palette 1', colors: a },
    { label: 'Colorful Palette 2', colors: [a[1], a[3], a[5], mix(a[1], -0.4), mix(a[3], -0.4), mix(a[5], -0.4)] },
    { label: 'Colorful Palette 3', colors: [a[2], a[4], a[0], mix(a[2], -0.4), mix(a[4], -0.4), mix(a[0], -0.4)] },
    { label: 'Colorful Palette 4', colors: [a[5], a[4], a[3], a[2], a[1], a[0]] },
  ];
  a.forEach((c, i) => out.push({ label: `Monochromatic Palette ${i + 1}`, colors: [c, mix(c, -0.25), mix(c, 0.25), mix(c, -0.5), mix(c, 0.5), mix(c, -0.7)] }));
  out.push({ label: 'Monochromatic Palette 7', colors: ['#7F7F7F', '#595959', '#A6A6A6', '#404040', '#BFBFBF', '#262626'] });
  return out;
}

function switchRowColumn(ctx: Ctx): void {
  const c = chartOf(ctx);
  if (!c) return;
  const hit = resolveRef(ctx.doc, c.range);
  if (!hit) {
    ctx.notify('This chart\'s data is not a single range: use Select Data to change it.');
    return;
  }
  const seriesIn = c.seriesIn === 'rows' ? 'columns' : 'rows';
  const layout = guessLayout(ctx.doc, hit.sheet, hit.range, c.type, seriesIn);
  const { categories, series } = seriesForLayout(hit.sheet, hit.range, layout, c.type, c.series);
  ctx.doc.updateChart(c.id, { seriesIn, headers: layout.headers, categories, series });
}

const noChart = (ctx: Ctx) => !chartOf(ctx);

export const CD = {
  addElement: { id: 'add-chart-element', label: 'Add Chart Element', icon: 'add_chart', tint: TINT.blue, menu: elementMenu, disabled: noChart } as Command,
  quickLayout: {
    id: 'quick-layout',
    label: 'Quick Layout',
    icon: 'dashboard',
    tint: TINT.blue,
    menu: (ctx) => QUICK_LAYOUTS.map((l) => ({ title: l.label, onClick: () => updateChart(ctx, l.patch) })),
    disabled: noChart,
  } as Command,
  colors: {
    id: 'change-colors',
    label: 'Change Colors',
    icon: 'palette',
    tint: TINT.orange,
    menu: (ctx) =>
      colorPalettes(ctx).map((p) => ({
        title: p.label,
        swatch: { kind: 'palette', colors: p.colors },
        checked: JSON.stringify(chartOf(ctx)?.palette ?? CHART_PALETTE.slice(0, 6)) === JSON.stringify(p.colors),
        onClick: () => updateChart(ctx, (c) => ({ palette: p.colors, series: c.series.map((s) => ({ ...s, color: undefined })) })),
      })),
    disabled: noChart,
  } as Command,
  styles: {
    id: 'chart-styles',
    label: 'Chart Styles',
    icon: 'style',
    tint: TINT.blue,
    menu: (ctx) => CHART_STYLES.map((s) => ({ title: s.label, checked: (chartOf(ctx)?.style ?? 1) === s.patch.style, onClick: () => updateChart(ctx, s.patch) })),
    disabled: noChart,
  } as Command,
  switchRowCol: { id: 'switch-row-column', label: 'Switch Row/Column', icon: 'swap_horiz', run: switchRowColumn, disabled: noChart } as Command,
  selectData: { id: 'select-data', label: 'Select Data', icon: 'dataset', run: (ctx) => ctx.selectedObject && ctx.openDialog({ kind: 'select-data', chartId: ctx.selectedObject }), disabled: noChart } as Command,
  changeType: { id: 'change-chart-type', label: 'Change Chart Type', icon: 'insert_chart', tint: TINT.blue, run: (ctx) => ctx.selectedObject && ctx.openDialog({ kind: 'insert-chart', changeId: ctx.selectedObject }), disabled: noChart } as Command,
  move: { id: 'move-chart', label: 'Move Chart', icon: 'drive_file_move', run: (ctx) => ctx.selectedObject && ctx.openDialog({ kind: 'move-chart', chartId: ctx.selectedObject }), disabled: noChart } as Command,
  altText: { id: 'alt-text', label: 'Alt Text', icon: 'accessibility', run: (ctx) => ctx.selectedObject && ctx.openDialog({ kind: 'alt-text', id: ctx.selectedObject }), disabled: (ctx) => !ctx.selectedObject } as Command,
  formatPane: { id: 'format-pane', label: 'Format Pane', icon: 'view_sidebar', checked: (ctx) => ctx.panel === 'format', run: (ctx) => ctx.setPanel(ctx.panel === 'format' ? null : 'format'), disabled: (ctx) => !ctx.selectedObject } as Command,
  deleteObject: {
    id: 'delete-object',
    label: 'Delete',
    icon: 'delete',
    run: (ctx) => {
      if (!ctx.selectedObject) return;
      ctx.doc.deleteObject(ctx.selectedObject);
      ctx.selectObject(null);
    },
    disabled: (ctx) => !ctx.selectedObject,
  } as Command,
};

export const chartDesignTab: Tab = {
  id: 'chart-design',
  label: 'Chart Design',
  contextual: true,
  visible: (ctx) => !!chartOf(ctx),
  groups: () => [
    { id: 'chart-layouts', label: 'Chart Layouts', icon: 'add_chart', controls: [{ type: 'large', cmd: CD.addElement }, { type: 'large', cmd: CD.quickLayout }] },
    { id: 'chart-styles', label: 'Chart Styles', icon: 'style', controls: [{ type: 'large', cmd: CD.colors }, { type: 'large', cmd: CD.styles }] },
    { id: 'chart-data', label: 'Data', icon: 'dataset', controls: [{ type: 'large', cmd: CD.switchRowCol }, { type: 'large', cmd: CD.selectData }] },
    { id: 'chart-type', label: 'Type', icon: 'insert_chart', controls: [{ type: 'large', cmd: CD.changeType }] },
    { id: 'chart-location', label: 'Location', icon: 'drive_file_move', controls: [{ type: 'large', cmd: CD.move }] },
    { id: 'chart-accessibility', label: 'Accessibility', icon: 'accessibility', controls: [{ type: 'large', cmd: CD.altText }] },
    { id: 'chart-format', label: 'Format', icon: 'view_sidebar', controls: [{ type: 'large', cmd: CD.formatPane }, { type: 'large', cmd: CD.deleteObject }] },
  ],
};

// -----------------------------------------------------------------------------------------------
// Shape Format and Picture Format

const shapeOf = (ctx: Ctx) => {
  const s = selected(ctx);
  return s?.kind === 'shape' ? s.obj : null;
};

/** Change Shape: gives the selected shape another outline. */
export function changeShapeTo(ctx: Ctx, it: ShapeGalleryItem): void {
  const s = shapeOf(ctx);
  if (s) ctx.doc.updateShape(s.id, { preset: it.preset, arrowStart: it.arrowStart || undefined, arrowEnd: it.arrowEnd || undefined });
}

/** Shape Styles: applies a style preset to the selected shape. */
export function applyShapeStyle(ctx: Ctx, st: (typeof SHAPE_STYLES)[number]): void {
  const s = shapeOf(ctx);
  if (s) ctx.doc.updateShape(s.id, { fill: st.fill, line: { color: st.line, width: s.line?.width ?? 1 }, fontColor: st.font });
}

const toHex = (ctx: Ctx, c: ColorSpec | null) => (c ? `#${resolveColor(c, ctx.doc.wb.theme, '000000')}` : null);

const shapeFill: Control = {
  type: 'color',
  cmd: { id: 'shape-fill', label: 'Shape Fill', icon: 'format_color_fill', disabled: (ctx) => !shapeOf(ctx) },
  value: (ctx) => shapeOf(ctx)?.fill ?? null,
  apply: (ctx, c) => {
    const s = shapeOf(ctx);
    if (s) ctx.doc.updateShape(s.id, { fill: toHex(ctx, c) });
  },
  noneLabel: 'No Fill',
  initial: '#5B9BD5',
};

const shapeOutline: Control = {
  type: 'color',
  cmd: { id: 'shape-outline', label: 'Shape Outline', icon: 'border_color', disabled: (ctx) => !shapeOf(ctx) },
  value: (ctx) => shapeOf(ctx)?.line?.color ?? null,
  apply: (ctx, c) => {
    const s = shapeOf(ctx);
    if (!s) return;
    const hex = toHex(ctx, c);
    ctx.doc.updateShape(s.id, { line: hex ? { color: hex, width: s.line?.width ?? 1, dash: s.line?.dash } : null });
  },
  noneLabel: 'No Outline',
  initial: '#41719C',
};

const shapeFont: Control = {
  type: 'color',
  cmd: { id: 'shape-font-color', label: 'Text Fill', icon: 'format_color_text', disabled: (ctx) => !shapeOf(ctx) },
  value: (ctx) => shapeOf(ctx)?.fontColor ?? null,
  apply: (ctx, c) => {
    const s = shapeOf(ctx);
    if (s) ctx.doc.updateShape(s.id, { fontColor: toHex(ctx, c) ?? '#000000' });
  },
  noneLabel: 'Automatic',
  initial: '#000000',
};

const outlineWeight: Command = {
  id: 'shape-outline-weight',
  label: 'Weight',
  icon: 'line_weight',
  disabled: (ctx) => !shapeOf(ctx)?.line,
  menu: (ctx) =>
    [0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((pt) => {
      const px = Math.round(((pt * 4) / 3) * 2) / 2;
      const s = shapeOf(ctx);
      return { title: `${pt} pt`, checked: s?.line?.width === px, onClick: () => s?.line && ctx.doc.updateShape(s.id, { line: { ...s.line, width: px } }) };
    }),
};

const outlineDash: Command = {
  id: 'shape-outline-dash',
  label: 'Dashes',
  icon: 'line_style',
  disabled: (ctx) => !shapeOf(ctx)?.line,
  menu: (ctx) =>
    (['solid', 'dashed', 'dotted', 'dash-dot'] as const).map((d) => {
      const s = shapeOf(ctx);
      return { title: d === 'solid' ? 'Solid' : d === 'dashed' ? 'Dash' : d === 'dotted' ? 'Round Dot' : 'Dash Dot', checked: (s?.line?.dash ?? 'solid') === d, onClick: () => s?.line && ctx.doc.updateShape(s.id, { line: { ...s.line, dash: d === 'solid' ? undefined : d } }) };
    }),
};

export const SF = {
  changeShape: { id: 'change-shape', label: 'Change Shape', icon: 'interests', panel: { kind: 'change-shape' }, disabled: (ctx) => !shapeOf(ctx) } as Command,
  editText: { id: 'edit-shape-text', label: 'Edit Text', icon: 'edit', run: (ctx) => ctx.selectedObject && ctx.openDialog({ kind: 'shape-text', id: ctx.selectedObject }), disabled: (ctx) => !shapeOf(ctx) } as Command,
  styles: { id: 'shape-styles', label: 'Shape Styles', icon: 'style', tint: TINT.blue, panel: { kind: 'shape-styles' }, disabled: (ctx) => !shapeOf(ctx) } as Command,
  bold: { id: 'shape-bold', label: 'Bold', icon: 'format_bold', checked: (ctx) => !!shapeOf(ctx)?.bold, run: (ctx) => { const s = shapeOf(ctx); if (s) ctx.doc.updateShape(s.id, { bold: !s.bold || undefined }); }, disabled: (ctx) => !shapeOf(ctx) } as Command,
  italic: { id: 'shape-italic', label: 'Italic', icon: 'format_italic', checked: (ctx) => !!shapeOf(ctx)?.italic, run: (ctx) => { const s = shapeOf(ctx); if (s) ctx.doc.updateShape(s.id, { italic: !s.italic || undefined }); }, disabled: (ctx) => !shapeOf(ctx) } as Command,
  changePicture: {
    id: 'change-picture',
    label: 'Change Picture',
    icon: 'add_photo_alternate',
    run: (ctx) => {
      const id = ctx.selectedObject;
      const hit = id ? ctx.doc.objectById(id) : null;
      if (!hit || hit.kind !== 'image') return;
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = 'image/*';
      input.onchange = async () => {
        const file = input.files?.[0];
        if (!file) return;
        const data = new Uint8Array(await file.arrayBuffer());
        ctx.doc.replaceImage(hit.obj.id, data, file.type || 'image/png');
      };
      input.click();
    },
  } as Command,
};

const arrangeGroup = {
  id: 'arrange',
  label: 'Arrange',
  icon: 'flip_to_front' as const,
  controls: [
    { type: 'large' as const, cmd: P.forward, split: true },
    { type: 'large' as const, cmd: P.backward, split: true },
    { type: 'large' as const, cmd: P.selectionPane },
    { type: 'stack' as const, items: [{ type: 'small' as const, cmd: P.align }, { type: 'small' as const, cmd: P.rotate }] },
  ],
};

export const shapeFormatTab: Tab = {
  id: 'shape-format',
  label: 'Shape Format',
  contextual: true,
  visible: (ctx) => !!shapeOf(ctx),
  groups: () => [
    { id: 'insert-shapes', label: 'Insert Shapes', icon: 'interests', controls: [{ type: 'large', cmd: SF.changeShape }, { type: 'large', cmd: SF.editText }] },
    {
      id: 'shape-styles',
      label: 'Shape Styles',
      icon: 'style',
      controls: [
        { type: 'large', cmd: SF.styles },
        { type: 'stack', items: [shapeFill, shapeOutline, shapeFont] },
        { type: 'stack', items: [{ type: 'small', cmd: outlineWeight }, { type: 'small', cmd: outlineDash }, { type: 'row', items: [{ type: 'small', cmd: SF.bold, iconOnly: true }, { type: 'small', cmd: SF.italic, iconOnly: true }] }] },
      ],
    },
    { id: 'shape-accessibility', label: 'Accessibility', icon: 'accessibility', controls: [{ type: 'large', cmd: CD.altText }] },
    arrangeGroup,
    { id: 'shape-size', label: 'Size', icon: 'aspect_ratio', controls: [{ type: 'view', id: 'shape-size-fields', view: { kind: 'object-size' } }] },
    { id: 'shape-format', label: 'Format', icon: 'view_sidebar', controls: [{ type: 'large', cmd: CD.formatPane }, { type: 'large', cmd: CD.deleteObject }] },
  ],
};

export const pictureFormatTab: Tab = {
  id: 'picture-format',
  label: 'Picture Format',
  contextual: true,
  visible: (ctx) => selected(ctx)?.kind === 'image',
  groups: () => [
    { id: 'picture-adjust', label: 'Adjust', icon: 'image', controls: [{ type: 'large', cmd: SF.changePicture }] },
    { id: 'picture-accessibility', label: 'Accessibility', icon: 'accessibility', controls: [{ type: 'large', cmd: CD.altText }] },
    arrangeGroup,
    { id: 'picture-size', label: 'Size', icon: 'aspect_ratio', controls: [{ type: 'view', id: 'picture-size-fields', view: { kind: 'object-size' } }, { type: 'large', cmd: CD.deleteObject }] },
  ],
};

// -----------------------------------------------------------------------------------------------
// Sparkline

const groupOf = (ctx: Ctx): SparklineGroup | null => {
  const { r, c } = ctx.doc.selection.active;
  return ctx.doc.sparklineAt(r, c)?.group ?? null;
};

function updateGroup(ctx: Ctx, patch: Partial<SparklineGroup>): void {
  const g = groupOf(ctx);
  if (g) ctx.doc.updateSparklineGroup(g.id, patch);
}

function pointFlag(id: string, label: string, key: 'high' | 'low' | 'negative' | 'first' | 'last' | 'markers'): Command {
  return {
    id,
    label,
    checked: (ctx) => !!groupOf(ctx)?.[key],
    disabled: (ctx) => key === 'markers' && groupOf(ctx)?.type !== 'line',
    run: (ctx) => {
      const g = groupOf(ctx);
      if (g) updateGroup(ctx, { [key]: !g[key] });
    },
  };
}

function sparkColor(id: string, label: string, key: SparklineColors, icon: Command['icon']): Control {
  return {
    type: 'color',
    cmd: { id, label, icon },
    value: (ctx) => {
      const g = groupOf(ctx);
      return g ? `#${resolveColor(g.colors[key], ctx.doc.wb.theme, '000000')}` : null;
    },
    apply: (ctx, c) => {
      const g = groupOf(ctx);
      if (g && c) updateGroup(ctx, { colors: { ...g.colors, [key]: c } });
    },
    noneLabel: 'Automatic',
    initial: '#2E75B6',
  };
}

function axisMenu(ctx: Ctx): MenuItem[] {
  const g = groupOf(ctx);
  if (!g) return [];
  const custom = async (which: 'min' | 'max') => {
    const cur = which === 'min' ? g.manualMin : g.manualMax;
    const v = await ctx.prompt(`Enter the ${which === 'min' ? 'minimum' : 'maximum'} value for the vertical axis:`, String(cur ?? 0), { title: 'Sparkline Vertical Axis Setting' });
    if (v === null) return;
    const n = Number(v);
    if (!Number.isFinite(n)) {
      ctx.notify('Enter a number.');
      return;
    }
    updateGroup(ctx, which === 'min' ? { minAxisType: 'custom', manualMin: n } : { maxAxisType: 'custom', manualMax: n });
  };
  return [
    { type: 'header', title: 'Horizontal Axis Options' },
    { title: 'Show Axis', checked: g.displayXAxis, onClick: () => updateGroup(ctx, { displayXAxis: !g.displayXAxis }) },
    { title: 'Plot Data Right-to-Left', checked: g.rightToLeft, onClick: () => updateGroup(ctx, { rightToLeft: !g.rightToLeft }) },
    { type: 'header', title: 'Vertical Axis Minimum Value Options' },
    { title: 'Automatic for Each Sparkline', checked: g.minAxisType === 'individual', onClick: () => updateGroup(ctx, { minAxisType: 'individual' }) },
    { title: 'Same for All Sparklines', checked: g.minAxisType === 'group', onClick: () => updateGroup(ctx, { minAxisType: 'group' }) },
    { title: 'Custom Value...', checked: g.minAxisType === 'custom', onClick: () => void custom('min') },
    { type: 'header', title: 'Vertical Axis Maximum Value Options' },
    { title: 'Automatic for Each Sparkline', checked: g.maxAxisType === 'individual', onClick: () => updateGroup(ctx, { maxAxisType: 'individual' }) },
    { title: 'Same for All Sparklines', checked: g.maxAxisType === 'group', onClick: () => updateGroup(ctx, { maxAxisType: 'group' }) },
    { title: 'Custom Value...', checked: g.maxAxisType === 'custom', onClick: () => void custom('max') },
  ];
}

/** The Style gallery of the Sparkline tab: darker and lighter accent colours. */
export function sparklineStyles(ctx: Ctx): { key: string; title: string; stroke: string; apply: () => void }[] {
  const theme = ctx.doc.wb.theme;
  return [0, 1].flatMap((row) =>
    [4, 5, 6, 7, 8, 9].map((accent) => {
      const colors = defaultSparklineColors(accent);
      if (row) colors.colorSeries = { theme: accent };
      return {
        key: `${row}-${accent}`,
        title: `Sparkline Style Accent ${accent - 3}${row ? ', Lighter' : ', Darker 50%'}`,
        stroke: `#${resolveColor(colors.colorSeries, theme, '000000')}`,
        apply: () => updateGroup(ctx, { colors }),
      };
    }),
  );
}

export const SP = {
  editData: {
    id: 'sparkline-edit-data',
    label: 'Edit Data',
    icon: 'dataset',
    menu: (ctx) => [
      { title: 'Edit Group Location & Data...', onClick: () => ctx.openDialog({ kind: 'sparkline', type: groupOf(ctx)?.type ?? 'line', groupId: groupOf(ctx)?.id }) },
      { title: "Edit Single Sparkline's Data...", onClick: () => ctx.openDialog({ kind: 'sparkline', type: groupOf(ctx)?.type ?? 'line', groupId: groupOf(ctx)?.id, single: true }) },
    ],
  } as Command,
  line: { id: 'sparkline-type-line', label: 'Line', icon: 'sparkline_line', checked: (ctx) => groupOf(ctx)?.type === 'line', run: (ctx) => updateGroup(ctx, { type: 'line' }) } as Command,
  column: { id: 'sparkline-type-column', label: 'Column', icon: 'sparkline_column', checked: (ctx) => groupOf(ctx)?.type === 'column', run: (ctx) => updateGroup(ctx, { type: 'column' }) } as Command,
  winLoss: { id: 'sparkline-type-winloss', label: 'Win/Loss', icon: 'sparkline_winloss', checked: (ctx) => groupOf(ctx)?.type === 'stacked', run: (ctx) => updateGroup(ctx, { type: 'stacked' }) } as Command,
  high: pointFlag('sparkline-high', 'High Point', 'high'),
  low: pointFlag('sparkline-low', 'Low Point', 'low'),
  negative: pointFlag('sparkline-negative', 'Negative Points', 'negative'),
  first: pointFlag('sparkline-first', 'First Point', 'first'),
  last: pointFlag('sparkline-last', 'Last Point', 'last'),
  markers: pointFlag('sparkline-markers', 'Markers', 'markers'),
  styles: { id: 'sparkline-styles', label: 'Style', icon: 'style', panel: { kind: 'sparkline-styles' } } as Command,
  weight: {
    id: 'sparkline-weight',
    label: 'Weight',
    icon: 'line_weight',
    disabled: (ctx) => groupOf(ctx)?.type !== 'line',
    menu: (ctx) => [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((pt) => ({ title: `${pt} pt`, checked: groupOf(ctx)?.lineWeight === pt, onClick: () => updateGroup(ctx, { lineWeight: pt }) })),
  } as Command,
  axis: { id: 'sparkline-axis', label: 'Axis', icon: 'linear_scale', menu: axisMenu } as Command,
  group: { id: 'sparkline-group', label: 'Group', icon: 'group_work', run: (ctx) => ctx.doc.groupSparklines(ctx.doc.selection.ranges) } as Command,
  ungroup: { id: 'sparkline-ungroup', label: 'Ungroup', icon: 'call_split', run: (ctx) => ctx.doc.ungroupSparklines(ctx.doc.selection.ranges) } as Command,
  clear: {
    id: 'sparkline-clear',
    label: 'Clear',
    icon: 'delete_sweep',
    run: (ctx) => ctx.doc.clearSparklines(ctx.doc.selection.ranges, false),
    menu: (ctx) => [
      { title: 'Clear Selected Sparklines', onClick: () => ctx.doc.clearSparklines(ctx.doc.selection.ranges, false) },
      { title: 'Clear Selected Sparkline Groups', onClick: () => ctx.doc.clearSparklines(ctx.doc.selection.ranges, true) },
    ],
  } as Command,
};

export const sparklineTab: Tab = {
  id: 'sparkline',
  label: 'Sparkline',
  contextual: true,
  visible: (ctx) => ctx.state.sparklineGroup !== null,
  groups: () => [
    { id: 'sparkline-data', label: 'Sparkline', icon: 'dataset', controls: [{ type: 'large', cmd: SP.editData }] },
    { id: 'sparkline-type', label: 'Type', icon: 'sparkline_line', controls: [{ type: 'large', cmd: SP.line }, { type: 'large', cmd: SP.column }, { type: 'large', cmd: SP.winLoss }] },
    {
      id: 'sparkline-show',
      label: 'Show',
      icon: 'visibility',
      controls: [
        { type: 'stack', items: [{ type: 'check', cmd: SP.high }, { type: 'check', cmd: SP.low }, { type: 'check', cmd: SP.negative }] },
        { type: 'stack', items: [{ type: 'check', cmd: SP.first }, { type: 'check', cmd: SP.last }, { type: 'check', cmd: SP.markers }] },
      ],
    },
    {
      id: 'sparkline-style',
      label: 'Style',
      icon: 'style',
      controls: [
        { type: 'large', cmd: SP.styles },
        { type: 'stack', items: [sparkColor('sparkline-color', 'Sparkline Color', 'colorSeries', 'border_color'), sparkColor('sparkline-marker-color', 'Marker Color', 'colorMarkers', 'circle'), { type: 'small', cmd: SP.weight }] },
        { type: 'stack', items: [sparkColor('sparkline-negative-color', 'Negative Points Color', 'colorNegative', 'remove'), sparkColor('sparkline-high-color', 'High Point Color', 'colorHigh', 'north_east'), sparkColor('sparkline-low-color', 'Low Point Color', 'colorLow', 'south_east')] },
      ],
    },
    {
      id: 'sparkline-group',
      label: 'Group',
      icon: 'group_work',
      controls: [{ type: 'large', cmd: SP.axis }, { type: 'stack', items: [{ type: 'small', cmd: SP.group }, { type: 'small', cmd: SP.ungroup }, { type: 'small', cmd: SP.clear, split: true }] }],
    },
  ],
};
