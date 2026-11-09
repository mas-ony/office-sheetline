/** Insert tab: Tables, Charts (dialog and quick menus), Sparklines, Illustrations, Links, Comments, Text. */
import type { MenuItem } from '../menuItems';
import * as cmd from '../commands';
import { TINT, type Command, type Ctx, type Tab } from './types';
import { CHART_CATEGORIES, insertChart, subtypeById, type ChartSubtype } from '../charts';
import { insertTextBox } from '../objects';

const sub = (ctx: Ctx, s: ChartSubtype | undefined): MenuItem | null =>
  s ? { title: s.label, icon: s.icon, testId: `chart-${s.id}`, onClick: () => insertChart(ctx, s) } : null;

/** A quick chart menu: subtypes of some categories, then "More ... Charts". */
function chartMenu(ids: string[][], more: string, category: string): (ctx: Ctx) => MenuItem[] {
  return (ctx) => {
    const items: MenuItem[] = [];
    ids.forEach((group, i) => {
      if (i) items.push({ type: 'line' });
      for (const id of group) {
        const it = sub(ctx, subtypeById(id));
        if (it) items.push(it);
      }
    });
    items.push({ type: 'line' }, { title: more, icon: 'insert_chart', onClick: () => ctx.openDialog({ kind: 'insert-chart', category }) });
    return items;
  };
}

const noData = (ctx: Ctx) => ctx.state.readOnly;

export const I = {
  table: { id: 'insert-table', label: 'Table', icon: 'table_chart', tint: TINT.blue, shortcut: 'Ctrl+T', keywords: 'list object format as table', run: (ctx) => ctx.openDialog({ kind: 'create-table' }), disabled: (ctx) => ctx.state.tableId !== null } as Command,
  charts: { id: 'insert-chart', label: 'Charts', tip: 'Insert Chart', icon: 'insert_chart', tint: TINT.blue, keywords: 'graph plot', run: (ctx) => ctx.openDialog({ kind: 'insert-chart' }), disabled: noData } as Command,
  chartDefault: { id: 'insert-chart-default', label: 'Insert Default Chart', icon: 'bar_chart', shortcut: 'Alt+F1', run: (ctx) => void insertChart(ctx, subtypeById('column-clustered')!) } as Command,
  column: { id: 'chart-column-menu', label: 'Column or Bar Chart', icon: 'bar_chart', tint: TINT.blue, menu: chartMenu([['column-clustered', 'column-stacked', 'column-percent'], ['bar-clustered', 'bar-stacked', 'bar-percent']], 'More Column Charts...', 'column') } as Command,
  line: { id: 'chart-line-menu', label: 'Line or Area Chart', icon: 'show_chart', tint: TINT.blue, menu: chartMenu([['line', 'line-stacked', 'line-percent', 'line-markers', 'line-stacked-markers', 'line-percent-markers'], ['area', 'area-stacked', 'area-percent']], 'More Line Charts...', 'line') } as Command,
  pie: { id: 'chart-pie-menu', label: 'Pie or Doughnut Chart', icon: 'pie_chart', tint: TINT.orange, menu: chartMenu([['pie'], ['doughnut']], 'More Pie Charts...', 'pie') } as Command,
  scatter: { id: 'chart-scatter-menu', label: 'Scatter (X, Y) or Bubble Chart', icon: 'scatter_plot', tint: TINT.blue, menu: chartMenu([['scatter', 'scatter-smooth-markers', 'scatter-smooth', 'scatter-lines-markers', 'scatter-lines'], ['bubble']], 'More Scatter Charts...', 'scatter') } as Command,
  hierarchy: { id: 'chart-hierarchy-menu', label: 'Hierarchy Chart', icon: 'view_quilt', tint: TINT.green, menu: chartMenu([['treemap', 'sunburst']], 'More Hierarchy Charts...', 'treemap') } as Command,
  waterfall: { id: 'chart-waterfall-menu', label: 'Waterfall, Funnel, Stock or Radar Chart', icon: 'waterfall_chart', tint: TINT.blue, menu: chartMenu([['waterfall'], ['funnel'], ['stock-hlc', 'stock-ohlc', 'stock-vhlc', 'stock-vohlc'], ['radar', 'radar-markers', 'radar-filled']], 'More Stock Charts...', 'stock') } as Command,
  statistic: { id: 'chart-statistic-menu', label: 'Statistic Chart', icon: 'equalizer', tint: TINT.blue, menu: chartMenu([['histogram', 'pareto'], ['box-whisker']], 'More Statistical Charts...', 'statistic') } as Command,
  combo: { id: 'chart-combo-menu', label: 'Combo Chart', icon: 'addchart', tint: TINT.blue, menu: chartMenu([['combo-column-line', 'combo-column-line-secondary', 'combo-area-column']], 'Create Custom Combo Chart...', 'combo') } as Command,
  sparkLine: { id: 'sparkline-line', label: 'Line', tip: 'Insert Line Sparkline', icon: 'sparkline_line', tint: TINT.blue, run: (ctx) => ctx.openDialog({ kind: 'sparkline', type: 'line' }), disabled: noData } as Command,
  sparkColumn: { id: 'sparkline-column', label: 'Column', tip: 'Insert Column Sparkline', icon: 'sparkline_column', tint: TINT.blue, run: (ctx) => ctx.openDialog({ kind: 'sparkline', type: 'column' }), disabled: noData } as Command,
  sparkWinLoss: { id: 'sparkline-winloss', label: 'Win/Loss', tip: 'Insert Win/Loss Sparkline', icon: 'sparkline_winloss', tint: TINT.blue, run: (ctx) => ctx.openDialog({ kind: 'sparkline', type: 'stacked' }), disabled: noData } as Command,
  picture: { id: 'insert-picture', label: 'Picture', tip: 'Insert Picture from this device', icon: 'image', tint: TINT.blue, keywords: 'image photo', run: (ctx) => cmd.uploadImage(ctx), disabled: noData } as Command,
  shapes: { id: 'insert-shapes', label: 'Shapes', icon: 'interests', tint: TINT.blue, panel: { kind: 'shapes' }, disabled: noData } as Command,
  textBox: { id: 'insert-text-box', label: 'Text Box', icon: 'text_fields', tint: TINT.gray, run: (ctx) => void insertTextBox(ctx), disabled: noData } as Command,
  link: { id: 'insert-link', label: 'Link', tip: 'Insert Link', icon: 'link', tint: TINT.blue, shortcut: 'Ctrl+K', keywords: 'hyperlink url', run: (ctx) => ctx.openDialog({ kind: 'hyperlink' }), disabled: noData } as Command,
  comment: { id: 'insert-comment', label: 'Comment', tip: 'New Comment', icon: 'add_comment', tint: TINT.purple, run: (ctx) => cmd.newComment(ctx), disabled: noData } as Command,
  symbol: { id: 'insert-symbol', label: 'Symbol', icon: 'functions', keywords: 'character special', run: (ctx) => ctx.openDialog({ kind: 'symbol' }), disabled: noData } as Command,
};

export const insertTab: Tab = {
  id: 'insert',
  label: 'Insert',
  groups: () => [
    { id: 'tables', label: 'Tables', icon: 'table_chart', controls: [{ type: 'large', cmd: I.table }] },
    {
      id: 'charts',
      label: 'Charts',
      icon: 'insert_chart',
      launcher: { id: 'charts-launcher', label: 'See All Charts', run: (ctx) => ctx.openDialog({ kind: 'insert-chart' }) },
      controls: [
        { type: 'large', cmd: I.charts },
        { type: 'stack', items: [{ type: 'small', cmd: I.column, iconOnly: true }, { type: 'small', cmd: I.line, iconOnly: true }, { type: 'small', cmd: I.pie, iconOnly: true }] },
        { type: 'stack', items: [{ type: 'small', cmd: I.scatter, iconOnly: true }, { type: 'small', cmd: I.hierarchy, iconOnly: true }, { type: 'small', cmd: I.waterfall, iconOnly: true }] },
        { type: 'stack', items: [{ type: 'small', cmd: I.statistic, iconOnly: true }, { type: 'small', cmd: I.combo, iconOnly: true }] },
      ],
    },
    {
      id: 'sparklines',
      label: 'Sparklines',
      icon: 'sparkline_line',
      controls: [{ type: 'large', cmd: I.sparkLine }, { type: 'large', cmd: I.sparkColumn }, { type: 'large', cmd: I.sparkWinLoss }],
    },
    {
      id: 'illustrations',
      label: 'Illustrations',
      icon: 'image',
      controls: [{ type: 'large', cmd: I.picture }, { type: 'large', cmd: I.shapes }],
    },
    { id: 'links', label: 'Links', icon: 'link', controls: [{ type: 'large', cmd: I.link }] },
    { id: 'comments', label: 'Comments', icon: 'add_comment', controls: [{ type: 'large', cmd: I.comment }] },
    { id: 'text', label: 'Text', icon: 'text_fields', controls: [{ type: 'large', cmd: I.textBox }, { type: 'large', cmd: I.symbol }] },
  ],
  toolbar: () => [
    { type: 'small', cmd: I.table },
    { type: 'small', cmd: I.charts },
    { type: 'small', cmd: I.column, iconOnly: true },
    { type: 'small', cmd: I.line, iconOnly: true },
    { type: 'small', cmd: I.pie, iconOnly: true },
    { type: 'small', cmd: I.scatter, iconOnly: true },
    { type: 'small', cmd: I.combo, iconOnly: true },
    { type: 'small', cmd: I.sparkLine },
    { type: 'small', cmd: I.sparkColumn },
    { type: 'small', cmd: I.sparkWinLoss },
    { type: 'small', cmd: I.picture },
    { type: 'small', cmd: I.shapes },
    { type: 'small', cmd: I.textBox },
    { type: 'small', cmd: I.link },
    { type: 'small', cmd: I.comment },
  ],
};

export { CHART_CATEGORIES };
