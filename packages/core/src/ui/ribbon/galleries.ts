/** Ribbon galleries and menus: paste options, cell styles, conditional format presets, borders. */
import type { MenuItem, Preview } from '../menuItems';
import type { IconName } from '../icons';
import type { Ctx } from './types';
import * as cmd from '../commands';
import { cellStylePatch, type CellStyleDef } from '../cellStyles';
import { resolveColor } from '../../core/model/color';
import { COLOR_SCALES, DATA_BAR_COLORS, ICON_SET_GROUPS, defaultThresholds, type CfRule, type IconSetName } from '../../core/condFormats';
import { drawIcon } from '../../grid/decor';
import type { PasteOptions } from '../../core/document';

/** Inline styles as CSS property names (camel case) and values with their units. */
export type StyleMap = Record<string, string>;

// -----------------------------------------------------------------------------------------------
// Paste

export interface PasteChoice {
  label: string;
  icon: IconName;
  opts: PasteOptions;
  id: string;
}

export const PASTE_GROUPS: { title: string; items: PasteChoice[] }[] = [
  {
    title: 'Paste',
    items: [
      { id: 'all', label: 'Paste', icon: 'content_paste', opts: { what: 'all' } },
      { id: 'formulas', label: 'Formulas', icon: 'functions', opts: { what: 'formulas' } },
      { id: 'formulas-numfmt', label: 'Formulas & Number Formatting', icon: 'percent', opts: { what: 'formulasNumFmt' } },
      { id: 'no-borders', label: 'No Borders', icon: 'border_clear', opts: { what: 'allExceptBorders' } },
      { id: 'col-widths', label: 'Keep Source Column Widths', icon: 'width_normal', opts: { what: 'colWidths' } },
      { id: 'transpose', label: 'Transpose', icon: 'swap_horiz', opts: { what: 'all', transpose: true } },
    ],
  },
  {
    title: 'Paste Values',
    items: [
      { id: 'values', label: 'Values', icon: '123', opts: { what: 'values' } },
      { id: 'values-numfmt', label: 'Values & Number Formatting', icon: 'numbers', opts: { what: 'valuesNumFmt' } },
      { id: 'values-fmt', label: 'Values & Source Formatting', icon: 'format_paint', opts: { what: 'valuesSourceFmt' } },
    ],
  },
  {
    title: 'Other Paste Options',
    items: [
      { id: 'formats', label: 'Formatting', icon: 'format_paint', opts: { what: 'formats' } },
      { id: 'link', label: 'Paste Link', icon: 'link', opts: { what: 'all', link: true } },
    ],
  },
];

// -----------------------------------------------------------------------------------------------
// Cell styles

/** The look of a cell style on its gallery button. */
export function cellStylePreview(ctx: Ctx, def: CellStyleDef): StyleMap {
  const theme = ctx.doc.wb.theme;
  const p = def.patch;
  const css: StyleMap = {};
  const fill = p.fill && 'fg' in p.fill && p.fill.fg ? `#${resolveColor(p.fill.fg, theme, 'FFFFFF')}` : undefined;
  if (fill) css.background = fill;
  if (p.font?.color) css.color = `#${resolveColor(p.font.color, theme, '000000')}`;
  if (p.font?.bold) css.fontWeight = '700';
  if (p.font?.italic) css.fontStyle = 'italic';
  if (p.font?.underline) css.textDecoration = 'underline';
  if (p.font?.size) css.fontSize = `${Math.min(18, Math.max(11, p.font.size))}px`;
  const b = p.border;
  if (b && typeof b === 'object') {
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const l = b[side];
      if (l) css[`border${side[0].toUpperCase()}${side.slice(1)}`] = `${l.style === 'double' ? 3 : l.style === 'thick' ? 3 : l.style === 'medium' ? 2 : 1}px ${l.style === 'double' ? 'double' : 'solid'} #${resolveColor(l.color, theme, '000000')}`;
    }
  }
  return css;
}

export function applyCellStyle(ctx: Ctx, def: CellStyleDef): void {
  const base = ctx.doc.wb.styles.defaultStyle.font;
  cmd.format(ctx, cellStylePatch(def, { name: base.name, size: base.size }), `Cell style: ${def.name}`);
}

// -----------------------------------------------------------------------------------------------
// Conditional formatting presets

function svg(markup: string, w = 30, h = 30): Preview {
  return { kind: 'svg', markup, width: w, height: h };
}

function dataBarSvg(color: string, gradient: boolean): string {
  const bars = [0.95, 0.7, 0.45, 0.25];
  const id = `g${color.slice(1)}`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="30" height="30" viewBox="0 0 30 30"><defs><linearGradient id="${id}"><stop offset="0" stop-color="${color}"/><stop offset="1" stop-color="#fff"/></linearGradient></defs><rect x=".5" y=".5" width="29" height="29" fill="#fff" stroke="#c8c8c8"/>${bars.map((b, i) => `<rect x="3" y="${3 + i * 6.5}" width="${24 * b}" height="5" fill="${gradient ? `url(#${id})` : color}" stroke="${color}" stroke-width=".5"/>`).join('')}</svg>`;
}

function colorScaleSvg(colors: string[]): string {
  const cells: string[] = [];
  const mix = (a: string, b: string, t: number) => {
    const p = (h: string) => [1, 3, 5].map((k) => parseInt(h.slice(k, k + 2), 16));
    const x = p(a);
    const y = p(b);
    return `#${x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('')}`;
  };
  for (let r = 0; r < 3; r++) {
    const t = (2 - r) / 2;
    const c = colors.length === 3 ? (t >= 0.5 ? mix(colors[1], colors[2], (t - 0.5) * 2) : mix(colors[0], colors[1], t * 2)) : mix(colors[0], colors[1], t);
    for (let k = 0; k < 3; k++) cells.push(`<rect x="${2 + k * 9}" y="${2 + r * 9}" width="8" height="8" fill="${c}"/>`);
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="30" height="30" viewBox="0 0 30 30"><rect x=".5" y=".5" width="29" height="29" fill="#fff" stroke="#c8c8c8"/>${cells.join('')}</svg>`;
}

/** Draws the icons of an icon set on a canvas (16 CSS pixels per icon). */
export function paintIconSet(c: HTMLCanvasElement, set: IconSetName): void {
  const g = c.getContext('2d');
  if (!g) return;
  const n = Number(set[0]);
  const dpr = window.devicePixelRatio || 1;
  c.width = n * 16 * dpr;
  c.height = 16 * dpr;
  g.scale(dpr, dpr);
  for (let i = 0; i < n; i++) drawIcon(g, { set, index: n - 1 - i }, i * 16 + 1, 1, 14);
}

function addRule(ctx: Ctx, rule: CfRule): void {
  cmd.commit(ctx);
  ctx.doc.addCondFormat(rule);
}

export function condFormatMenu(ctx: Ctx): MenuItem[] {
  const preset = (title: string, kind: Parameters<typeof cfPreset>[1]): MenuItem => ({ title, onClick: () => cfPreset(ctx, kind), testId: `cf-${kind}` });
  const more = (type?: string): MenuItem => ({ title: 'More Rules...', onClick: () => ctx.openDialog({ kind: 'cf-rule', ruleType: type }) });
  const table = ctx.doc.tableAt(ctx.doc.selection.active.r, ctx.doc.selection.active.c);
  return [
    {
      title: 'Highlight Cells Rules',
      icon: 'rule',
      testId: 'cf-highlight',
      submenu: [
        preset('Greater Than...', 'greater'),
        preset('Less Than...', 'less'),
        preset('Between...', 'between'),
        preset('Equal To...', 'equal'),
        preset('Text that Contains...', 'text'),
        preset('A Date Occurring...', 'date'),
        preset('Duplicate Values...', 'duplicate'),
        { type: 'line' },
        more('cellIs'),
      ],
    },
    {
      title: 'Top/Bottom Rules',
      icon: 'leaderboard',
      testId: 'cf-topbottom',
      submenu: [
        preset('Top 10 Items...', 'top10'),
        preset('Top 10%...', 'top10pct'),
        preset('Bottom 10 Items...', 'bottom10'),
        preset('Bottom 10%...', 'bottom10pct'),
        preset('Above Average...', 'above'),
        preset('Below Average...', 'below'),
        { type: 'line' },
        more('top10'),
      ],
    },
    { type: 'line' },
    {
      title: 'Data Bars',
      icon: 'bar_chart',
      testId: 'cf-databars',
      submenu: [
        {
          type: 'gallery',
          title: 'Gradient Fill',
          columns: 3,
          gallery: DATA_BAR_COLORS.map((d) => ({
            title: d.label,
            testId: `cf-bar-gradient-${d.label.split(' ')[0].toLowerCase()}`,
            preview: svg(dataBarSvg(d.color, true)),
            onClick: () => addRule(ctx, { type: 'dataBar', min: { type: 'min' }, max: { type: 'max' }, color: d.color, gradient: true, showValue: true }),
          })),
        },
        {
          type: 'gallery',
          title: 'Solid Fill',
          columns: 3,
          gallery: DATA_BAR_COLORS.map((d) => ({
            title: d.label,
            testId: `cf-bar-solid-${d.label.split(' ')[0].toLowerCase()}`,
            preview: svg(dataBarSvg(d.color, false)),
            onClick: () => addRule(ctx, { type: 'dataBar', min: { type: 'min' }, max: { type: 'max' }, color: d.color, gradient: false, showValue: true }),
          })),
        },
        { type: 'line' },
        more('dataBar'),
      ],
    },
    {
      title: 'Color Scales',
      icon: 'palette',
      testId: 'cf-colorscales',
      submenu: [
        {
          type: 'gallery',
          title: 'Color Scales',
          columns: 4,
          gallery: COLOR_SCALES.map((sc, i) => ({
            title: sc.label,
            testId: `cf-scale-${i}`,
            preview: svg(colorScaleSvg(sc.colors)),
            onClick: () => addRule(ctx, {
              type: 'colorScale',
              points: sc.colors.length === 3
                ? [{ cfvo: { type: 'min' }, color: sc.colors[0] }, { cfvo: { type: 'percentile', value: '50' }, color: sc.colors[1] }, { cfvo: { type: 'max' }, color: sc.colors[2] }]
                : [{ cfvo: { type: 'min' }, color: sc.colors[0] }, { cfvo: { type: 'max' }, color: sc.colors[1] }],
            }),
          })),
        },
        { type: 'line' },
        more('colorScale'),
      ],
    },
    {
      title: 'Icon Sets',
      icon: 'star_half',
      testId: 'cf-iconsets',
      submenu: [
        ...ICON_SET_GROUPS.map((g): MenuItem => ({
          type: 'gallery',
          title: g.title,
          columns: 2,
          gallery: g.sets.map((s) => ({
            title: s.label,
            testId: `cf-icons-${s.name}`,
            preview: { kind: 'icon-set', set: s.name },
            onClick: () => addRule(ctx, { type: 'iconSet', iconSet: s.name, reverse: false, showValue: true, thresholds: defaultThresholds(s.name) }),
          })),
        })),
        { type: 'line' },
        more('iconSet'),
      ],
    },
    { type: 'line' },
    { title: 'New Rule...', icon: 'post_add', testId: 'cf-new', onClick: () => ctx.openDialog({ kind: 'cf-rule' }) },
    {
      title: 'Clear Rules',
      icon: 'delete_sweep',
      testId: 'cf-clear',
      submenu: [
        { title: 'Clear Rules from Selected Cells', testId: 'cf-clear-selection', onClick: () => ctx.doc.clearCondFormats('selection') },
        { title: 'Clear Rules from Entire Sheet', testId: 'cf-clear-sheet', onClick: () => ctx.doc.clearCondFormats('sheet') },
        { title: 'Clear Rules from This Table', disabled: !table, onClick: () => ctx.doc.clearCondFormats('table') },
      ],
    },
    { title: 'Manage Rules...', icon: 'tune', testId: 'cf-manage', onClick: () => ctx.openDialog({ kind: 'cf-manager' }) },
  ];
}

function cfPreset(ctx: Ctx, preset: 'greater' | 'less' | 'between' | 'equal' | 'text' | 'date' | 'duplicate' | 'top10' | 'top10pct' | 'bottom10' | 'bottom10pct' | 'above' | 'below'): void {
  cmd.commit(ctx);
  ctx.openDialog({ kind: 'cf-preset', preset });
}

// -----------------------------------------------------------------------------------------------
// Borders

function borderIcon(kind: string): Preview {
  const s = 'currentColor';
  const dot = (x: number, y: number) => `<rect x="${x}" y="${y}" width="1" height="1" fill="#9a9a9a"/>`;
  const dots: string[] = [];
  for (let i = 2; i <= 16; i += 2) {
    dots.push(dot(i, 2), dot(i, 16), dot(2, i), dot(16, i), dot(i, 9), dot(9, i));
  }
  const parts: Record<string, string> = {
    bottom: `<path d="M1.5 16.5h15" stroke="${s}"/>`,
    top: `<path d="M1.5 1.5h15" stroke="${s}"/>`,
    left: `<path d="M1.5 1.5v15" stroke="${s}"/>`,
    right: `<path d="M16.5 1.5v15" stroke="${s}"/>`,
    none: '',
    all: `<path d="M1.5 1.5h15v15h-15zM9 1.5v15M1.5 9h15" stroke="${s}" fill="none"/>`,
    outside: `<path d="M1.5 1.5h15v15h-15z" stroke="${s}" fill="none"/>`,
    'thick-box': `<path d="M2 2h14v14H2z" stroke="${s}" stroke-width="2" fill="none"/>`,
    'bottom-double': `<path d="M1.5 14.5h15M1.5 16.5h15" stroke="${s}"/>`,
    'thick-bottom': `<path d="M1.5 16h15" stroke="${s}" stroke-width="2"/>`,
    'top-bottom': `<path d="M1.5 1.5h15M1.5 16.5h15" stroke="${s}"/>`,
    'top-thick-bottom': `<path d="M1.5 1.5h15" stroke="${s}"/><path d="M1.5 16h15" stroke="${s}" stroke-width="2"/>`,
    'top-double-bottom': `<path d="M1.5 1.5h15M1.5 14.5h15M1.5 16.5h15" stroke="${s}"/>`,
    inside: `<path d="M9 1.5v15M1.5 9h15" stroke="${s}"/>`,
  };
  return svg(`<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 18 18">${dots.join('')}${parts[kind] ?? ''}</svg>`, 18, 18);
}

export function borderMenu(ctx: Ctx): MenuItem[] {
  const it = (title: string, kind: cmd.BorderPreset): MenuItem => ({ title, swatch: borderIcon(kind), testId: `border-${kind}`, onClick: () => cmd.applyBorder(ctx, kind) });
  return [
    it('Bottom Border', 'bottom'),
    it('Top Border', 'top'),
    it('Left Border', 'left'),
    it('Right Border', 'right'),
    { type: 'line' },
    it('No Border', 'none'),
    it('All Borders', 'all'),
    it('Outside Borders', 'outside'),
    it('Thick Box Border', 'thick-box'),
    { type: 'line' },
    it('Bottom Double Border', 'bottom-double'),
    it('Thick Bottom Border', 'thick-bottom'),
    it('Top and Bottom Border', 'top-bottom'),
    it('Top and Thick Bottom Border', 'top-thick-bottom'),
    it('Top and Double Bottom Border', 'top-double-bottom'),
    { type: 'line' },
    { title: 'More Borders...', swatch: borderIcon('all'), testId: 'border-more', onClick: () => ctx.openDialog({ kind: 'format-cells', tab: 'border' }) },
  ];
}
