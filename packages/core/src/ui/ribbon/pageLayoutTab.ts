/** Page Layout tab: Themes, Page Setup, Scale to Fit, Sheet Options, Arrange. */
import type { MenuItem } from '../menuItems';
import { TINT, type Command, type Control, type Ctx, type Tab } from './types';
import { COLOR_SCHEMES, FONT_SCHEMES, THEME_PRESETS, colorScheme, fontScheme } from '../../core/themes';
import { MARGIN_PRESETS, PAPER } from '../../core/model/page';
import { makeRange } from '../../core/refs';
import { alignObject, rotateShape } from '../objects';
import { commit } from '../commands';

const page = (ctx: Ctx) => ctx.doc.sheet.page;

/** The tiles of the Themes gallery. */
export function themeTiles(ctx: Ctx): { name: string; active: boolean; testId: string; background: string; color: string; font: string; accents: string[]; apply: () => void }[] {
  const current = ctx.doc.wb.themeName;
  return THEME_PRESETS.map((t) => {
    const colors = colorScheme(t.colors)!;
    const fonts = fontScheme(t.fonts)!;
    const c = colors.colors;
    return {
      name: t.name,
      active: current === t.name,
      testId: `theme-${t.name.replace(/\s+/g, '-')}`,
      background: `#${c.lt1}`,
      color: `#${c.dk1}`,
      font: `"${fonts.major}", sans-serif`,
      accents: [c.accent1, c.accent2, c.accent3, c.accent4, c.accent5, c.accent6].map((x) => `#${x}`),
      apply: () => ctx.doc.applyTheme({ name: t.name, colors: { ...c }, majorFont: fonts.major, minorFont: fonts.minor }),
    };
  });
}

function marginMenu(ctx: Ctx): MenuItem[] {
  const m = page(ctx).margins;
  const same = (x: typeof m) => (['top', 'bottom', 'left', 'right', 'header', 'footer'] as const).every((k) => Math.abs(x[k] - m[k]) < 1e-6);
  return [
    ...MARGIN_PRESETS.map((p) => ({
      title: `${p.label}  —  Top ${p.margins.top}"  Bottom ${p.margins.bottom}"  Left ${p.margins.left}"  Right ${p.margins.right}"`,
      icon: 'margin' as const,
      checked: same(p.margins),
      testId: `margins-${p.name}`,
      onClick: () => ctx.doc.setPageSetup({ margins: { ...p.margins } }),
    })),
    { type: 'line' },
    { title: 'Custom Margins...', onClick: () => ctx.openDialog({ kind: 'page-setup', tab: 'margins' }) },
  ];
}

const scaleOptions = () => [
  { value: 'Automatic' },
  ...Array.from({ length: 9 }, (_, i) => ({ value: `${i + 1} page${i ? 's' : ''}` })),
];

function fitValue(n: number, fit: boolean): string {
  return !fit || !n ? 'Automatic' : `${n} page${n > 1 ? 's' : ''}`;
}

function setFit(ctx: Ctx, axis: 'w' | 'h', v: string): void {
  const n = v === 'Automatic' ? 0 : parseInt(v, 10) || 0;
  const p = page(ctx);
  const width = axis === 'w' ? n : p.fitToPage ? p.fitToWidth : 0;
  const height = axis === 'h' ? n : p.fitToPage ? p.fitToHeight : 0;
  if (!width && !height) ctx.doc.setPageSetup({ fitToPage: false, fitToWidth: 1, fitToHeight: 1 });
  else ctx.doc.setPageSetup({ fitToPage: true, fitToWidth: width, fitToHeight: height });
}

const objectSelected = (ctx: Ctx) => !ctx.selectedObject;

export const P = {
  themes: { id: 'themes', label: 'Themes', icon: 'palette', tint: TINT.blue, panel: { kind: 'themes' } } as Command,
  colors: {
    id: 'theme-colors',
    label: 'Colors',
    tip: 'Theme Colors',
    icon: 'color_lens',
    tint: TINT.orange,
    menu: (ctx) => [
      ...COLOR_SCHEMES.map((s) => ({
        title: s.name,
        swatch: { kind: 'theme', colors: [s.colors.dk2, s.colors.lt2, s.colors.accent1, s.colors.accent2, s.colors.accent3, s.colors.accent4, s.colors.accent5, s.colors.accent6].map((c) => `#${c}`) },
        checked: JSON.stringify(s.colors) === JSON.stringify(ctx.doc.wb.theme.colors),
        onClick: () => ctx.doc.applyTheme({ colors: { ...s.colors } }),
      })),
      { type: 'line' },
      { title: 'Customize Colors...', icon: 'tune', onClick: () => ctx.openDialog({ kind: 'customize-colors' }) },
    ],
  } as Command,
  fonts: {
    id: 'theme-fonts',
    label: 'Fonts',
    tip: 'Theme Fonts',
    icon: 'text_format',
    tint: TINT.blue,
    menu: (ctx) =>
      FONT_SCHEMES.map((f) => ({
        title: `${f.name}  —  ${f.major} / ${f.minor}`,
        swatch: { kind: 'font', family: `"${f.major}", sans-serif` },
        checked: ctx.doc.wb.theme.majorFont === f.major && ctx.doc.wb.theme.minorFont === f.minor,
        onClick: () => ctx.doc.applyTheme({ majorFont: f.major, minorFont: f.minor }),
      })),
  } as Command,
  margins: { id: 'margins', label: 'Margins', icon: 'margin', menu: marginMenu } as Command,
  orientation: {
    id: 'orientation-page',
    label: 'Orientation',
    tip: 'Page Orientation',
    icon: 'page_orientation',
    menu: (ctx) => [
      { title: 'Portrait', icon: 'crop_portrait', checked: page(ctx).orientation === 'portrait', testId: 'orientation-portrait', onClick: () => ctx.doc.setPageSetup({ orientation: 'portrait' }) },
      { title: 'Landscape', icon: 'crop_landscape', checked: page(ctx).orientation === 'landscape', testId: 'orientation-landscape', onClick: () => ctx.doc.setPageSetup({ orientation: 'landscape' }) },
    ],
  } as Command,
  size: {
    id: 'paper-size',
    label: 'Size',
    tip: 'Paper Size',
    icon: 'description',
    menu: (ctx) => [
      ...PAPER.map((p) => ({ title: `${p.label}  ${p.size}`, checked: page(ctx).paperSize === p.code, onClick: () => ctx.doc.setPageSetup({ paperSize: p.code }) })),
      { type: 'line' },
      { title: 'More Paper Sizes...', onClick: () => ctx.openDialog({ kind: 'page-setup', tab: 'page' }) },
    ],
  } as Command,
  printArea: {
    id: 'print-area',
    label: 'Print Area',
    icon: 'print_area',
    menu: (ctx) => [
      { title: 'Set Print Area', testId: 'print-area-set', onClick: () => { commit(ctx); ctx.doc.setPrintArea(ctx.doc.bounded(ctx.doc.range, ctx.doc.sheet) ?? ctx.doc.range); } },
      { title: 'Clear Print Area', testId: 'print-area-clear', disabled: !ctx.doc.printArea(), onClick: () => ctx.doc.setPrintArea(null) },
    ],
  } as Command,
  breaks: {
    id: 'page-breaks',
    label: 'Breaks',
    icon: 'insert_page_break',
    menu: (ctx) => [
      { title: 'Insert Page Break', onClick: () => ctx.doc.insertPageBreak() },
      { title: 'Remove Page Break', onClick: () => ctx.doc.removePageBreak() },
      { title: 'Reset All Page Breaks', onClick: () => ctx.doc.resetPageBreaks() },
    ],
  } as Command,
  titles: { id: 'print-titles', label: 'Print Titles', icon: 'print_titles', run: (ctx) => ctx.openDialog({ kind: 'page-setup', tab: 'sheet' }) } as Command,
  pageSetup: { id: 'page-setup', label: 'Page Setup', run: (ctx) => ctx.openDialog({ kind: 'page-setup', tab: 'page' }) } as Command,
  scaleSetup: { id: 'scale-setup', label: 'Page Setup: Scale', run: (ctx) => ctx.openDialog({ kind: 'page-setup', tab: 'page' }) } as Command,
  sheetSetup: { id: 'sheet-options-setup', label: 'Page Setup: Sheet', run: (ctx) => ctx.openDialog({ kind: 'page-setup', tab: 'sheet' }) } as Command,
  gridView: { id: 'gridlines-view', label: 'View', tip: 'View Gridlines', checked: (ctx) => ctx.doc.sheet.showGridLines, run: (ctx) => ctx.doc.setGridLines(!ctx.doc.sheet.showGridLines) } as Command,
  gridPrint: { id: 'gridlines-print', label: 'Print', tip: 'Print Gridlines', checked: (ctx) => page(ctx).gridLines, run: (ctx) => ctx.doc.setPageSetup({ gridLines: !page(ctx).gridLines }) } as Command,
  headView: { id: 'headings-view', label: 'View', tip: 'View Headings', checked: (ctx) => ctx.doc.sheet.showRowHeaders && ctx.doc.sheet.showColHeaders, run: (ctx) => { const on = !(ctx.doc.sheet.showRowHeaders && ctx.doc.sheet.showColHeaders); ctx.doc.setSheetView({ showRowHeaders: on, showColHeaders: on }); } } as Command,
  headPrint: { id: 'headings-print', label: 'Print', tip: 'Print Headings', checked: (ctx) => page(ctx).headings, run: (ctx) => ctx.doc.setPageSetup({ headings: !page(ctx).headings }) } as Command,
  forward: {
    id: 'bring-forward',
    label: 'Bring Forward',
    icon: 'flip_to_front',
    disabled: objectSelected,
    run: (ctx) => ctx.selectedObject && ctx.doc.reorderObject(ctx.selectedObject, 'forward'),
    menu: (ctx) => [
      { title: 'Bring Forward', icon: 'flip_to_front', onClick: () => ctx.selectedObject && ctx.doc.reorderObject(ctx.selectedObject, 'forward') },
      { title: 'Bring to Front', icon: 'flip_to_front', onClick: () => ctx.selectedObject && ctx.doc.reorderObject(ctx.selectedObject, 'front') },
    ],
  } as Command,
  backward: {
    id: 'send-backward',
    label: 'Send Backward',
    icon: 'flip_to_back',
    disabled: objectSelected,
    run: (ctx) => ctx.selectedObject && ctx.doc.reorderObject(ctx.selectedObject, 'backward'),
    menu: (ctx) => [
      { title: 'Send Backward', icon: 'flip_to_back', onClick: () => ctx.selectedObject && ctx.doc.reorderObject(ctx.selectedObject, 'backward') },
      { title: 'Send to Back', icon: 'flip_to_back', onClick: () => ctx.selectedObject && ctx.doc.reorderObject(ctx.selectedObject, 'back') },
    ],
  } as Command,
  selectionPane: { id: 'selection-pane', label: 'Selection Pane', icon: 'layers', checked: (ctx) => ctx.panel === 'selection', run: (ctx) => ctx.setPanel(ctx.panel === 'selection' ? null : 'selection') } as Command,
  align: {
    id: 'align-objects',
    label: 'Align',
    icon: 'align_horizontal_left',
    disabled: objectSelected,
    menu: (ctx) => [
      { title: 'Align Left', icon: 'align_horizontal_left', onClick: () => alignObject(ctx, 'left') },
      { title: 'Align Center', icon: 'align_horizontal_center', onClick: () => alignObject(ctx, 'center') },
      { title: 'Align Right', icon: 'align_horizontal_right', onClick: () => alignObject(ctx, 'right') },
      { type: 'line' },
      { title: 'Align Top', icon: 'align_vertical_top', onClick: () => alignObject(ctx, 'top') },
      { title: 'Align Middle', icon: 'align_vertical_center', onClick: () => alignObject(ctx, 'middle') },
      { title: 'Align Bottom', icon: 'align_vertical_bottom', onClick: () => alignObject(ctx, 'bottom') },
    ],
  } as Command,
  rotate: {
    id: 'rotate-objects',
    label: 'Rotate',
    icon: 'rotate_right',
    disabled: (ctx) => !ctx.selectedObject || ctx.doc.objectById(ctx.selectedObject)?.kind !== 'shape',
    menu: (ctx) => [
      { title: 'Rotate Right 90°', icon: 'rotate_right', onClick: () => rotateShape(ctx, 'right') },
      { title: 'Rotate Left 90°', icon: 'rotate_left', onClick: () => rotateShape(ctx, 'left') },
      { title: 'Flip Vertical', icon: 'swap_vert', onClick: () => rotateShape(ctx, 'flipV') },
      { title: 'Flip Horizontal', icon: 'swap_horiz', onClick: () => rotateShape(ctx, 'flipH') },
    ],
  } as Command,
};

const widthCombo: Control = {
  type: 'combo',
  id: 'fit-width',
  label: 'Width',
  width: 96,
  value: (ctx) => fitValue(page(ctx).fitToWidth, page(ctx).fitToPage),
  options: scaleOptions,
  onChange: (ctx, v) => setFit(ctx, 'w', v),
};

const heightCombo: Control = {
  type: 'combo',
  id: 'fit-height',
  label: 'Height',
  width: 96,
  value: (ctx) => fitValue(page(ctx).fitToHeight, page(ctx).fitToPage),
  options: scaleOptions,
  onChange: (ctx, v) => setFit(ctx, 'h', v),
};

const scaleCombo: Control = {
  type: 'combo',
  id: 'fit-scale',
  label: 'Scale',
  width: 72,
  editable: true,
  value: (ctx) => `${page(ctx).scale}%`,
  options: () => [10, 25, 50, 75, 100, 125, 150, 200, 400].map((v) => ({ value: `${v}%` })),
  onChange: (ctx, v) => {
    const n = parseInt(v, 10);
    if (!Number.isFinite(n) || n < 10 || n > 400) {
      ctx.notify('Enter a scale between 10% and 400%.');
      return;
    }
    ctx.doc.setPageSetup({ scale: n, fitToPage: false });
  },
  disabled: (ctx) => page(ctx).fitToPage,
};

export const pageLayoutTab: Tab = {
  id: 'page-layout',
  label: 'Page Layout',
  groups: () => [
    {
      id: 'themes',
      label: 'Themes',
      icon: 'palette',
      controls: [{ type: 'large', cmd: P.themes }, { type: 'stack', items: [{ type: 'small', cmd: P.colors }, { type: 'small', cmd: P.fonts }] }],
    },
    {
      id: 'page-setup',
      label: 'Page Setup',
      icon: 'description',
      launcher: P.pageSetup,
      controls: [
        { type: 'large', cmd: P.margins },
        { type: 'large', cmd: P.orientation },
        { type: 'large', cmd: P.size },
        { type: 'large', cmd: P.printArea },
        { type: 'large', cmd: P.breaks },
        { type: 'large', cmd: P.titles },
      ],
    },
    {
      id: 'scale-to-fit',
      label: 'Scale to Fit',
      icon: 'fit_screen',
      launcher: P.scaleSetup,
      controls: [{ type: 'stack', items: [widthCombo, heightCombo, scaleCombo] }],
    },
    {
      id: 'sheet-options',
      label: 'Sheet Options',
      icon: 'grid_on',
      launcher: P.sheetSetup,
      controls: [
        {
          type: 'view',
          id: 'sheet-options-grid',
          view: {
            kind: 'sheet-options',
            columns: [
              { head: 'Gridlines', items: [P.gridView, P.gridPrint] },
              { head: 'Headings', items: [P.headView, P.headPrint] },
            ],
          },
        },
      ],
    },
    {
      id: 'arrange',
      label: 'Arrange',
      icon: 'flip_to_front',
      controls: [
        { type: 'large', cmd: P.forward, split: true },
        { type: 'large', cmd: P.backward, split: true },
        { type: 'large', cmd: P.selectionPane },
        { type: 'stack', items: [{ type: 'small', cmd: P.align }, { type: 'small', cmd: P.rotate }] },
      ],
    },
  ],
  toolbar: () => [
    { type: 'small', cmd: P.themes },
    { type: 'small', cmd: P.colors },
    { type: 'small', cmd: P.fonts },
    { type: 'small', cmd: P.margins },
    { type: 'small', cmd: P.orientation },
    { type: 'small', cmd: P.size },
    { type: 'small', cmd: P.printArea },
    { type: 'small', cmd: P.breaks },
    { type: 'small', cmd: P.titles },
    widthCombo,
    heightCombo,
    scaleCombo,
    { type: 'check', cmd: { ...P.gridPrint, label: 'Print Gridlines' } },
    { type: 'check', cmd: { ...P.headPrint, label: 'Print Headings' } },
    { type: 'small', cmd: P.selectionPane },
  ],
};

export { makeRange };
