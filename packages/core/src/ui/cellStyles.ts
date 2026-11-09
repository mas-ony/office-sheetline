/**
 * Built-in cell styles of the Cell Styles gallery (the Office names), applied as direct formatting.
 * Colours are theme colours where Office uses them, so the styles follow theme changes.
 */
import type { ColorSpec } from '../core/model/color';
import type { BorderLine, StylePatch } from '../core/model/style';

export interface CellStyleDef {
  name: string;
  group: 'Good, Bad and Neutral' | 'Data and Model' | 'Titles and Headings' | 'Themed Cell Styles' | 'Number Format';
  patch: StylePatch;
}

const rgb = (hex: string): ColorSpec => ({ rgb: `FF${hex}` });
const solid = (fg: ColorSpec) => ({ pattern: 'solid', fg, bg: { indexed: 64 } as ColorSpec });
const ln = (style: BorderLine['style'], color: ColorSpec): BorderLine => ({ style, color });
const box = (l: BorderLine) => ({ top: l, bottom: l, left: l, right: l });

/** Resets what the Normal style controls. */
const NORMAL: StylePatch = {
  font: { name: null, size: null, bold: null, italic: null, underline: null, strike: null, color: { theme: 1 }, scheme: 'minor' },
  fill: null,
  border: null,
  numFmt: 'General',
};

function themed(accent: number, tint: number | undefined, white: boolean): StylePatch {
  return { font: { color: white ? { theme: 0 } : { theme: 1 } }, fill: solid(tint === undefined ? { theme: 3 + accent } : { theme: 3 + accent, tint }) };
}

export const CELL_STYLES: CellStyleDef[] = [
  { name: 'Normal', group: 'Good, Bad and Neutral', patch: NORMAL },
  { name: 'Bad', group: 'Good, Bad and Neutral', patch: { font: { color: rgb('9C0006') }, fill: solid(rgb('FFC7CE')) } },
  { name: 'Good', group: 'Good, Bad and Neutral', patch: { font: { color: rgb('006100') }, fill: solid(rgb('C6EFCE')) } },
  { name: 'Neutral', group: 'Good, Bad and Neutral', patch: { font: { color: rgb('9C5700') }, fill: solid(rgb('FFEB9C')) } },

  { name: 'Calculation', group: 'Data and Model', patch: { font: { bold: true, color: rgb('FA7D00') }, fill: solid(rgb('F2F2F2')), border: box(ln('thin', rgb('7F7F7F'))) } },
  { name: 'Check Cell', group: 'Data and Model', patch: { font: { bold: true, color: { theme: 0 } }, fill: solid(rgb('A5A5A5')), border: box(ln('double', rgb('3F3F3F'))) } },
  { name: 'Explanatory Text', group: 'Data and Model', patch: { font: { italic: true, color: rgb('7F7F7F') } } },
  { name: 'Input', group: 'Data and Model', patch: { font: { color: rgb('3F3F76') }, fill: solid(rgb('FFCC99')), border: box(ln('thin', rgb('7F7F7F'))) } },
  { name: 'Linked Cell', group: 'Data and Model', patch: { font: { color: rgb('FA7D00') }, border: { bottom: ln('double', rgb('FF8001')) } } },
  { name: 'Note', group: 'Data and Model', patch: { fill: solid(rgb('FFFFCC')), border: box(ln('thin', rgb('B2B2B2'))) } },
  { name: 'Output', group: 'Data and Model', patch: { font: { bold: true, color: rgb('3F3F3F') }, fill: solid(rgb('F2F2F2')), border: box(ln('thin', rgb('3F3F3F'))) } },
  { name: 'Warning Text', group: 'Data and Model', patch: { font: { color: rgb('FF0000') } } },
  { name: 'Hyperlink', group: 'Data and Model', patch: { font: { color: { theme: 10 }, underline: 'single' } } },
  { name: 'Followed Hyperlink', group: 'Data and Model', patch: { font: { color: { theme: 11 }, underline: 'single' } } },

  { name: 'Heading 1', group: 'Titles and Headings', patch: { font: { bold: true, size: 15, color: { theme: 3 } }, border: { bottom: ln('thick', { theme: 4 }) } } },
  { name: 'Heading 2', group: 'Titles and Headings', patch: { font: { bold: true, size: 13, color: { theme: 3 } }, border: { bottom: ln('thick', { theme: 4, tint: 0.5 }) } } },
  { name: 'Heading 3', group: 'Titles and Headings', patch: { font: { bold: true, size: 11, color: { theme: 3 } }, border: { bottom: ln('medium', { theme: 4, tint: 0.4 }) } } },
  { name: 'Heading 4', group: 'Titles and Headings', patch: { font: { bold: true, size: 11, color: { theme: 3 } } } },
  { name: 'Title', group: 'Titles and Headings', patch: { font: { size: 18, color: { theme: 3 }, scheme: 'major' } } },
  { name: 'Total', group: 'Titles and Headings', patch: { font: { bold: true }, border: { top: ln('thin', { theme: 4 }), bottom: ln('double', { theme: 4 }) } } },

  ...[1, 2, 3, 4, 5, 6].flatMap((a): CellStyleDef[] => [
    { name: `20% - Accent${a}`, group: 'Themed Cell Styles', patch: themed(a, 0.8, false) },
  ]),
  ...[1, 2, 3, 4, 5, 6].map((a): CellStyleDef => ({ name: `40% - Accent${a}`, group: 'Themed Cell Styles', patch: themed(a, 0.6, false) })),
  ...[1, 2, 3, 4, 5, 6].map((a): CellStyleDef => ({ name: `60% - Accent${a}`, group: 'Themed Cell Styles', patch: themed(a, 0.4, true) })),
  ...[1, 2, 3, 4, 5, 6].map((a): CellStyleDef => ({ name: `Accent${a}`, group: 'Themed Cell Styles', patch: themed(a, undefined, true) })),

  { name: 'Comma', group: 'Number Format', patch: { numFmt: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' } },
  { name: 'Comma [0]', group: 'Number Format', patch: { numFmt: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)' } },
  { name: 'Currency', group: 'Number Format', patch: { numFmt: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' } },
  { name: 'Currency [0]', group: 'Number Format', patch: { numFmt: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)' } },
  { name: 'Percent', group: 'Number Format', patch: { numFmt: '0%' } },
];

export const CELL_STYLE_GROUPS: CellStyleDef['group'][] = ['Good, Bad and Neutral', 'Data and Model', 'Titles and Headings', 'Themed Cell Styles', 'Number Format'];

/**
 * Applying a style first resets font, fill and borders (to the workbook's default font), then sets
 * the style's own properties. Number format styles change only the number format.
 */
export function cellStylePatch(def: CellStyleDef, defaultFont: { name: string; size: number }): StylePatch {
  if (def.group === 'Number Format') return def.patch;
  const reset: StylePatch = { ...NORMAL, font: { ...NORMAL.font, name: defaultFont.name, size: defaultFont.size } };
  if (def.name === 'Normal') return reset;
  const { numFmt: _ignored, ...rest } = reset;
  void _ignored;
  return { ...rest, ...def.patch, font: { ...reset.font, ...(def.patch.font ?? {}) } };
}
