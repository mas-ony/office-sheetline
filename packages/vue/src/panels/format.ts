/** Colours and line dashes of the Format pane. */
import { type ColorSpec, type LineDash, resolveColor, type UIContextValue } from '@sheetline/core/views';

export const hexOf = (ctx: UIContextValue, c: ColorSpec | null) => (c ? `#${resolveColor(c, ctx.doc.wb.theme, '000000')}` : null);
export const specOf = (hex: string | null | undefined): ColorSpec | null => (hex ? { rgb: `FF${hex.replace('#', '').toUpperCase()}` } : null);

export const DASHES: { dash: LineDash; label: string }[] = [
  { dash: 'solid', label: 'Solid' },
  { dash: 'dashed', label: 'Dash' },
  { dash: 'dotted', label: 'Round Dot' },
  { dash: 'dash-dot', label: 'Dash Dot' },
];

export type Line = { color: string; width: number; dash?: LineDash };
