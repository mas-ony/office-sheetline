/**
 * What the conditional formatting dialogs of both interfaces share: the quick rules' texts, rule
 * types, threshold and operator choices, new-rule defaults and small colour helpers.
 */
import { type CellIsOperator, type CfRule, type Cfvo, type CfvoType, COLOR_SCALES, DATA_BAR_COLORS, defaultThresholds, type IconSetName, type TextOperator } from '../core/condFormats';
import type { SpreadsheetDocument } from '../core/document';
import { type ColorSpec, resolveColor, type Theme } from '../core/model/color';
import { drawIcon } from '../grid/decor';
import type { CfPresetKind } from './context';

/** CSS colour of a colour choice (undefined for none). */
export const cfHex = (c: ColorSpec | null, theme: Theme) => (c ? `#${resolveColor(c, theme, '000000')}` : undefined);

/** The colour choice of a CSS colour (null for none). */
export const cfSpec = (css: string | undefined): ColorSpec | null => (css ? { rgb: `FF${css.replace('#', '').toUpperCase()}` } : null);

/** An icon of an icon set as an image URL. */
export function cfIconUrl(set: IconSetName, i: number): string {
  const canvas = document.createElement('canvas');
  canvas.width = 32;
  canvas.height = 32;
  const ctx = canvas.getContext('2d');
  if (!ctx) return '';
  drawIcon(ctx, { set, index: i }, 0, 0, 32);
  return canvas.toDataURL();
}

/** The average of the numbers in the selection (the quick rules' starting value). */
export function selectionAverage(doc: SpreadsheetDocument): number {
  let sum = 0;
  let n = 0;
  for (const r of doc.selection.ranges) {
    const b = doc.bounded(r, doc.sheet);
    if (!b) continue;
    doc.sheet.forEachCellIn(b, (_r, _c, cell) => {
      if (typeof cell.v === 'number') {
        sum += cell.v;
        n++;
      }
    });
  }
  return n ? Math.round((sum / n) * 100) / 100 : 0;
}

/** Title and lead text of each quick rule (Greater Than..., Top 10 Items...). */
export const CF_PRESET_TEXT: Record<CfPresetKind, { title: string; lead: string }> = {
  greater: { title: 'Greater Than', lead: 'Format cells that are GREATER THAN:' },
  less: { title: 'Less Than', lead: 'Format cells that are LESS THAN:' },
  between: { title: 'Between', lead: 'Format cells that are BETWEEN:' },
  equal: { title: 'Equal To', lead: 'Format cells that are EQUAL TO:' },
  text: { title: 'Text That Contains', lead: 'Format cells that contain the text:' },
  date: { title: 'A Date Occurring', lead: 'Format cells that contain a date occurring:' },
  duplicate: { title: 'Duplicate Values', lead: 'Format cells that contain:' },
  top10: { title: 'Top 10 Items', lead: 'Format cells that rank in the TOP:' },
  top10pct: { title: 'Top 10%', lead: 'Format cells that rank in the TOP:' },
  bottom10: { title: 'Bottom 10 Items', lead: 'Format cells that rank in the BOTTOM:' },
  bottom10pct: { title: 'Bottom 10%', lead: 'Format cells that rank in the BOTTOM:' },
  above: { title: 'Above Average', lead: 'Format cells that are ABOVE AVERAGE:' },
  below: { title: 'Below Average', lead: 'Format cells that are BELOW AVERAGE:' },
};

/** Formula text of a value box: references and formulas as typed, numbers and text as constants. */
export function cfOperand(text: string): string {
  const t = text.trim();
  if (t.startsWith('=')) return t.slice(1);
  if (t !== '' && Number.isFinite(Number(t))) return t;
  if (/^\$?[A-Z]{1,3}\$?\d+$/i.test(t)) return t;
  return `"${t.replace(/"/g, '""')}"`;
}

// -----------------------------------------------------------------------------------------------
// New / Edit Formatting Rule

export type CfRuleKind = 'values' | 'contain' | 'rank' | 'average' | 'unique' | 'formula';

export const CF_RULE_KINDS: { kind: CfRuleKind; label: string }[] = [
  { kind: 'values', label: 'Format all cells based on their values' },
  { kind: 'contain', label: 'Format only cells that contain' },
  { kind: 'rank', label: 'Format only top or bottom ranked values' },
  { kind: 'average', label: 'Format only values that are above or below average' },
  { kind: 'unique', label: 'Format only unique or duplicate values' },
  { kind: 'formula', label: 'Use a formula to determine which cells to format' },
];

/** The rule type of a rule (or of a ribbon hint for a new one). */
export function cfRuleKindOf(rule: CfRule | null, hint?: string): CfRuleKind {
  const t = rule?.type ?? hint;
  switch (t) {
    case 'colorScale':
    case 'dataBar':
    case 'iconSet':
      return 'values';
    case 'top10':
      return 'rank';
    case 'aboveAverage':
      return 'average';
    case 'duplicateValues':
    case 'uniqueValues':
      return 'unique';
    case 'expression':
      return 'formula';
    default:
      return 'contain';
  }
}

export const CFVO_TYPE_OPTIONS: { type: CfvoType; label: string }[] = [
  { type: 'min', label: 'Lowest Value' },
  { type: 'num', label: 'Number' },
  { type: 'percent', label: 'Percent' },
  { type: 'formula', label: 'Formula' },
  { type: 'percentile', label: 'Percentile' },
  { type: 'max', label: 'Highest Value' },
];

export const CELL_IS_OPERATORS: { op: CellIsOperator; label: string }[] = [
  { op: 'between', label: 'between' },
  { op: 'notBetween', label: 'not between' },
  { op: 'equal', label: 'equal to' },
  { op: 'notEqual', label: 'not equal to' },
  { op: 'greaterThan', label: 'greater than' },
  { op: 'lessThan', label: 'less than' },
  { op: 'greaterThanOrEqual', label: 'greater than or equal to' },
  { op: 'lessThanOrEqual', label: 'less than or equal to' },
];

export const TEXT_OPERATOR_OPTIONS: { op: TextOperator; label: string }[] = [
  { op: 'containsText', label: 'containing' },
  { op: 'notContainsText', label: 'not containing' },
  { op: 'beginsWith', label: 'beginning with' },
  { op: 'endsWith', label: 'ending with' },
];

export type CfContainType = 'cellIs' | 'text' | 'timePeriod' | 'blanks' | 'noBlanks' | 'errors' | 'noErrors';

/** The rule New Rule starts from for a ribbon hint (data bar, colour scale, icon set). */
export function cfInitialRule(hint?: string): CfRule | null {
  return hint === 'dataBar' ? { type: 'dataBar', min: { type: 'min' }, max: { type: 'max' }, color: DATA_BAR_COLORS[0].color, gradient: true, showValue: true } : hint === 'colorScale' ? { type: 'colorScale', points: COLOR_SCALES[0].colors.map((color, i, a) => ({ cfvo: { type: i === 0 ? 'min' : i === a.length - 1 ? 'max' : 'percentile', value: i === 0 || i === a.length - 1 ? undefined : '50' } as Cfvo, color })) } : hint === 'iconSet' ? { type: 'iconSet', iconSet: '3TrafficLights1', reverse: false, showValue: true, thresholds: defaultThresholds('3TrafficLights1') } : null;
}
