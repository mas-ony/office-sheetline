/**
 * The Format Cells dialog's lists and number-format logic, shared by both interfaces: number
 * categories and their codes, border edges and line previews, fill patterns, alignments.
 */
import type { BorderLine, BorderLineStyle, HAlign, UnderlineType, VAlign } from '../core/model/style';

export type FormatTab = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection';

// -----------------------------------------------------------------------------------------------
// Number formats

export type NumberCategory = 'General' | 'Number' | 'Currency' | 'Accounting' | 'Date' | 'Time' | 'Percentage' | 'Fraction' | 'Scientific' | 'Text' | 'Special' | 'Custom';

export const NUMBER_CATEGORIES: NumberCategory[] = ['General', 'Number', 'Currency', 'Accounting', 'Date', 'Time', 'Percentage', 'Fraction', 'Scientific', 'Text', 'Special', 'Custom'];

export const NUMBER_CATEGORY_DESCRIPTIONS: Record<NumberCategory, string> = {
  General: 'General format cells have no specific number format.',
  Number: 'Number is used for general display of numbers. Currency and Accounting offer specialized formatting for monetary value.',
  Currency: 'Currency formats are used for general monetary values. Use Accounting formats to align decimal points in a column.',
  Accounting: 'Accounting formats line up the currency symbols and decimal points in a column.',
  Date: 'Date formats display date and time serial numbers as date values.',
  Time: 'Time formats display date and time serial numbers as time values.',
  Percentage: 'Percentage formats multiply the cell value by 100 and display the result with a percent symbol.',
  Fraction: 'Fraction formats display numbers as fractions.',
  Scientific: 'Scientific formats display numbers in exponential notation.',
  Text: 'Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered.',
  Special: 'Special formats are useful for tracking list and database values.',
  Custom: 'Type the number format code, using one of the existing codes as a starting point.',
};

export const CURRENCY_SYMBOLS: { label: string; code: string }[] = [
  { label: 'None', code: '' },
  { label: '$', code: '"$"' },
  { label: '€', code: '[$€-x-euro2] ' },
  { label: '£', code: '[$£-809]' },
  { label: '¥', code: '[$¥-804]' },
  { label: 'Rp', code: '[$Rp-421]' },
  { label: '₹', code: '[$₹-4009]' },
  { label: 'CHF', code: '[$CHF-807] ' },
];

export const DATE_FORMAT_TYPES: { label: string; code: string }[] = [
  { label: '*3/14/2012', code: 'm/d/yyyy' },
  { label: '*Wednesday, March 14, 2012', code: '[$-F800]dddd, mmmm dd, yyyy' },
  { label: '3/14', code: 'm/d;@' },
  { label: '3/14/12', code: 'm/d/yy;@' },
  { label: '03/14/12', code: 'mm/dd/yy;@' },
  { label: '14-Mar', code: 'd-mmm;@' },
  { label: '14-Mar-12', code: 'd-mmm-yy;@' },
  { label: '14-Mar-2012', code: 'dd-mmm-yyyy' },
  { label: 'Mar-12', code: 'mmm-yy;@' },
  { label: 'March-12', code: 'mmmm-yy;@' },
  { label: 'March 14, 2012', code: '[$-409]mmmm d, yyyy;@' },
  { label: '3/14/12 1:30 PM', code: 'm/d/yy h:mm AM/PM;@' },
  { label: '3/14/12 13:30', code: 'm/d/yy h:mm;@' },
  { label: 'M', code: 'mmmmm;@' },
  { label: 'M-12', code: 'mmmmm-yy;@' },
  { label: '2012-03-14', code: 'yyyy-mm-dd;@' },
];

export const TIME_FORMAT_TYPES: { label: string; code: string }[] = [
  { label: '*1:30:55 PM', code: '[$-F400]h:mm:ss AM/PM' },
  { label: '13:30', code: 'h:mm;@' },
  { label: '1:30 PM', code: 'h:mm AM/PM;@' },
  { label: '13:30:55', code: 'h:mm:ss;@' },
  { label: '1:30:55 PM', code: 'h:mm:ss AM/PM;@' },
  { label: '30:55.2', code: 'mm:ss.0;@' },
  { label: '37:30:55', code: '[h]:mm:ss;@' },
  { label: '3/14/12 1:30 PM', code: 'm/d/yy h:mm AM/PM;@' },
  { label: '3/14/12 13:30', code: 'm/d/yy h:mm;@' },
];

export const FRACTION_FORMAT_TYPES: { label: string; code: string }[] = [
  { label: 'Up to one digit (1/4)', code: '# ?/?' },
  { label: 'Up to two digits (21/25)', code: '# ??/??' },
  { label: 'Up to three digits (312/943)', code: '# ???/???' },
  { label: 'As halves (1/2)', code: '# ?/2' },
  { label: 'As quarters (2/4)', code: '# ?/4' },
  { label: 'As eighths (4/8)', code: '# ?/8' },
  { label: 'As sixteenths (8/16)', code: '# ??/16' },
  { label: 'As tenths (3/10)', code: '# ?/10' },
  { label: 'As hundredths (30/100)', code: '# ??/100' },
];

export const SPECIAL_FORMAT_TYPES: { label: string; code: string }[] = [
  { label: 'Zip Code', code: '00000' },
  { label: 'Zip Code + 4', code: '00000-0000' },
  { label: 'Phone Number', code: '[<=9999999]###-####;(###) ###-####' },
  { label: 'Social Security Number', code: '000-00-0000' },
];

export interface NumberFormatState {
  category: NumberCategory;
  decimals: number;
  separator: boolean;
  symbol: string;
  negative: number;
  /** Code chosen from a type list (date, time, fraction, special) or typed (custom). */
  code: string;
}

const dec = (n: number) => (n > 0 ? `.${'0'.repeat(Math.min(30, n))}` : '');

/** The format code of the Number tab settings. */
export function numberCode(s: NumberFormatState): string {
  const d = Math.max(0, Math.min(30, s.decimals));
  const neg = (base: string) => [base, `${base};[Red]${base}`, `${base}_);\\(${base}\\)`, `${base}_);[Red]\\(${base}\\)`][s.negative] ?? base;
  switch (s.category) {
    case 'General':
      return 'General';
    case 'Number':
      return s.negative === 0 ? `${s.separator ? '#,##0' : '0'}${dec(d)}` : neg(`${s.separator ? '#,##0' : '0'}${dec(d)}`);
    case 'Currency':
      return neg(`${s.symbol}#,##0${dec(d)}`);
    case 'Accounting': {
      const n = `#,##0${dec(d)}`;
      const sym = s.symbol;
      return `_(${sym}* ${n}_);_(${sym}* \\(${n}\\);_(${sym}* "-"${d ? '?'.repeat(d) : ''}_);_(@_)`;
    }
    case 'Percentage':
      return `0${dec(d)}%`;
    case 'Scientific':
      return `0${dec(d)}E+00`;
    case 'Text':
      return '@';
    default:
      return s.code || 'General';
  }
}

/** The Number tab settings showing a format code (Custom when no category fits). */
export function numberStateOf(code: string): NumberFormatState {
  const base: NumberFormatState = { category: 'Custom', decimals: 2, separator: false, symbol: '"$"', negative: 0, code };
  const decimals = /\.(0+)/.exec(code)?.[1].length ?? 0;
  if (/^general$/i.test(code)) return { ...base, category: 'General' };
  if (code === '@') return { ...base, category: 'Text' };
  if (/^0(\.0+)?%$/.test(code)) return { ...base, category: 'Percentage', decimals };
  if (/^0(\.0+)?E\+00$/i.test(code)) return { ...base, category: 'Scientific', decimals };
  if (/^_\(/.test(code) || /^_-/.test(code)) {
    const sym = CURRENCY_SYMBOLS.find((x) => x.code && code.includes(x.code.trim()))?.code ?? '';
    return { ...base, category: 'Accounting', decimals, symbol: sym };
  }
  const first = code.split(';')[0].replace(/_\)$/, '');
  const negative = /;\[Red\]\\?\(/.test(code) ? 3 : /;\\?\(/.test(code) ? 2 : /;\[Red\]/.test(code) ? 1 : 0;
  const sym = CURRENCY_SYMBOLS.find((x) => x.code && first.startsWith(x.code));
  if (sym && /#,##0(\.0+)?$/.test(first)) return { ...base, category: 'Currency', decimals, symbol: sym.code, negative };
  if (/^(#,##)?0(\.0+)?$/.test(first)) return { ...base, category: 'Number', decimals, separator: first.startsWith('#,##'), negative };
  for (const [cat, list] of [['Date', DATE_FORMAT_TYPES], ['Time', TIME_FORMAT_TYPES], ['Fraction', FRACTION_FORMAT_TYPES], ['Special', SPECIAL_FORMAT_TYPES]] as const) {
    if (list.some((t) => t.code === code)) return { ...base, category: cat, code };
  }
  if (/[dmy]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/[hs]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]/g, ''))) return { ...base, category: 'Custom', code };
  return base;
}

// -----------------------------------------------------------------------------------------------
// Borders, fill, alignment and font

export type BorderEdge = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV' | 'diagUp' | 'diagDown';

/** undefined = unchanged, null = removed. */
export type BorderEdges = Partial<Record<BorderEdge, BorderLine | null>>;

export const BORDER_LINE_PREVIEW: Record<BorderLineStyle, string> = {
  thin: '1', medium: '2', thick: '3', dashed: '1 4 2', dotted: '1 1 1', double: 'd', hair: '1 1 1',
  mediumDashed: '2 6 3', dashDot: '1 6 2 2 2', mediumDashDot: '2 6 2 2 2', dashDotDot: '1 6 2 2 2 2 2', mediumDashDotDot: '2 6 2 2 2 2 2', slantDashDot: '2 8 2 3 2',
};

export const FILL_PATTERNS = [
  'none', 'solid', 'gray125', 'gray0625', 'lightGray', 'mediumGray', 'darkGray', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',
  'lightGrid', 'lightTrellis', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
];

export const H_ALIGN_OPTIONS: { value: HAlign; label: string }[] = [
  { value: 'general', label: 'General' },
  { value: 'left', label: 'Left (Indent)' },
  { value: 'center', label: 'Center' },
  { value: 'right', label: 'Right (Indent)' },
  { value: 'fill', label: 'Fill' },
  { value: 'justify', label: 'Justify' },
  { value: 'centerContinuous', label: 'Center Across Selection' },
  { value: 'distributed', label: 'Distributed (Indent)' },
];

export const V_ALIGN_OPTIONS: { value: VAlign; label: string }[] = [
  { value: 'top', label: 'Top' },
  { value: 'center', label: 'Center' },
  { value: 'bottom', label: 'Bottom' },
  { value: 'justify', label: 'Justify' },
  { value: 'distributed', label: 'Distributed' },
];

export const UNDERLINE_OPTIONS: { value: UnderlineType | 'none'; label: string }[] = [
  { value: 'none', label: 'None' },
  { value: 'single', label: 'Single' },
  { value: 'double', label: 'Double' },
  { value: 'singleAccounting', label: 'Single Accounting' },
  { value: 'doubleAccounting', label: 'Double Accounting' },
];

/** Excel's dialog degrees (-90..90) from the file's textRotation (0..180, 255 = stacked). */
export function textRotationDegrees(rotation: number | undefined): number {
  if (!rotation || rotation === 255) return 0;
  return rotation <= 90 ? rotation : 90 - rotation;
}
