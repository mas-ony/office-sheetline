/**
 * Conditional formatting rules as data (the rule types of Excel's New Formatting Rule dialog), and
 * their SpreadsheetML form: <cfRule> elements plus <dxf> differential formats. Rules read from
 * files are converted for display in the Rules Manager; unchanged rules are written back as read.
 */
import { attr, child, childrenNamed, el, numAttr, textContent, txt, type XmlElement } from './xml/xml';
import { readColor } from './xlsx/styles';
import { colorAttrs } from './xlsx/styles';
import type { ColorSpec } from './model/color';
import { resolveColor, type Theme } from './model/color';
import { cellName, type Range } from './refs';
import type { ConditionalFormat, Sheet } from './model/sheet';
import { cloneXml } from './xml/xml';
import { formulaFromFile, formulaToFile } from './xlsx/formulaText';

export type CfvoType = 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula';

export interface Cfvo {
  type: CfvoType;
  value?: string;
}

/** Differential format of a rule (what the matching cells look like). */
export interface CfFormat {
  color?: string;
  fill?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Border colour (all four sides, thin). */
  border?: string;
  numFmt?: string;
}

export type CellIsOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
export type TextOperator = 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith';
export type TimePeriod = 'yesterday' | 'today' | 'tomorrow' | 'last7Days' | 'lastWeek' | 'thisWeek' | 'nextWeek' | 'lastMonth' | 'thisMonth' | 'nextMonth';

export const ICON_SETS = [
  '3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '3Stars', '3Triangles',
  '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '5Boxes',
] as const;
export type IconSetName = (typeof ICON_SETS)[number];

/** Icon sets that exist only in the Excel 2010 extension. */
export const X14_ICON_SETS = new Set<IconSetName>(['3Stars', '3Triangles', '5Boxes']);

export type CfRule =
  | { type: 'cellIs'; operator: CellIsOperator; formula1: string; formula2?: string; format: CfFormat }
  | { type: 'text'; operator: TextOperator; text: string; format: CfFormat }
  | { type: 'timePeriod'; period: TimePeriod; format: CfFormat }
  | { type: 'blanks' | 'noBlanks' | 'errors' | 'noErrors'; format: CfFormat }
  | { type: 'top10'; bottom: boolean; percent: boolean; rank: number; format: CfFormat }
  | { type: 'aboveAverage'; above: boolean; equal: boolean; stdDev?: number; format: CfFormat }
  | { type: 'duplicateValues' | 'uniqueValues'; format: CfFormat }
  | { type: 'expression'; formula: string; format: CfFormat }
  | { type: 'colorScale'; points: { cfvo: Cfvo; color: string }[] }
  | { type: 'dataBar'; min: Cfvo; max: Cfvo; color: string; gradient: boolean; showValue: boolean; border?: string; negativeColor?: string }
  | { type: 'iconSet'; iconSet: IconSetName; reverse: boolean; showValue: boolean; thresholds: { cfvo: Cfvo; gte: boolean }[] };

export interface RuleEntry {
  /** Id of the conditional format (one rule each once edited). */
  id: string;
  rule: CfRule | null;
  /** The rule element (kept for types the editor does not model). */
  raw: XmlElement;
  ranges: Range[];
  priority: number;
  stopIfTrue: boolean;
}

// -----------------------------------------------------------------------------------------------
// Presets of the Home > Conditional Format menus

export const HIGHLIGHT_PRESETS: { label: string; format: CfFormat }[] = [
  { label: 'Light Red Fill with Dark Red Text', format: { fill: '#FFC7CE', color: '#9C0006' } },
  { label: 'Yellow Fill with Dark Yellow Text', format: { fill: '#FFEB9C', color: '#9C5700' } },
  { label: 'Green Fill with Dark Green Text', format: { fill: '#C6EFCE', color: '#006100' } },
  { label: 'Light Red Fill', format: { fill: '#FFC7CE' } },
  { label: 'Red Text', format: { color: '#9C0006' } },
  { label: 'Red Border', format: { border: '#9C0006' } },
];

export const DATA_BAR_COLORS: { label: string; color: string }[] = [
  { label: 'Blue Data Bar', color: '#638EC6' },
  { label: 'Green Data Bar', color: '#63C384' },
  { label: 'Red Data Bar', color: '#FF555A' },
  { label: 'Orange Data Bar', color: '#FFB628' },
  { label: 'Light Blue Data Bar', color: '#008AEF' },
  { label: 'Purple Data Bar', color: '#D6007B' },
];

const G = '#63BE7B';
const Y = '#FFEB84';
const R = '#F8696B';
const W = '#FCFCFF';
const B = '#5A8AC6';

/** Colour scale presets; names list the colour of the highest values first, `colors` go from min to max. */
export const COLOR_SCALES: { label: string; colors: string[] }[] = [
  { label: 'Green - Yellow - Red Color Scale', colors: [R, Y, G] },
  { label: 'Red - Yellow - Green Color Scale', colors: [G, Y, R] },
  { label: 'Green - White - Red Color Scale', colors: [R, W, G] },
  { label: 'Red - White - Green Color Scale', colors: [G, W, R] },
  { label: 'Blue - White - Red Color Scale', colors: [R, W, B] },
  { label: 'Red - White - Blue Color Scale', colors: [B, W, R] },
  { label: 'White - Red Color Scale', colors: [R, W] },
  { label: 'Red - White Color Scale', colors: [W, R] },
  { label: 'Green - White Color Scale', colors: [W, G] },
  { label: 'White - Green Color Scale', colors: [G, W] },
  { label: 'Green - Yellow Color Scale', colors: [Y, G] },
  { label: 'Yellow - Green Color Scale', colors: [G, Y] },
];

export const ICON_SET_GROUPS: { title: string; sets: { name: IconSetName; label: string }[] }[] = [
  {
    title: 'Directional',
    sets: [
      { name: '3Arrows', label: '3 Arrows (Colored)' },
      { name: '3ArrowsGray', label: '3 Arrows (Gray)' },
      { name: '3Triangles', label: '3 Triangles' },
      { name: '4ArrowsGray', label: '4 Arrows (Gray)' },
      { name: '4Arrows', label: '4 Arrows (Colored)' },
      { name: '5ArrowsGray', label: '5 Arrows (Gray)' },
      { name: '5Arrows', label: '5 Arrows (Colored)' },
    ],
  },
  {
    title: 'Shapes',
    sets: [
      { name: '3TrafficLights1', label: '3 Traffic Lights (Unrimmed)' },
      { name: '3TrafficLights2', label: '3 Traffic Lights (Rimmed)' },
      { name: '3Signs', label: '3 Signs' },
      { name: '4TrafficLights', label: '4 Traffic Lights' },
      { name: '4RedToBlack', label: 'Red To Black' },
    ],
  },
  {
    title: 'Indicators',
    sets: [
      { name: '3Symbols', label: '3 Symbols (Circled)' },
      { name: '3Symbols2', label: '3 Symbols (Uncircled)' },
      { name: '3Flags', label: '3 Flags' },
    ],
  },
  {
    title: 'Ratings',
    sets: [
      { name: '3Stars', label: '3 Stars' },
      { name: '4Rating', label: '4 Ratings' },
      { name: '5Quarters', label: '5 Quarters' },
      { name: '5Rating', label: '5 Ratings' },
      { name: '5Boxes', label: '5 Boxes' },
    ],
  },
];

export function iconCount(set: IconSetName): number {
  return Number(set[0]);
}

/** Default thresholds of an icon set: equal percent steps. */
export function defaultThresholds(set: IconSetName): { cfvo: Cfvo; gte: boolean }[] {
  const n = iconCount(set);
  return Array.from({ length: n }, (_, i) => ({ cfvo: { type: 'percent' as CfvoType, value: String(Math.round((i * 100) / n)) }, gte: true }));
}

export const TIME_PERIODS: { period: TimePeriod; label: string }[] = [
  { period: 'yesterday', label: 'Yesterday' },
  { period: 'today', label: 'Today' },
  { period: 'tomorrow', label: 'Tomorrow' },
  { period: 'last7Days', label: 'In the last 7 days' },
  { period: 'lastWeek', label: 'Last week' },
  { period: 'thisWeek', label: 'This week' },
  { period: 'nextWeek', label: 'Next week' },
  { period: 'lastMonth', label: 'Last month' },
  { period: 'thisMonth', label: 'This month' },
  { period: 'nextMonth', label: 'Next month' },
];

// -----------------------------------------------------------------------------------------------
// Differential formats

function hexRgb(css: string): ColorSpec {
  return { rgb: `FF${css.replace('#', '').toUpperCase()}` };
}

export function formatToDxf(f: CfFormat): XmlElement {
  const kids: XmlElement[] = [];
  const font: XmlElement[] = [];
  if (f.bold) font.push(el('b'));
  if (f.italic) font.push(el('i'));
  if (f.strike) font.push(el('strike'));
  if (f.underline) font.push(el('u'));
  if (f.color) font.push(el('color', colorAttrs(hexRgb(f.color))));
  if (font.length) kids.push(el('font', {}, font));
  if (f.numFmt) kids.push(el('numFmt', { numFmtId: 200, formatCode: f.numFmt }));
  if (f.fill) kids.push(el('fill', {}, [el('patternFill', {}, [el('bgColor', colorAttrs(hexRgb(f.fill)))])]));
  if (f.border) {
    const side = (n: string) => el(n, { style: 'thin' }, [el('color', colorAttrs(hexRgb(f.border!)))]);
    kids.push(el('border', {}, [side('left'), side('right'), side('top'), side('bottom')]));
  }
  return el('dxf', {}, kids);
}

export function dxfToFormat(d: XmlElement | undefined, theme: Theme): CfFormat {
  const f: CfFormat = {};
  if (!d) return f;
  const font = child(d, 'font');
  const flag = (n: string) => {
    const e = child(font, n);
    return !!e && attr(e, 'val') !== '0' && attr(e, 'val') !== 'false';
  };
  if (font) {
    if (flag('b')) f.bold = true;
    if (flag('i')) f.italic = true;
    if (flag('strike')) f.strike = true;
    if (child(font, 'u') && attr(child(font, 'u'), 'val') !== 'none') f.underline = true;
    const color = readColor(child(font, 'color'));
    if (color) f.color = `#${resolveColor(color, theme, '000000')}`;
  }
  const pf = child(child(d, 'fill'), 'patternFill');
  const fill = readColor(child(pf, 'bgColor')) ?? readColor(child(pf, 'fgColor'));
  if (fill) f.fill = `#${resolveColor(fill, theme, 'FFFFFF')}`;
  const border = child(d, 'border');
  const anySide = ['left', 'right', 'top', 'bottom'].map((n) => child(border, n)).find((e) => e && attr(e, 'style'));
  if (anySide) f.border = `#${resolveColor(readColor(child(anySide, 'color')), theme, '000000')}`;
  const nf = child(d, 'numFmt');
  if (nf) f.numFmt = attr(nf, 'formatCode');
  return f;
}

// -----------------------------------------------------------------------------------------------
// Rules <-> XML

function cfvoXml(v: Cfvo): XmlElement {
  return el('cfvo', { type: v.type, val: v.type === 'min' || v.type === 'max' ? undefined : formulaToFile(v.value ?? '0') });
}

function readCfvo(e: XmlElement | undefined): Cfvo {
  const type = (attr(e, 'type') ?? 'num') as CfvoType;
  const val = attr(e, 'val');
  return { type: ['min', 'max', 'num', 'percent', 'percentile', 'formula'].includes(type) ? type : 'num', value: val !== undefined ? formulaFromFile(val) : undefined };
}

const quote = (s: string) => `"${s.replace(/"/g, '""')}"`;

/** The formula Excel stores with text, blank, error and date rules (relative to the first cell). */
function impliedFormula(rule: CfRule, first: string): string | undefined {
  switch (rule.type) {
    case 'text': {
      const t = quote(rule.text);
      switch (rule.operator) {
        case 'containsText': return `NOT(ISERROR(SEARCH(${t},${first})))`;
        case 'notContainsText': return `ISERROR(SEARCH(${t},${first}))`;
        case 'beginsWith': return `LEFT(${first},LEN(${t}))=${t}`;
        case 'endsWith': return `RIGHT(${first},LEN(${t}))=${t}`;
      }
      return undefined;
    }
    case 'blanks': return `LEN(TRIM(${first}))=0`;
    case 'noBlanks': return `LEN(TRIM(${first}))>0`;
    case 'errors': return `ISERROR(${first})`;
    case 'noErrors': return `NOT(ISERROR(${first}))`;
    case 'timePeriod':
      switch (rule.period) {
        case 'yesterday': return `FLOOR(${first},1)=TODAY()-1`;
        case 'today': return `FLOOR(${first},1)=TODAY()`;
        case 'tomorrow': return `FLOOR(${first},1)=TODAY()+1`;
        case 'last7Days': return `AND(TODAY()-FLOOR(${first},1)<=6,FLOOR(${first},1)<=TODAY())`;
        case 'lastWeek': return `AND(TODAY()-ROUNDDOWN(${first},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${first},0)<(WEEKDAY(TODAY())+7))`;
        case 'thisWeek': return `AND(TODAY()-ROUNDDOWN(${first},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${first},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
        case 'nextWeek': return `AND(ROUNDDOWN(${first},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${first},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
        case 'lastMonth': return `AND(MONTH(${first})=MONTH(EDATE(TODAY(),0-1)),YEAR(${first})=YEAR(EDATE(TODAY(),0-1)))`;
        case 'thisMonth': return `AND(MONTH(${first})=MONTH(TODAY()),YEAR(${first})=YEAR(TODAY()))`;
        case 'nextMonth': return `AND(MONTH(${first})=MONTH(EDATE(TODAY(),0+1)),YEAR(${first})=YEAR(EDATE(TODAY(),0+1)))`;
      }
      return undefined;
    default:
      return undefined;
  }
}

/** Builds the <cfRule> of a rule. `dxfId` is the index of its format in the workbook's dxfs. */
export function ruleToXml(rule: CfRule, priority: number, stopIfTrue: boolean, dxfId: number | undefined, ranges: Range[]): XmlElement {
  const first = ranges.length ? cellName(ranges[0].r0, ranges[0].c0) : 'A1';
  const base = { priority, stopIfTrue: stopIfTrue ? 1 : undefined };
  const f = (s: string) => el('formula', {}, [txt(formulaToFile(s))]);
  switch (rule.type) {
    case 'cellIs': {
      const kids = [f(rule.formula1)];
      if ((rule.operator === 'between' || rule.operator === 'notBetween') && rule.formula2 !== undefined) kids.push(f(rule.formula2));
      return el('cfRule', { type: 'cellIs', dxfId, ...base, operator: rule.operator }, kids);
    }
    case 'text':
      return el('cfRule', { type: rule.operator, dxfId, ...base, operator: rule.operator === 'notContainsText' ? 'notContains' : rule.operator, text: rule.text }, [f(impliedFormula(rule, first)!)]);
    case 'timePeriod':
      return el('cfRule', { type: 'timePeriod', dxfId, ...base, timePeriod: rule.period }, [f(impliedFormula(rule, first)!)]);
    case 'blanks':
    case 'noBlanks':
    case 'errors':
    case 'noErrors': {
      const type = { blanks: 'containsBlanks', noBlanks: 'notContainsBlanks', errors: 'containsErrors', noErrors: 'notContainsErrors' }[rule.type];
      return el('cfRule', { type, dxfId, ...base }, [f(impliedFormula(rule, first)!)]);
    }
    case 'top10':
      return el('cfRule', { type: 'top10', dxfId, ...base, percent: rule.percent ? 1 : undefined, bottom: rule.bottom ? 1 : undefined, rank: rule.rank });
    case 'aboveAverage':
      return el('cfRule', { type: 'aboveAverage', dxfId, ...base, aboveAverage: rule.above ? undefined : 0, equalAverage: rule.equal ? 1 : undefined, stdDev: rule.stdDev || undefined });
    case 'duplicateValues':
    case 'uniqueValues':
      return el('cfRule', { type: rule.type, dxfId, ...base });
    case 'expression':
      return el('cfRule', { type: 'expression', dxfId, ...base }, [f(rule.formula)]);
    case 'colorScale':
      return el('cfRule', { type: 'colorScale', ...base }, [
        el('colorScale', {}, [...rule.points.map((p) => cfvoXml(p.cfvo)), ...rule.points.map((p) => el('color', colorAttrs(hexRgb(p.color))))]),
      ]);
    case 'dataBar':
      return el('cfRule', { type: 'dataBar', ...base }, [
        el('dataBar', { showValue: rule.showValue ? undefined : 0 }, [cfvoXml(rule.min), cfvoXml(rule.max), el('color', colorAttrs(hexRgb(rule.color)))]),
      ]);
    case 'iconSet':
      return el('cfRule', { type: 'iconSet', ...base }, [
        el('iconSet', { iconSet: rule.iconSet === '3TrafficLights1' ? undefined : rule.iconSet, showValue: rule.showValue ? undefined : 0, reverse: rule.reverse ? 1 : undefined }, rule.thresholds.map((t) => {
          const x = cfvoXml(t.cfvo);
          if (!t.gte) x.attrs.gte = '0';
          return x;
        })),
      ]);
  }
}

/** Reads a <cfRule> into a rule (null for types the editor does not model). `x14` holds Excel 2010 rule details by id. */
export function ruleFromXml(e: XmlElement, dxfs: XmlElement[], theme: Theme, x14?: Map<string, XmlElement>): CfRule | null {
  const type = attr(e, 'type');
  const dxfId = numAttr(e, 'dxfId');
  const format = dxfToFormat(dxfId !== undefined ? dxfs[dxfId] : undefined, theme);
  const formulas = childrenNamed(e, 'formula').map((x) => formulaFromFile(textContent(x)));
  switch (type) {
    case 'cellIs':
      return { type: 'cellIs', operator: (attr(e, 'operator') ?? 'equal') as CellIsOperator, formula1: formulas[0] ?? '0', formula2: formulas[1], format };
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return { type: 'text', operator: type, text: attr(e, 'text') ?? '', format };
    case 'timePeriod':
      return { type: 'timePeriod', period: (attr(e, 'timePeriod') ?? 'today') as TimePeriod, format };
    case 'containsBlanks':
      return { type: 'blanks', format };
    case 'notContainsBlanks':
      return { type: 'noBlanks', format };
    case 'containsErrors':
      return { type: 'errors', format };
    case 'notContainsErrors':
      return { type: 'noErrors', format };
    case 'top10':
      return { type: 'top10', bottom: attr(e, 'bottom') === '1', percent: attr(e, 'percent') === '1', rank: numAttr(e, 'rank') ?? 10, format };
    case 'aboveAverage':
      return { type: 'aboveAverage', above: attr(e, 'aboveAverage') !== '0', equal: attr(e, 'equalAverage') === '1', stdDev: numAttr(e, 'stdDev'), format };
    case 'duplicateValues':
    case 'uniqueValues':
      return { type, format };
    case 'expression':
      return { type: 'expression', formula: formulas[0] ?? 'FALSE', format };
    case 'colorScale': {
      const cs = child(e, 'colorScale');
      const cfvos = childrenNamed(cs, 'cfvo').map(readCfvo);
      const colors = childrenNamed(cs, 'color').map((c) => `#${resolveColor(readColor(c), theme, 'FFFFFF')}`);
      return { type: 'colorScale', points: cfvos.map((cfvo, i) => ({ cfvo, color: colors[i] ?? '#FFFFFF' })) };
    }
    case 'dataBar': {
      const db = child(e, 'dataBar');
      const cfvos = childrenNamed(db, 'cfvo').map(readCfvo);
      const ext = x14Of(e, x14);
      const xdb = ext ? child(ext, 'x14:dataBar') : undefined;
      const border = xdb && attr(xdb, 'border') === '1' ? readColor(child(xdb, 'x14:borderColor')) : undefined;
      const neg = xdb ? readColor(child(xdb, 'x14:negativeFillColor')) : undefined;
      return {
        type: 'dataBar',
        min: cfvos[0] ?? { type: 'min' },
        max: cfvos[1] ?? { type: 'max' },
        color: `#${resolveColor(readColor(child(db, 'color')), theme, '638EC6')}`,
        gradient: !xdb || attr(xdb, 'gradient') !== '0',
        showValue: attr(db, 'showValue') !== '0',
        border: border ? `#${resolveColor(border, theme, '638EC6')}` : undefined,
        negativeColor: neg ? `#${resolveColor(neg, theme, 'FF0000')}` : undefined,
      };
    }
    case 'iconSet': {
      const is = child(e, 'iconSet');
      const set = (attr(is, 'iconSet') ?? '3TrafficLights1') as IconSetName;
      return {
        type: 'iconSet',
        iconSet: ICON_SETS.includes(set) ? set : '3TrafficLights1',
        reverse: attr(is, 'reverse') === '1',
        showValue: attr(is, 'showValue') !== '0',
        thresholds: childrenNamed(is, 'cfvo').map((c) => ({ cfvo: readCfvo(c), gte: attr(c, 'gte') !== '0' })),
      };
    }
    default:
      return null;
  }
}

/** The Excel 2010 (x14) rule that extends a classic rule, by the id in its extension list. */
export function x14Of(e: XmlElement, x14?: Map<string, XmlElement>): XmlElement | undefined {
  if (!x14) return undefined;
  for (const ext of childrenNamed(child(e, 'extLst'), 'ext')) {
    const id = textContent(child(ext, 'x14:id') ?? ext).trim();
    if (id && x14.has(id)) return x14.get(id);
  }
  return undefined;
}

/** Short description for the Rules Manager. */
export function describeRule(rule: CfRule | null, raw: XmlElement): string {
  if (!rule) return `Rule (${attr(raw, 'type') ?? 'unknown'})`;
  const opText: Record<CellIsOperator, string> = {
    between: 'between', notBetween: 'not between', equal: '=', notEqual: '<>', greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=',
  };
  switch (rule.type) {
    case 'cellIs':
      return rule.operator === 'between' || rule.operator === 'notBetween'
        ? `Cell value ${opText[rule.operator]} ${rule.formula1} and ${rule.formula2 ?? ''}`
        : `Cell value ${opText[rule.operator]} ${rule.formula1}`;
    case 'text':
      return { containsText: 'Cell value contains', notContainsText: 'Cell value does not contain', beginsWith: 'Cell value begins with', endsWith: 'Cell value ends with' }[rule.operator] + ` "${rule.text}"`;
    case 'timePeriod':
      return TIME_PERIODS.find((p) => p.period === rule.period)?.label ?? 'Date occurring';
    case 'blanks': return 'Blanks';
    case 'noBlanks': return 'No Blanks';
    case 'errors': return 'Errors';
    case 'noErrors': return 'No Errors';
    case 'top10': return `${rule.bottom ? 'Bottom' : 'Top'} ${rule.rank}${rule.percent ? '%' : ''}`;
    case 'aboveAverage': return `${rule.above ? 'Above' : 'Below'}${rule.equal ? ' or Equal to' : ''} Average${rule.stdDev ? ` (${rule.stdDev} std dev)` : ''}`;
    case 'duplicateValues': return 'Duplicate Values';
    case 'uniqueValues': return 'Unique Values';
    case 'expression': return `Formula: =${rule.formula}`;
    case 'colorScale': return `Graded Color Scale`;
    case 'dataBar': return 'Data Bar';
    case 'iconSet': return 'Icon Set';
  }
}

let counter = 1;

export function newCfId(): string {
  return `cf${Date.now().toString(36)}${(counter++).toString(36)}`;
}

/** Gives every block an id and one rule each (blocks with several rules are split). */
export function normalizeCondFormats(sheet: Sheet): void {
  const out: ConditionalFormat[] = [];
  for (const cf of sheet.condFormats) {
    const rules = cf.raw.children.filter((k): k is XmlElement => k.type === 'element' && k.name === 'cfRule');
    if (rules.length <= 1) {
      out.push(cf.id ? cf : { ...cf, id: newCfId() });
      continue;
    }
    rules.forEach((rule, i) => {
      const raw = el('conditionalFormatting', { ...cf.raw.attrs }, [cloneXml(rule)]);
      out.push({ sqref: cf.sqref.map((r) => ({ ...r })), raw, id: i === 0 && cf.id ? cf.id : newCfId() });
    });
  }
  sheet.condFormats = out;
}

export { hexRgb };
