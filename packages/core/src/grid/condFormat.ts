/**
 * Conditional formatting for display: evaluates the rule types (cell value, expression, text,
 * dates, blanks/errors, duplicates, top/bottom, above/below average, colour scales, data bars, icon
 * sets) and returns what to draw. Rules are kept unchanged in the file.
 */
import type { SpreadsheetDocument } from '../core/document';
import type { Sheet } from '../core/model/sheet';
import type { FontStyle } from '../core/model/style';
import { attr, child, childrenNamed, numAttr, textContent, type XmlElement } from '../core/xml/xml';
import { translateFormula } from '../core/formula/adjust';
import { ArrayValue, RefValue, compareScalars, toBoolean } from '../core/formula/runtime';
import { CellError, type Scalar } from '../core/values';
import { rangeContains, type Range } from '../core/refs';
import { formatValue } from '../core/format/display';
import { dxfToFormat, iconCount, ruleFromXml, type CfFormat, type CfRule, type Cfvo, type IconSetName, type TimePeriod } from '../core/condFormats';
import { serialFromYMD, weekdayOf, ymdFromSerial } from '../core/formula/dates';

export interface CfIcon {
  set: IconSetName;
  /** 0 = lowest icon of the set. */
  index: number;
}

export interface CfResult {
  fill?: string;
  color?: string;
  font?: Partial<FontStyle>;
  /** Thin border of this colour on all sides. */
  border?: string;
  /** Number format override. */
  numFmt?: string;
  bar?: { ratio: number; color: string; gradient: boolean; border?: string; negative?: boolean; axis?: number };
  icon?: CfIcon;
  /** Data bars and icon sets can hide the cell value. */
  hideValue?: boolean;
}

interface Rule {
  type: string;
  priority: number;
  stop: boolean;
  format?: CfFormat;
  spec: CfRule | null;
  raw: XmlElement;
  ranges: Range[];
}

let version = 0;

/** Invalidates cached rule results (call after data changes). */
export function bumpConditionalFormats(): void {
  version++;
}

interface SheetCache {
  version: number;
  list: unknown;
  dxfs: unknown;
  rules: Rule[];
  results: Map<number, CfResult | null>;
  stats: Map<Rule, { values: number[]; min: number; max: number; avg: number; std: number }>;
}

const caches = new WeakMap<Sheet, SheetCache>();

function rulesOf(doc: SpreadsheetDocument, sheet: Sheet): Rule[] {
  const rules: Rule[] = [];
  for (const cf of sheet.condFormats) {
    for (const r of childrenNamed(cf.raw, 'cfRule')) {
      const dxfId = numAttr(r, 'dxfId');
      rules.push({
        type: attr(r, 'type') ?? '',
        priority: numAttr(r, 'priority') ?? 1000,
        stop: attr(r, 'stopIfTrue') === '1',
        format: dxfId !== undefined ? dxfToFormat(doc.wb.dxfs[dxfId], doc.wb.theme) : undefined,
        spec: ruleFromXml(r, doc.wb.dxfs, doc.wb.theme, sheet.x14Rules),
        raw: r,
        ranges: cf.sqref,
      });
    }
  }
  return rules.sort((a, b) => a.priority - b.priority);
}

function valueAt(sheet: Sheet, r: number, c: number): Scalar {
  return sheet.getCell(r, c)?.v ?? null;
}

function evalFormula(doc: SpreadsheetDocument, sheet: Sheet, formula: string, rule: Rule, r: number, c: number): Scalar {
  const origin = rule.ranges[0];
  const f = translateFormula(formula, r - origin.r0, c - origin.c0);
  const v = doc.engine.evaluateText(f, sheet, r, c);
  if (v instanceof RefValue) return doc.engine.cellValue(v.single.sheet, v.single.range.r0, v.single.range.c0);
  if (v instanceof ArrayValue) return v.get(0, 0);
  return v;
}

function numbersInRanges(sheet: Sheet, ranges: Range[]): number[] {
  const out: number[] = [];
  for (const range of ranges) sheet.forEachCellIn(range, (_r, _c, cell) => {
    if (typeof cell.v === 'number') out.push(cell.v);
  });
  return out;
}

function stats(cache: SheetCache, sheet: Sheet, rule: Rule) {
  let s = cache.stats.get(rule);
  if (!s) {
    const values = numbersInRanges(sheet, rule.ranges).sort((a, b) => a - b);
    const n = values.length;
    const avg = n ? values.reduce((a, b) => a + b, 0) / n : 0;
    const std = n > 1 ? Math.sqrt(values.reduce((a, b) => a + (b - avg) ** 2, 0) / (n - 1)) : 0;
    s = { values, min: values[0] ?? 0, max: values[n - 1] ?? 0, avg, std };
    cache.stats.set(rule, s);
  }
  return s;
}

function cfvoValue(doc: SpreadsheetDocument, sheet: Sheet, cfvo: Cfvo, st: { values: number[]; min: number; max: number }, rule: Rule): number {
  const val = cfvo.value;
  switch (cfvo.type) {
    case 'min':
      return st.min;
    case 'max':
      return st.max;
    case 'percent':
      return st.min + ((st.max - st.min) * Number(val ?? 0)) / 100;
    case 'percentile': {
      const p = Number(val ?? 50) / 100;
      const h = (st.values.length - 1) * p;
      const lo = Math.floor(h);
      return st.values.length ? st.values[lo] + (h - lo) * ((st.values[lo + 1] ?? st.values[lo]) - st.values[lo]) : 0;
    }
    case 'formula': {
      const v = evalFormula(doc, sheet, val ?? '0', rule, rule.ranges[0].r0, rule.ranges[0].c0);
      return typeof v === 'number' ? v : 0;
    }
    default: {
      const n = Number(val ?? 0);
      if (Number.isFinite(n)) return n;
      const v = evalFormula(doc, sheet, val ?? '0', rule, rule.ranges[0].r0, rule.ranges[0].c0);
      return typeof v === 'number' ? v : 0;
    }
  }
}

function mix(a: string, b: string, t: number): string {
  const p = (h: string) => {
    const x = h.replace('#', '');
    return [parseInt(x.slice(0, 2), 16), parseInt(x.slice(2, 4), 16), parseInt(x.slice(4, 6), 16)];
  };
  const [r1, g1, b1] = p(a);
  const [r2, g2, b2] = p(b);
  const m = (x: number, y: number) => Math.round(x + (y - x) * t).toString(16).padStart(2, '0');
  return `#${m(r1, r2)}${m(g1, g2)}${m(b1, b2)}`;
}

/** Start and end (exclusive) serial dates of a "date occurring" period. */
function periodRange(period: TimePeriod, today: number): [number, number] {
  const [y, m] = ymdFromSerial(today);
  // Weeks start on Sunday, as in Excel.
  const weekday = weekdayOf(today);
  const monthStart = (offset: number) => serialFromYMD(y, m + offset, 1);
  switch (period) {
    case 'yesterday': return [today - 1, today];
    case 'today': return [today, today + 1];
    case 'tomorrow': return [today + 1, today + 2];
    case 'last7Days': return [today - 6, today + 1];
    case 'lastWeek': return [today - weekday - 7, today - weekday];
    case 'thisWeek': return [today - weekday, today - weekday + 7];
    case 'nextWeek': return [today - weekday + 7, today - weekday + 14];
    case 'lastMonth': return [monthStart(-1), monthStart(0)];
    case 'thisMonth': return [monthStart(0), monthStart(1)];
    case 'nextMonth': return [monthStart(1), monthStart(2)];
  }
}

function todaySerial(): number {
  const now = new Date();
  return serialFromYMD(now.getFullYear(), now.getMonth() + 1, now.getDate());
}

function evaluate(doc: SpreadsheetDocument, sheet: Sheet, cache: SheetCache, rule: Rule, r: number, c: number): CfResult | null | 'skip' {
  const v = valueAt(sheet, r, c);
  const apply = (): CfResult => {
    const out: CfResult = {};
    const f = rule.format;
    if (!f) return out;
    if (f.fill) out.fill = f.fill;
    if (f.color) out.color = f.color;
    const font: Partial<FontStyle> = {};
    if (f.bold) font.bold = true;
    if (f.italic) font.italic = true;
    if (f.strike) font.strike = true;
    if (f.underline) font.underline = 'single';
    if (Object.keys(font).length) out.font = font;
    if (f.border) out.border = f.border;
    if (f.numFmt) out.numFmt = f.numFmt;
    return out;
  };
  const formulas = childrenNamed(rule.raw, 'formula').map((f) => textContent(f));
  switch (rule.type) {
    case 'cellIs': {
      if (v === null) return null;
      const a = evalFormula(doc, sheet, formulas[0] ?? '0', rule, r, c);
      const b = formulas[1] !== undefined ? evalFormula(doc, sheet, formulas[1], rule, r, c) : null;
      const cmp = compareScalars(v, a);
      const op = attr(rule.raw, 'operator') ?? 'equal';
      const ok =
        op === 'lessThan' ? cmp < 0 : op === 'lessThanOrEqual' ? cmp <= 0 : op === 'greaterThan' ? cmp > 0 : op === 'greaterThanOrEqual' ? cmp >= 0
          : op === 'notEqual' ? cmp !== 0 : op === 'between' ? cmp >= 0 && compareScalars(v, b) <= 0 : op === 'notBetween' ? cmp < 0 || compareScalars(v, b) > 0 : cmp === 0;
      return ok ? apply() : null;
    }
    case 'expression': {
      const res = evalFormula(doc, sheet, formulas[0] ?? 'FALSE', rule, r, c);
      const b = res instanceof CellError ? false : toBoolean(res);
      return b === true ? apply() : null;
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith': {
      const text = (attr(rule.raw, 'text') ?? '').toLowerCase();
      const s = v === null ? '' : formatValue(v, 'General').text.toLowerCase();
      const ok = rule.type === 'containsText' ? s.includes(text) : rule.type === 'notContainsText' ? !s.includes(text) : rule.type === 'beginsWith' ? s.startsWith(text) : s.endsWith(text);
      return ok ? apply() : null;
    }
    case 'timePeriod': {
      if (typeof v !== 'number') return null;
      const [a, b] = periodRange((attr(rule.raw, 'timePeriod') ?? 'today') as TimePeriod, todaySerial());
      const day = Math.floor(v);
      return day >= a && day < b ? apply() : null;
    }
    case 'containsBlanks':
      return v === null || (typeof v === 'string' && !v.trim()) ? apply() : null;
    case 'notContainsBlanks':
      return v !== null && !(typeof v === 'string' && !v.trim()) ? apply() : null;
    case 'containsErrors':
      return v instanceof CellError ? apply() : null;
    case 'notContainsErrors':
      return !(v instanceof CellError) ? apply() : null;
    case 'duplicateValues':
    case 'uniqueValues': {
      if (v === null) return null;
      let count = 0;
      for (const range of rule.ranges) sheet.forEachCellIn(range, (_r, _c, cell) => {
        if (cell.v !== null && compareScalars(cell.v, v) === 0) count++;
      });
      return (rule.type === 'duplicateValues' ? count > 1 : count === 1) ? apply() : null;
    }
    case 'top10': {
      if (typeof v !== 'number') return null;
      const st = stats(cache, sheet, rule);
      const rank = numAttr(rule.raw, 'rank') ?? 10;
      const n = attr(rule.raw, 'percent') === '1' ? Math.max(1, Math.floor((st.values.length * rank) / 100)) : rank;
      const bottom = attr(rule.raw, 'bottom') === '1';
      const threshold = bottom ? st.values[Math.min(n, st.values.length) - 1] : st.values[Math.max(0, st.values.length - n)];
      return (bottom ? v <= threshold : v >= threshold) ? apply() : null;
    }
    case 'aboveAverage': {
      if (typeof v !== 'number') return null;
      const st = stats(cache, sheet, rule);
      const above = attr(rule.raw, 'aboveAverage') !== '0';
      const eq = attr(rule.raw, 'equalAverage') === '1';
      const sd = numAttr(rule.raw, 'stdDev') ?? 0;
      const limit = st.avg + (above ? 1 : -1) * sd * st.std;
      const ok = above ? (eq ? v >= limit : v > limit) : eq ? v <= limit : v < limit;
      return ok ? apply() : null;
    }
    case 'colorScale': {
      const spec = rule.spec;
      if (typeof v !== 'number' || spec?.type !== 'colorScale') return null;
      const st = stats(cache, sheet, rule);
      const points = spec.points.map((p) => cfvoValue(doc, sheet, p.cfvo, st, rule));
      const colors = spec.points.map((p) => p.color);
      if (points.length < 2) return null;
      let fill: string;
      if (points.length === 2 || v <= points[1]) {
        const t = points[1] === points[0] ? 0 : (v - points[0]) / (points[1] - points[0]);
        fill = mix(colors[0], colors[1], Math.max(0, Math.min(1, t)));
      } else {
        const t = points[2] === points[1] ? 1 : (v - points[1]) / (points[2] - points[1]);
        fill = mix(colors[1], colors[2], Math.max(0, Math.min(1, t)));
      }
      if (points.length === 2 && v > points[1]) fill = colors[1];
      return { fill };
    }
    case 'dataBar': {
      const spec = rule.spec;
      if (typeof v !== 'number' || spec?.type !== 'dataBar') return null;
      const st = stats(cache, sheet, rule);
      const lo = cfvoValue(doc, sheet, spec.min, st, rule);
      const hi = cfvoValue(doc, sheet, spec.max, st, rule);
      const negative = v < 0;
      if (lo < 0 && hi > 0) {
        // Bars grow from an axis between the negative and positive values.
        const axis = -lo / (hi - lo);
        const ratio = negative ? Math.min(1, v / lo) * axis : Math.min(1, v / hi) * (1 - axis);
        return { bar: { ratio, color: negative && spec.negativeColor ? spec.negativeColor : spec.color, gradient: spec.gradient, border: spec.border, negative, axis }, hideValue: !spec.showValue };
      }
      const base = Math.min(0, lo);
      const ratio = hi === base ? 1 : (v - base) / (hi - base);
      return { bar: { ratio: Math.max(0, Math.min(1, ratio)), color: spec.color, gradient: spec.gradient, border: spec.border }, hideValue: !spec.showValue };
    }
    case 'iconSet': {
      const spec = rule.spec;
      if (typeof v !== 'number' || spec?.type !== 'iconSet') return null;
      const st = stats(cache, sheet, rule);
      const n = iconCount(spec.iconSet);
      let index = 0;
      spec.thresholds.forEach((t, i) => {
        if (i === 0) return;
        const limit = cfvoValue(doc, sheet, t.cfvo, st, rule);
        if (t.gte ? v >= limit : v > limit) index = i;
      });
      if (spec.reverse) index = n - 1 - index;
      return { icon: { set: spec.iconSet, index }, hideValue: !spec.showValue };
    }
    default:
      return 'skip';
  }
}

/** Returns a function giving the conditional format of a cell, or null when the sheet has none. */
export function condFormatFor(doc: SpreadsheetDocument, sheet: Sheet): ((r: number, c: number) => CfResult | null) | null {
  if (!sheet.condFormats.length) return null;
  let cache = caches.get(sheet);
  if (!cache || cache.version !== version || cache.list !== sheet.condFormats || cache.dxfs !== doc.wb.dxfs) {
    cache = { version, list: sheet.condFormats, dxfs: doc.wb.dxfs, rules: rulesOf(doc, sheet), results: new Map(), stats: new Map() };
    caches.set(sheet, cache);
  }
  const c0 = cache;
  return (r, c) => {
    const key = r * 16384 + c;
    const hit = c0.results.get(key);
    if (hit !== undefined) return hit;
    let out = null as CfResult | null;
    for (const rule of c0.rules) {
      if (!rule.ranges.some((rg) => rangeContains(rg, r, c))) continue;
      let res: CfResult | null | 'skip';
      try {
        res = evaluate(doc, sheet, c0, rule, r, c);
      } catch {
        res = null;
      }
      if (res === 'skip' || !res) continue;
      // Earlier (higher priority) rules win for each property.
      const prev: CfResult = out ?? {};
      out = { ...res, ...prev, font: { ...(res.font ?? {}), ...(prev.font ?? {}) } };
      if (rule.stop) break;
    }
    c0.results.set(key, out);
    return out;
  };
}

export { child };
