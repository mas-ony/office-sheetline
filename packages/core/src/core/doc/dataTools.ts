/**
 * Data tools: multi-level sort (values, cell colour, font colour, custom lists), Text to Columns,
 * Remove Duplicates, inserting and deleting cells with shifting, and Fill > Series.
 */
import { SpreadsheetDocument } from '../document';
import type { Cell, Sheet } from '../model/sheet';
import type { Transaction } from '../history';
import { CellError, type Scalar } from '../values';
import { compareScalars } from '../formula/runtime';
import { shiftFormula, translateFormula } from '../formula/adjust';
import { parseInput } from '../format/input';
import { resolveColor } from '../model/color';
import { MAX_COLS, MAX_ROWS, makeRange, rangesIntersect, type Range } from '../refs';
import { satisfies } from '../validation';

export interface SortLevel {
  /** Column (sorting rows) or row (sorting columns) the level sorts by. */
  index: number;
  /** For colour levels: 'asc' puts the colour on top (left), 'desc' at the bottom (right). */
  order: 'asc' | 'desc' | 'custom';
  sortOn: 'values' | 'cellColor' | 'fontColor';
  /** '#RRGGBB' of colour levels ('' = no colour). */
  color?: string;
  /** Order of a custom list level. */
  list?: string[];
}

export interface SortOptions {
  hasHeader: boolean;
  /** Sort left to right (columns) instead of top to bottom. */
  byColumns?: boolean;
  caseSensitive?: boolean;
}

/** Excel's built-in custom lists. */
export const CUSTOM_LISTS: string[][] = [
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
];

export interface TextToColumnsOptions {
  type: 'delimited' | 'fixed';
  delimiters: { tab: boolean; semicolon: boolean; comma: boolean; space: boolean; other: string };
  /** Treat consecutive delimiters as one. */
  consecutive: boolean;
  qualifier: '"' | "'" | '';
  /** Fixed width: break positions in characters. */
  breaks: number[];
  /** Top-left cell of the result (defaults to the source cell). */
  destination?: { r: number; c: number };
  /** Per result column: General (numbers, dates, booleans parsed), Text, or skipped. */
  formats?: ('general' | 'text' | 'skip')[];
}

export interface SeriesOptions {
  by: 'rows' | 'columns';
  type: 'linear' | 'growth' | 'date' | 'autofill';
  unit: 'day' | 'weekday' | 'month' | 'year';
  step: number;
  stop?: number;
  /** Best-fit trend from the selected values (linear and growth). */
  trend: boolean;
}

declare module '../document' {
  interface SpreadsheetDocument {
    sortLevels(range: Range, levels: SortLevel[], opts: SortOptions, sheet?: Sheet): void;
    /** Splits the cells of one column; returns an error message, or null. */
    textToColumns(range: Range, opts: TextToColumnsOptions, sheet?: Sheet): string | null;
    /** Cells the split would overwrite (to confirm first). */
    textToColumnsOverwrites(range: Range, opts: TextToColumnsOptions, sheet?: Sheet): boolean;
    removeDuplicates(range: Range, columns: number[], hasHeader: boolean, sheet?: Sheet): { removed: number; remaining: number };
    insertCells(range: Range, shift: 'down' | 'right', sheet?: Sheet): string | null;
    deleteCells(range: Range, shift: 'up' | 'left', sheet?: Sheet): string | null;
    fillSeriesRange(range: Range, opts: SeriesOptions, sheet?: Sheet): void;
    /** Cells whose values break their data validation rule (Circle Invalid Data). */
    invalidCells(limit?: number, sheet?: Sheet): { r: number; c: number }[];
  }
}

const proto = SpreadsheetDocument.prototype;

// -------------------------------------------------------------------------------------------------
// Sort

function colorOf(doc: SpreadsheetDocument, sheet: Sheet, r: number, c: number, which: 'cellColor' | 'fontColor'): string {
  const style = doc.styleAt(r, c, sheet);
  if (which === 'cellColor') {
    const f = style.fill;
    if (!f) return '';
    if ('gradientXml' in f) return f.color ? `#${resolveColor(f.color, doc.wb.theme, 'FFFFFF')}`.toUpperCase() : '';
    if (f.pattern === 'none') return '';
    return `#${resolveColor(f.fg ?? f.bg, doc.wb.theme, 'FFFFFF')}`.toUpperCase();
  }
  return `#${resolveColor(style.font.color, doc.wb.theme, '000000')}`.toUpperCase();
}

/** Excel order: numbers, text, logicals, errors; blanks always last. */
function compareValues(a: Scalar, b: Scalar, caseSensitive: boolean): number {
  const blankA = a === null || a === '';
  const blankB = b === null || b === '';
  if (blankA || blankB) return blankA === blankB ? 0 : blankA ? 2 : -2;
  const rank = (v: Scalar) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (a instanceof CellError || b instanceof CellError) return 0;
  if (typeof a === 'string' && typeof b === 'string' && caseSensitive) {
    const c = a.localeCompare(b, undefined, { sensitivity: 'case' });
    return c;
  }
  return compareScalars(a, b);
}

proto.sortLevels = function (this: SpreadsheetDocument, range, levels, opts, sheet = this.sheet) {
  const b = this.bounded(range, sheet) ?? range;
  const byCols = !!opts.byColumns;
  const start = (byCols ? b.c0 : b.r0) + (opts.hasHeader ? 1 : 0);
  const end = byCols ? b.c1 : b.r1;
  if (start >= end || !levels.length) return;
  const lines: { index: number; cells: (Cell | undefined)[] }[] = [];
  for (let i = start; i <= end; i++) {
    const cells: (Cell | undefined)[] = [];
    if (byCols) for (let r = b.r0; r <= b.r1; r++) cells.push(sheet.getCell(r, i));
    else for (let c = b.c0; c <= b.c1; c++) cells.push(sheet.getCell(i, c));
    lines.push({ index: i, cells });
  }
  const keyCache = new Map<string, Scalar | string>();
  const key = (line: number, level: SortLevel): Scalar | string => {
    const id = `${line}:${level.index}:${level.sortOn}`;
    if (keyCache.has(id)) return keyCache.get(id)!;
    const r = byCols ? level.index : line;
    const c = byCols ? line : level.index;
    const v = level.sortOn === 'values' ? this.engine.cellValue(sheet, r, c) : colorOf(this, sheet, r, c, level.sortOn);
    keyCache.set(id, v);
    return v;
  };
  lines.sort((x, y) => {
    for (const level of levels) {
      let cmp = 0;
      if (level.sortOn !== 'values') {
        const want = (level.color ?? '').toUpperCase();
        const hx = key(x.index, level) === want ? 0 : 1;
        const hy = key(y.index, level) === want ? 0 : 1;
        cmp = level.order === 'desc' ? hy - hx : hx - hy;
      } else if (level.order === 'custom' && level.list?.length) {
        const pos = (v: Scalar | string) => {
          const i = level.list!.findIndex((item) => item.toLowerCase() === String(v ?? '').toLowerCase());
          return i < 0 ? level.list!.length : i;
        };
        cmp = pos(key(x.index, level)) - pos(key(y.index, level));
      } else {
        const raw = compareValues(key(x.index, level) as Scalar, key(y.index, level) as Scalar, !!opts.caseSensitive);
        // Blanks stay last in both directions.
        cmp = Math.abs(raw) === 2 ? raw : level.order === 'desc' ? -raw : raw;
      }
      if (cmp) return cmp;
    }
    return x.index - y.index;
  });
  this.transact('Sort', (tx) => {
    lines.forEach((line, k) => {
      const to = start + k;
      const delta = to - line.index;
      line.cells.forEach((cell, j) => {
        const r = byCols ? b.r0 + j : to;
        const c = byCols ? to : b.c0 + j;
        const moved = cell?.f !== undefined && delta ? { ...cell, f: translateFormula(cell.f, byCols ? 0 : delta, byCols ? delta : 0) } : cell;
        this.put(tx, sheet, r, c, moved);
      });
    });
  });
};

// -------------------------------------------------------------------------------------------------
// Text to Columns

/** Splits one text into fields. */
export function splitText(text: string, opts: TextToColumnsOptions): string[] {
  if (opts.type === 'fixed') {
    const cuts = [...new Set(opts.breaks.filter((b) => b > 0))].sort((a, b) => a - b);
    const out: string[] = [];
    let prev = 0;
    for (const b of cuts) {
      out.push(text.slice(prev, b));
      prev = b;
    }
    out.push(text.slice(prev));
    return out.map((s) => s.trimEnd());
  }
  const d = opts.delimiters;
  const delims = new Set<string>();
  if (d.tab) delims.add('\t');
  if (d.semicolon) delims.add(';');
  if (d.comma) delims.add(',');
  if (d.space) delims.add(' ');
  if (d.other) delims.add(d.other[0]);
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  let lastWasDelim = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (opts.qualifier && ch === opts.qualifier) {
      if (quoted && text[i + 1] === opts.qualifier) {
        cur += ch;
        i++;
        continue;
      }
      quoted = !quoted;
      lastWasDelim = false;
      continue;
    }
    if (!quoted && delims.has(ch)) {
      if (opts.consecutive && lastWasDelim) continue;
      out.push(cur);
      cur = '';
      lastWasDelim = true;
      continue;
    }
    cur += ch;
    lastWasDelim = false;
  }
  out.push(cur);
  return out;
}

function splitRows(doc: SpreadsheetDocument, sheet: Sheet, range: Range, opts: TextToColumnsOptions): { r: number; fields: string[] }[] {
  const out: { r: number; fields: string[] }[] = [];
  const b = doc.bounded(range, sheet) ?? range;
  for (let r = b.r0; r <= b.r1; r++) {
    const text = doc.displayText(r, b.c0, sheet);
    out.push({ r, fields: text === '' ? [] : splitText(text, opts) });
  }
  return out;
}

proto.textToColumnsOverwrites = function (this: SpreadsheetDocument, range, opts, sheet = this.sheet) {
  const dest = opts.destination ?? { r: range.r0, c: range.c0 };
  for (const { r, fields } of splitRows(this, sheet, range, opts)) {
    const rr = dest.r + (r - range.r0);
    let k = 0;
    for (let i = 0; i < fields.length; i++) {
      if (opts.formats?.[i] === 'skip') continue;
      const c = dest.c + k++;
      if (c === range.c0 && rr === r) continue;
      const v = sheet.getCell(rr, c)?.v;
      if (v !== null && v !== undefined && v !== '') return true;
    }
  }
  return false;
};

proto.textToColumns = function (this: SpreadsheetDocument, range, opts, sheet = this.sheet) {
  if (range.c0 !== range.c1) return 'Text to Columns can convert only one column at a time. The range can be many rows tall but no more than one column wide. Try again by selecting cells in one column only.';
  const dest = opts.destination ?? { r: range.r0, c: range.c0 };
  const rows = splitRows(this, sheet, range, opts);
  this.transact('Text to Columns', (tx) => {
    for (const { r, fields } of rows) {
      const rr = dest.r + (r - range.r0);
      let k = 0;
      fields.forEach((field, i) => {
        const fmt = opts.formats?.[i] ?? 'general';
        if (fmt === 'skip') return;
        const c = dest.c + k++;
        if (c >= MAX_COLS) return;
        const old = sheet.getCell(rr, c);
        let v: Scalar;
        if (fmt === 'text') v = field;
        else {
          const parsed = parseInput(field.trim() === '' ? '' : field, { textFormat: false, date1904: this.wb.date1904 });
          v = parsed.formula !== undefined ? field : parsed.value;
        }
        this.put(tx, sheet, rr, c, v === null || v === '' ? (old?.s ? { v: null, s: old.s } : undefined) : { v, s: old?.s });
      });
    }
  });
  return null;
};

// -------------------------------------------------------------------------------------------------
// Remove Duplicates

proto.removeDuplicates = function (this: SpreadsheetDocument, range, columns, hasHeader, sheet = this.sheet) {
  const b = this.bounded(range, sheet) ?? range;
  const first = b.r0 + (hasHeader ? 1 : 0);
  const keep: number[] = [];
  const seen = new Set<string>();
  for (let r = first; r <= b.r1; r++) {
    const key = columns
      .map((c) => {
        const v = this.engine.cellValue(sheet, r, c);
        return typeof v === 'string' ? `s:${v.toLowerCase()}` : `${typeof v}:${String(v)}`;
      })
      .join('\u0001');
    if (seen.has(key)) continue;
    seen.add(key);
    keep.push(r);
  }
  const removed = b.r1 - first + 1 - keep.length;
  if (!removed) return { removed: 0, remaining: keep.length };
  this.transact('Remove Duplicates', (tx) => {
    const rows = keep.map((r) => Array.from({ length: b.c1 - b.c0 + 1 }, (_, j) => sheet.getCell(r, b.c0 + j)));
    for (let r = first; r <= b.r1; r++) {
      const k = r - first;
      for (let j = 0; j <= b.c1 - b.c0; j++) {
        const c = b.c0 + j;
        if (k < rows.length) {
          const cell = rows[k][j];
          const src = keep[k];
          this.put(tx, sheet, r, c, cell?.f !== undefined && src !== r ? { ...cell, f: translateFormula(cell.f, r - src, 0) } : cell);
        } else this.put(tx, sheet, r, c, undefined);
      }
    }
  });
  return { removed, remaining: keep.length };
};

// -------------------------------------------------------------------------------------------------
// Insert and delete cells

/** Moves the cells of a band (rows `from..` of columns lo..hi, or columns of rows) by `delta`. */
function shiftBand(doc: SpreadsheetDocument, tx: Transaction, sheet: Sheet, axis: 'row' | 'col', from: number, delta: number, lo: number, hi: number): void {
  tx.snapshotSheet(sheet);
  const ext = sheet.dataExtent();
  const last = axis === 'row' ? Math.max(ext.rows, from) : Math.max(ext.cols, from);
  const max = (axis === 'row' ? MAX_ROWS : MAX_COLS) - 1;
  const at = (i: number, k: number): [number, number] => (axis === 'row' ? [i, k] : [k, i]);
  if (delta > 0) {
    for (let i = Math.min(last, max - delta); i >= from; i--) {
      for (let k = lo; k <= hi; k++) {
        const [r, c] = at(i, k);
        const [r2, c2] = at(i + delta, k);
        sheet.setCell(r2, c2, sheet.getCell(r, c));
        sheet.setCell(r, c, undefined);
      }
    }
  } else {
    // Deleted: lines from + delta .. from - 1; the rest moves back.
    for (let i = from + delta; i < from; i++) for (let k = lo; k <= hi; k++) sheet.setCell(...at(i, k), undefined);
    for (let i = from; i <= last; i++) {
      for (let k = lo; k <= hi; k++) {
        const [r, c] = at(i, k);
        const [r2, c2] = at(i + delta, k);
        sheet.setCell(r2, c2, sheet.getCell(r, c));
        sheet.setCell(r, c, undefined);
      }
    }
  }
  const inBand = (r: Range) => (axis === 'row' ? r.c0 >= lo && r.c1 <= hi : r.r0 >= lo && r.r1 <= hi);
  const moveRange = (r: Range): Range | null => {
    if (!inBand(r)) return r;
    const a = axis === 'row' ? r.r0 : r.c0;
    const b = axis === 'row' ? r.r1 : r.c1;
    let na = a;
    let nb = b;
    if (delta > 0) {
      if (a >= from) na += delta;
      if (b >= from) nb += delta;
    } else {
      const ds = from + delta;
      if (a >= ds && b < from) return null;
      na = a >= from ? a + delta : a >= ds ? ds : a;
      nb = b >= from ? b + delta : b >= ds ? ds - 1 : b;
    }
    return axis === 'row' ? { ...r, r0: na, r1: nb } : { ...r, c0: na, c1: nb };
  };
  sheet.comments = sheet.comments.flatMap((t) => {
    const m = moveRange(makeRange(t.r, t.c));
    return m ? [{ ...t, r: m.r0, c: m.c0 }] : [];
  });
  sheet.hyperlinks = sheet.hyperlinks.flatMap((h) => {
    const m = moveRange(h.ref);
    return m ? [{ ...h, ref: m }] : [];
  });
  sheet.setMerges(sheet.merges.flatMap((m) => {
    const n = moveRange(m);
    return n ? [n] : [];
  }));
  sheet.validations = sheet.validations.map((v) => ({ ...v, sqref: v.sqref.flatMap((r) => moveRange(r) ?? []) })).filter((v) => v.sqref.length);
  sheet.condFormats = sheet.condFormats.map((cf) => ({ ...cf, sqref: cf.sqref.flatMap((r) => moveRange(r) ?? []) })).filter((cf) => cf.sqref.length);
  const name = sheet.name;
  doc.adjustAllFormulas(tx, (f, s) => shiftFormula(f, s.name, name, axis, from, delta, { lo, hi }));
}

function blockedByMerges(sheet: Sheet, block: Range, axis: 'row' | 'col', lo: number, hi: number): boolean {
  // A merge crossing the band's edge would be cut in two.
  return sheet.merges.some((m) => rangesIntersect(m, block) && (axis === 'row' ? m.c0 < lo || m.c1 > hi : m.r0 < lo || m.r1 > hi));
}

const MERGE_MESSAGE = 'This operation will cause some merged cells to unmerge. Unmerge the cells first, or select the whole merged area.';

proto.insertCells = function (this: SpreadsheetDocument, range, shift, sheet = this.sheet) {
  if (shift === 'down') {
    const h = range.r1 - range.r0 + 1;
    if (blockedByMerges(sheet, makeRange(range.r0, range.c0, MAX_ROWS - 1, range.c1), 'row', range.c0, range.c1)) return MERGE_MESSAGE;
    this.transact('Insert cells', (tx) => shiftBand(this, tx, sheet, 'row', range.r0, h, range.c0, range.c1));
  } else {
    const w = range.c1 - range.c0 + 1;
    if (blockedByMerges(sheet, makeRange(range.r0, range.c0, range.r1, MAX_COLS - 1), 'col', range.r0, range.r1)) return MERGE_MESSAGE;
    this.transact('Insert cells', (tx) => shiftBand(this, tx, sheet, 'col', range.c0, w, range.r0, range.r1));
  }
  return null;
};

proto.deleteCells = function (this: SpreadsheetDocument, range, shift, sheet = this.sheet) {
  if (shift === 'up') {
    const h = range.r1 - range.r0 + 1;
    if (blockedByMerges(sheet, makeRange(range.r0, range.c0, MAX_ROWS - 1, range.c1), 'row', range.c0, range.c1)) return MERGE_MESSAGE;
    this.transact('Delete cells', (tx) => shiftBand(this, tx, sheet, 'row', range.r1 + 1, -h, range.c0, range.c1));
  } else {
    const w = range.c1 - range.c0 + 1;
    if (blockedByMerges(sheet, makeRange(range.r0, range.c0, range.r1, MAX_COLS - 1), 'col', range.r0, range.r1)) return MERGE_MESSAGE;
    this.transact('Delete cells', (tx) => shiftBand(this, tx, sheet, 'col', range.c1 + 1, -w, range.r0, range.r1));
  }
  return null;
};

// -------------------------------------------------------------------------------------------------
// Fill > Series

function addDate(serial: number, unit: SeriesOptions['unit'], n: number): number {
  if (unit === 'day') return serial + n;
  if (unit === 'weekday') {
    let d = serial;
    let left = Math.abs(n);
    const dir = n < 0 ? -1 : 1;
    while (left > 0) {
      d += dir;
      const dow = ((Math.floor(d) % 7) + 7) % 7; // 1900 system: serial 1 is a Sunday
      if (dow !== 0 && dow !== 1) left--;
    }
    return d;
  }
  const base = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000);
  const y = base.getUTCFullYear();
  const m = base.getUTCMonth();
  const day = base.getUTCDate();
  const months = unit === 'month' ? n : n * 12;
  const target = new Date(Date.UTC(y, m + months, 1));
  const lastDay = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();
  target.setUTCDate(Math.min(day, lastDay));
  return (target.getTime() - Date.UTC(1899, 11, 30)) / 86400000 + (serial - Math.floor(serial));
}

proto.fillSeriesRange = function (this: SpreadsheetDocument, range, opts, sheet = this.sheet) {
  if (opts.type === 'autofill') {
    const source = opts.by === 'columns' ? makeRange(range.r0, range.c0, range.r0, range.c1) : makeRange(range.r0, range.c0, range.r1, range.c0);
    this.select(source);
    this.fill(range, true);
    this.select(range);
    return;
  }
  this.transact('Series', (tx) => {
    const lines = opts.by === 'columns' ? range.c1 - range.c0 + 1 : range.r1 - range.r0 + 1;
    const length = opts.by === 'columns' ? range.r1 - range.r0 + 1 : range.c1 - range.c0 + 1;
    for (let l = 0; l < lines; l++) {
      const at = (k: number): [number, number] => (opts.by === 'columns' ? [range.r0 + k, range.c0 + l] : [range.r0 + l, range.c0 + k]);
      const values: { k: number; v: number }[] = [];
      for (let k = 0; k < length; k++) {
        const v = sheet.getCell(...at(k))?.v;
        if (typeof v === 'number') values.push({ k, v });
      }
      if (!values.length) continue;
      let next: (k: number) => number;
      const start = values[0].v;
      if (opts.trend && values.length > 1 && opts.type !== 'date') {
        const n = values.length;
        const xs = values.map((p) => p.k);
        const ys = opts.type === 'growth' ? values.map((p) => Math.log(p.v)) : values.map((p) => p.v);
        const mx = xs.reduce((a, b) => a + b, 0) / n;
        const my = ys.reduce((a, b) => a + b, 0) / n;
        const slope = xs.reduce((a, x, i) => a + (x - mx) * (ys[i] - my), 0) / (xs.reduce((a, x) => a + (x - mx) ** 2, 0) || 1);
        const icpt = my - slope * mx;
        next = opts.type === 'growth' ? (k) => Math.exp(icpt + slope * k) : (k) => icpt + slope * k;
      } else if (opts.type === 'growth') next = (k) => start * opts.step ** k;
      else if (opts.type === 'date') next = (k) => addDate(start, opts.unit, opts.step * k);
      else next = (k) => start + opts.step * k;
      for (let k = opts.trend ? 0 : 1; k < length; k++) {
        const v = next(k);
        if (!Number.isFinite(v)) break;
        if (opts.stop !== undefined && (opts.step >= 0 || opts.type === 'growth' ? v > opts.stop : v < opts.stop)) break;
        const [r, c] = at(k);
        const old = sheet.getCell(r, c);
        const style = old?.s ?? sheet.getCell(...at(0))?.s;
        this.put(tx, sheet, r, c, { v: Number(v.toPrecision(15)), s: style });
      }
    }
  });
};

// -------------------------------------------------------------------------------------------------
// Circle Invalid Data

proto.invalidCells = function (this: SpreadsheetDocument, limit = 255, sheet = this.sheet) {
  const out: { r: number; c: number }[] = [];
  const display = (_v: Scalar, s: Sheet, r: number, c: number) => this.displayText(r, c, s);
  for (const v of sheet.validations) {
    if (!v.type || v.type === 'none') continue;
    for (const range of v.sqref) {
      const b = this.bounded(range, sheet);
      if (!b) continue;
      for (let r = b.r0; r <= b.r1; r++) {
        for (let c = b.c0; c <= b.c1; c++) {
          const value = this.engine.cellValue(sheet, r, c);
          if (value === null || value === '') continue;
          if (!satisfies(this.engine, sheet, v, r, c, value, display)) out.push({ r, c });
          if (out.length >= limit) return out;
        }
      }
    }
  }
  return out;
};
