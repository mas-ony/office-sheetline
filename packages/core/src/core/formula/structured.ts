/**
 * Structured references to tables: Table1[Col], Table1[[#Headers],[Col]], Table1[[#Data],[A]:[B]],
 * Table1[#All], [@Col] (this row), and bare table names. Resolved against the tables of the workbook.
 */
import type { Range } from '../refs';
import type { Sheet } from '../model/sheet';
import type { Workbook } from '../model/workbook';
import type { SheetTable } from '../model/table';
import { CellError, Errors } from '../values';

export type StructItem = '#All' | '#Data' | '#Headers' | '#Totals' | '#This Row';

export interface StructSpec {
  /** Table name ('' when unqualified, inside a table). */
  table: string;
  items: Set<StructItem>;
  /** First and last column names (null: every column). */
  cols: [string, string] | null;
}

const ITEMS: Record<string, StructItem> = {
  '#all': '#All',
  '#data': '#Data',
  '#headers': '#Headers',
  '#totals': '#Totals',
  '#this row': '#This Row',
};

/** Removes the ' escapes of special characters in a column name. */
export function unescapeName(s: string): string {
  return s.replace(/'(.)/g, '$1');
}

/** Splits at top-level commas (outside brackets, honouring ' escapes). */
function splitTop(s: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let cur = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "'" && i + 1 < s.length) {
      cur += ch + s[++i];
      continue;
    }
    if (ch === '[') depth++;
    else if (ch === ']') depth--;
    if (ch === ',' && depth === 0) {
      out.push(cur.trim());
      cur = '';
      continue;
    }
    cur += ch;
  }
  if (cur.trim() || out.length) out.push(cur.trim());
  return out;
}

function stripBrackets(s: string): string {
  const t = s.trim();
  return t.startsWith('[') && t.endsWith(']') ? t.slice(1, -1) : t;
}

/** Parses the text of a structured reference token. Returns null for malformed text. */
export function parseStructured(text: string, table: string): StructSpec | null {
  const body = text.slice(table.length).trim();
  if (!body.startsWith('[') || !body.endsWith(']')) return null;
  const inner = body.slice(1, -1);
  const spec: StructSpec = { table, items: new Set(), cols: null };
  const addCol = (a: string, b?: string) => {
    const first = unescapeName(stripBrackets(a));
    const last = b !== undefined ? unescapeName(stripBrackets(b)) : first;
    spec.cols = [first, last];
  };
  const part = (p: string) => {
    if (!p) return;
    const t = p.trim();
    if (t.startsWith('@')) {
      spec.items.add('#This Row');
      const rest = t.slice(1).trim();
      if (rest) addCol(rest);
      return;
    }
    // [A]:[B]
    const range = /^(\[(?:[^\]']|'.)*\])\s*:\s*(\[(?:[^\]']|'.)*\])$/.exec(t);
    if (range) {
      addCol(range[1], range[2]);
      return;
    }
    const name = stripBrackets(t);
    const item = ITEMS[name.toLowerCase()];
    if (item) spec.items.add(item);
    else addCol(name);
  };
  if (!inner.trim()) return spec;
  if (inner.trimStart().startsWith('[')) {
    for (const p of splitTop(inner)) part(p);
  } else if (inner.trimStart().startsWith('@')) {
    part(inner);
  } else {
    const item = ITEMS[inner.trim().toLowerCase()];
    if (item) spec.items.add(item);
    else addCol(inner);
  }
  return spec;
}

export function findTable(wb: Workbook, name: string): { sheet: Sheet; table: SheetTable } | undefined {
  const lower = name.toLowerCase();
  for (const sheet of wb.sheets) {
    for (const table of sheet.tables) if (table.name.toLowerCase() === lower) return { sheet, table };
  }
  return undefined;
}

export function tableAtCell(sheet: Sheet, r: number, c: number): SheetTable | undefined {
  return sheet.tables.find((t) => r >= t.ref.r0 && r <= t.ref.r1 && c >= t.ref.c0 && c <= t.ref.c1);
}

export function columnIndex(table: SheetTable, name: string): number {
  const lower = name.toLowerCase();
  return table.columns.findIndex((c) => c.name.toLowerCase() === lower);
}

/** The cells a structured reference covers, or an error. */
export function resolveStructured(wb: Workbook, spec: StructSpec, home: Sheet, row: number, col: number): { sheet: Sheet; range: Range } | CellError {
  const found = spec.table ? findTable(wb, spec.table) : (() => {
    const t = tableAtCell(home, row, col);
    return t ? { sheet: home, table: t } : undefined;
  })();
  if (!found) return Errors.REF;
  const { sheet, table: t } = found;
  let c0 = t.ref.c0;
  let c1 = t.ref.c1;
  if (spec.cols) {
    const a = columnIndex(t, spec.cols[0]);
    const b = columnIndex(t, spec.cols[1]);
    if (a < 0 || b < 0) return Errors.REF;
    c0 = t.ref.c0 + Math.min(a, b);
    c1 = t.ref.c0 + Math.max(a, b);
  }
  const header = t.headerRow ? t.ref.r0 : null;
  const totals = t.totalsRow ? t.ref.r1 : null;
  const dataR0 = t.ref.r0 + (t.headerRow ? 1 : 0);
  const dataR1 = t.ref.r1 - (t.totalsRow ? 1 : 0);
  const items = spec.items.size ? spec.items : new Set<StructItem>(['#Data']);
  if (items.has('#This Row')) {
    if (sheet !== home || row < dataR0 || row > dataR1) return Errors.VALUE;
    return { sheet, range: { r0: row, c0, r1: row, c1 } };
  }
  let r0 = Infinity;
  let r1 = -Infinity;
  const take = (a: number, b: number) => {
    if (b < a) return;
    r0 = Math.min(r0, a);
    r1 = Math.max(r1, b);
  };
  if (items.has('#All')) take(t.ref.r0, t.ref.r1);
  if (items.has('#Headers')) {
    if (header === null) return Errors.REF;
    take(header, header);
  }
  if (items.has('#Data')) take(dataR0, dataR1);
  if (items.has('#Totals')) {
    if (totals === null) return Errors.REF;
    take(totals, totals);
  }
  if (!Number.isFinite(r0)) return Errors.REF;
  return { sheet, range: { r0, c0, r1, c1 } };
}

/** Rewrites structured references in a formula text (rename of tables or columns, conversion to ranges). */
export function rewriteStructured(formula: string, tokens: { type: string; start: number; end: number; value?: unknown }[], fn: (text: string, table: string) => string | null): string {
  let out = '';
  let pos = 0;
  for (const t of tokens) {
    if (t.type !== 'structured' && t.type !== 'name') continue;
    const text = formula.slice(t.start, t.end);
    const next = fn(text, t.type === 'structured' ? String(t.value ?? '') : text);
    if (next === null || next === text) continue;
    out += formula.slice(pos, t.start) + next;
    pos = t.end;
  }
  return pos ? out + formula.slice(pos) : formula;
}
