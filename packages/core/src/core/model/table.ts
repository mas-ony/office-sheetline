/** Excel tables (ListObjects): a named range with a header row, optional totals row, a style and filters. */
import type { Range } from '../refs';
import type { ColumnFilter } from './sheet';
import type { XmlElement } from '../xml/xml';

export type TotalsFunction = 'none' | 'sum' | 'average' | 'count' | 'countNums' | 'max' | 'min' | 'stdDev' | 'var' | 'custom';

/** SUBTOTAL function numbers of the totals row functions (visible cells only, as Excel writes them). */
export const TOTALS_SUBTOTAL: Record<Exclude<TotalsFunction, 'none' | 'custom'>, number> = {
  average: 101,
  count: 103,
  countNums: 102,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
};

export const TOTALS_LABELS: { fn: TotalsFunction; label: string }[] = [
  { fn: 'none', label: 'None' },
  { fn: 'average', label: 'Average' },
  { fn: 'count', label: 'Count' },
  { fn: 'countNums', label: 'Count Numbers' },
  { fn: 'max', label: 'Max' },
  { fn: 'min', label: 'Min' },
  { fn: 'sum', label: 'Sum' },
  { fn: 'stdDev', label: 'StdDev' },
  { fn: 'var', label: 'Var' },
];

export interface TableColumn {
  /** Column id in the file (unique within the table). */
  id: number;
  name: string;
  totalsFunction?: TotalsFunction;
  totalsLabel?: string;
  /** Formula of a custom totals cell (without '='). */
  totalsFormula?: string;
  /** Formula filled into every data cell of the column (without '='). */
  calculatedFormula?: string;
  /** The source element (other attributes and children are kept). */
  raw?: XmlElement;
}

export interface TableStyleInfo {
  /** Built-in style name such as TableStyleMedium2; '' for no style. */
  name: string;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
}

export interface SheetTable {
  /** Workbook-wide id (the table part's id attribute). */
  id: number;
  name: string;
  /** The whole table: header row, data and totals row. */
  ref: Range;
  headerRow: boolean;
  totalsRow: boolean;
  columns: TableColumn[];
  style: TableStyleInfo;
  /** Filter buttons in the header row. */
  autoFilter: boolean;
  /** Filter criteria by absolute column index. */
  filters: Map<number, ColumnFilter>;
  /** Typing next to the table extends it (an editor option, shown as "Allow Auto Expand"; not saved in the file). */
  autoExpand?: boolean;
  comment?: string;
  /** The source <table> element and part, kept for attributes the editor does not model. */
  raw?: XmlElement;
  sourcePath?: string;
}

export const DEFAULT_TABLE_STYLE = 'TableStyleMedium2';

export function cloneTable(t: SheetTable): SheetTable {
  return {
    ...t,
    ref: { ...t.ref },
    columns: t.columns.map((c) => ({ ...c })),
    style: { ...t.style },
    filters: new Map([...t.filters].map(([k, v]) => [k, { ...v, values: v.values ? [...v.values] : undefined, conditions: v.conditions?.map((x) => ({ ...x })) }])),
  };
}

/** Rows of the data body (without header and totals rows); r1 < r0 when the table has no data rows. */
export function tableDataRange(t: SheetTable): Range {
  return { r0: t.ref.r0 + (t.headerRow ? 1 : 0), c0: t.ref.c0, r1: t.ref.r1 - (t.totalsRow ? 1 : 0), c1: t.ref.c1 };
}

export function tableHeaderRow(t: SheetTable): number | null {
  return t.headerRow ? t.ref.r0 : null;
}

export function tableTotalsRow(t: SheetTable): number | null {
  return t.totalsRow ? t.ref.r1 : null;
}

/** The filtered area of a table in auto filter form (header row included, totals row excluded). */
export function tableFilterRef(t: SheetTable): Range {
  return { r0: t.ref.r0, c0: t.ref.c0, r1: t.ref.r1 - (t.totalsRow ? 1 : 0), c1: t.ref.c1 };
}

/** Excel's rules for table names. Returns an error message or null. */
export function tableNameError(name: string): string | null {
  if (!name) return 'The table name cannot be empty.';
  if (name.length > 255) return 'The table name is too long.';
  if (!/^[A-Za-z_\\À-￿][A-Za-z0-9_.\\À-￿]*$/.test(name)) return 'The table name must start with a letter or an underscore and contain no spaces or symbols.';
  if (/^[A-Za-z]{1,3}\d+$/.test(name) || /^[Rr]\d*[Cc]\d*$/.test(name)) return 'The table name cannot look like a cell reference.';
  if (/^(r|c)$/i.test(name)) return 'This name is reserved.';
  return null;
}

/** Column names are unique (case-insensitively) within a table: "Name", "Name2", ... */
export function uniqueColumnName(base: string, taken: Iterable<string>, self?: string): string {
  const used = new Set([...taken].map((n) => n.toLowerCase()));
  if (self) used.delete(self.toLowerCase());
  let name = base.trim() || 'Column';
  if (!used.has(name.toLowerCase())) return name;
  let n = 2;
  while (used.has(`${name}${n}`.toLowerCase())) n++;
  name = `${name}${n}`;
  return name;
}

/** Characters that need escaping with ' inside structured references. */
export function escapeColumnName(name: string): string {
  return name.replace(/['#[\]]/g, (m) => `'${m}`);
}

export function columnRef(name: string): string {
  return /[\s'#[\]@,:+\-*/^&=<>(){}"!%$]/.test(name) || /^\d/.test(name) ? `[${escapeColumnName(name)}]` : escapeColumnName(name);
}
