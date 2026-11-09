/**
 * Table commands of the document: create, resize, convert to range, rename, style and options,
 * totals row, filters and sorting. Header cells always hold the column names; typing next to a
 * table extends it, and a formula typed into an empty table column fills the column.
 */
import { SpreadsheetDocument } from '../document';
import type { Transaction } from '../history';
import type { Cell, ColumnFilter, Sheet } from '../model/sheet';
import {
  DEFAULT_TABLE_STYLE, TOTALS_SUBTOTAL, escapeColumnName, tableDataRange, tableNameError, uniqueColumnName,
  type SheetTable, type TableColumn, type TotalsFunction,
} from '../model/table';
import { cellName, makeRange, rangesIntersect, type Range } from '../refs';
import { tableCellDecor } from '../tableStyles';
import { rgbSpec } from '../model/color';
import type { BorderLine, StylePatch } from '../model/style';
import { tokenize } from '../formula/tokenizer';
import { findTable, parseStructured, resolveStructured, rewriteStructured } from '../formula/structured';
import { translateFormula } from '../formula/adjust';

export interface TableOptions {
  headerRow: boolean;
  totalsRow: boolean;
  autoFilter: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  autoExpand: boolean;
}

declare module '../document' {
  interface SpreadsheetDocument {
    tableAt(r: number, c: number, sheet?: Sheet): SheetTable | undefined;
    tableById(id: number): { sheet: Sheet; table: SheetTable } | undefined;
    /** Formats the range as a table. Returns null (and a reason) when it cannot. */
    createTable(range: Range, opts?: { hasHeaders?: boolean; style?: string; name?: string }, sheet?: Sheet): SheetTable | string;
    resizeTable(id: number, ref: Range): string | null;
    renameTable(id: number, name: string): string | null;
    convertTableToRange(id: number): void;
    deleteTable(id: number, clearCells?: boolean): void;
    setTableStyle(id: number, style: string): void;
    setTableOptions(id: number, patch: Partial<TableOptions>): void;
    setTotalsFunction(id: number, colIndex: number, fn: TotalsFunction): void;
    setTableFilter(id: number, col: number, filter: ColumnFilter | null): void;
    clearTableFilters(id: number): void;
    sortTable(id: number, col: number, ascending: boolean): void;
    /** A table name not used by tables or defined names: Table1, Table2... */
    uniqueTableName(base?: string): string;
  }
}

const proto = SpreadsheetDocument.prototype;

function absRange(r: Range): string {
  const a = cellName(r.r0, r.c0, true, true);
  return r.r0 === r.r1 && r.c0 === r.c1 ? a : `${a}:${cellName(r.r1, r.c1, true, true)}`;
}

function allTables(doc: SpreadsheetDocument): { sheet: Sheet; table: SheetTable }[] {
  return doc.wb.sheets.flatMap((sheet) => sheet.tables.map((table) => ({ sheet, table })));
}

proto.tableAt = function (this: SpreadsheetDocument, r, c, sheet = this.sheet) {
  return sheet.tables.find((t) => r >= t.ref.r0 && r <= t.ref.r1 && c >= t.ref.c0 && c <= t.ref.c1);
};

proto.tableById = function (this: SpreadsheetDocument, id) {
  return allTables(this).find((x) => x.table.id === id);
};

proto.uniqueTableName = function (this: SpreadsheetDocument, base = 'Table') {
  const taken = new Set([...allTables(this).map((x) => x.table.name.toLowerCase()), ...this.wb.names.map((n) => n.name.toLowerCase())]);
  let n = 1;
  while (taken.has(`${base}${n}`.toLowerCase())) n++;
  return `${base}${n}`;
};

function nextTableId(doc: SpreadsheetDocument): number {
  return Math.max(0, ...allTables(doc).map((x) => x.table.id)) + 1;
}

/** Writes the column names into the header cells (as text). */
function writeHeaders(doc: SpreadsheetDocument, tx: Transaction, sheet: Sheet, t: SheetTable): void {
  if (!t.headerRow) return;
  t.columns.forEach((col, i) => {
    const r = t.ref.r0;
    const c = t.ref.c0 + i;
    const cell = sheet.getCell(r, c);
    if (cell?.v === col.name && cell.f === undefined) return;
    doc.put(tx, sheet, r, c, { v: col.name, s: cell?.s });
  });
}

function isEmptyRow(sheet: Sheet, r: number, c0: number, c1: number): boolean {
  for (let c = c0; c <= c1; c++) {
    const cell = sheet.getCell(r, c);
    if (cell && (cell.v !== null || cell.f !== undefined)) return false;
  }
  return true;
}

function totalsFormula(t: SheetTable, col: TableColumn): string | null {
  const fn = col.totalsFunction;
  if (!fn || fn === 'none') return null;
  if (fn === 'custom') return col.totalsFormula ?? null;
  return `SUBTOTAL(${TOTALS_SUBTOTAL[fn]},${t.name}[${escapeColumnName(col.name)}])`;
}

/** Fills the totals row cells from the column settings. */
function writeTotals(doc: SpreadsheetDocument, tx: Transaction, sheet: Sheet, t: SheetTable): void {
  if (!t.totalsRow) return;
  const r = t.ref.r1;
  t.columns.forEach((col, i) => {
    const c = t.ref.c0 + i;
    const cell = sheet.getCell(r, c);
    const f = totalsFormula(t, col);
    if (f) doc.put(tx, sheet, r, c, { v: null, f, s: cell?.s });
    else if (col.totalsLabel !== undefined) doc.put(tx, sheet, r, c, { v: col.totalsLabel, s: cell?.s });
    else doc.put(tx, sheet, r, c, cell?.s ? { v: null, s: cell.s } : undefined);
  });
}

proto.createTable = function (this: SpreadsheetDocument, range, opts = {}, sheet = this.sheet) {
  if (this.isReadOnly()) return 'The workbook is read-only.';
  const r = makeRange(range.r0, range.c0, range.r1, range.c1);
  if (r.c1 - r.c0 > 16383 || r.r1 - r.r0 > 1048574) return 'The table is too large.';
  if (sheet.tables.some((t) => rangesIntersect(t.ref, r))) return 'A table cannot overlap another table.';
  const hasHeaders = opts.hasHeaders ?? true;
  let created: SheetTable | null = null;
  this.transact('Create table', (tx) => {
    tx.snapshotSheet(sheet);
    let ref = { ...r };
    if (!hasHeaders) {
      // A header row is added above the data.
      if (ref.r0 > 0 && isEmptyRow(sheet, ref.r0 - 1, ref.c0, ref.c1)) ref = { ...ref, r0: ref.r0 - 1 };
      else {
        this.insertRows(ref.r0, 1, sheet);
        ref = { ...ref, r1: ref.r1 + 1 };
      }
    }
    if (ref.r1 === ref.r0) ref = { ...ref, r1: ref.r0 + 1 };
    // Merged cells are not allowed in tables.
    sheet.setMerges(sheet.merges.filter((m) => !rangesIntersect(m, ref)));
    // An auto filter of the area is replaced by the table's own filter.
    if (sheet.autoFilter && rangesIntersect(sheet.autoFilter.ref, ref)) {
      this.unhideFiltered(sheet);
      sheet.autoFilter = undefined;
    }
    const columns: TableColumn[] = [];
    for (let c = ref.c0; c <= ref.c1; c++) {
      const text = hasHeaders ? this.displayText(ref.r0, c, sheet).trim() : '';
      const name = uniqueColumnName(text || `Column${c - ref.c0 + 1}`, columns.map((x) => x.name));
      columns.push({ id: c - ref.c0 + 1, name });
    }
    const table: SheetTable = {
      id: nextTableId(this),
      name: opts.name && !tableNameError(opts.name) && !findTable(this.wb, opts.name) ? opts.name : this.uniqueTableName(),
      ref,
      headerRow: true,
      totalsRow: false,
      columns,
      style: { name: opts.style ?? DEFAULT_TABLE_STYLE, showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false },
      autoFilter: true,
      filters: new Map(),
      autoExpand: true,
    };
    sheet.tables = [...sheet.tables, table];
    writeHeaders(this, tx, sheet, table);
    created = table;
  });
  return created ?? 'The table could not be created.';
};

proto.resizeTable = function (this: SpreadsheetDocument, id, ref) {
  const hit = this.tableById(id);
  if (!hit) return 'The table does not exist.';
  const { sheet, table } = hit;
  const r = makeRange(ref.r0, ref.c0, ref.r1, ref.c1);
  if (r.r0 !== table.ref.r0) return 'The headers must remain in the same row, and the resulting table range must overlap the original table range.';
  if (!rangesIntersect(r, table.ref)) return 'The resulting table range must overlap the original table range.';
  if (sheet.tables.some((t) => t !== table && rangesIntersect(t.ref, r))) return 'A table cannot overlap another table.';
  if (r.r1 <= r.r0 && table.headerRow) return 'A table must have at least one data row.';
  this.transact('Resize table', (tx) => {
    tx.snapshotSheet(sheet);
    const columns: TableColumn[] = [];
    let nextId = Math.max(0, ...table.columns.map((c) => c.id)) + 1;
    for (let c = r.c0; c <= r.c1; c++) {
      const old = c >= table.ref.c0 && c <= table.ref.c1 ? table.columns[c - table.ref.c0] : undefined;
      if (old) columns.push(old);
      else {
        const text = table.headerRow ? this.displayText(r.r0, c, sheet).trim() : '';
        columns.push({ id: nextId++, name: uniqueColumnName(text || `Column${nextId - 1}`, columns.map((x) => x.name)) });
      }
    }
    const filters = new Map([...table.filters].filter(([c]) => c >= r.c0 && c <= r.c1));
    const next: SheetTable = { ...table, ref: r, columns, filters };
    sheet.tables = sheet.tables.map((t) => (t === table ? next : t));
    sheet.setMerges(sheet.merges.filter((m) => !rangesIntersect(m, r)));
    writeHeaders(this, tx, sheet, next);
    this.applyFilterInTx(sheet);
  });
  return null;
};

/** Rewrites formulas of the workbook (cells, names, validations) that refer to the table. */
function rewriteFormulas(doc: SpreadsheetDocument, tx: Transaction, fn: (text: string, table: string) => string | null): void {
  const rewrite = (f: string) => {
    if (!f.includes('[') && !/[A-Za-z_]/.test(f)) return f;
    try {
      return rewriteStructured(f, tokenize(f), fn);
    } catch {
      return f;
    }
  };
  doc.adjustAllFormulas(tx, (f) => rewrite(f));
  for (const n of doc.wb.names) n.formula = rewrite(n.formula);
}

proto.renameTable = function (this: SpreadsheetDocument, id, name) {
  const hit = this.tableById(id);
  if (!hit) return 'The table does not exist.';
  const clean = name.trim();
  const err = tableNameError(clean);
  if (err) return err;
  const other = findTable(this.wb, clean);
  if ((other && other.table !== hit.table) || this.wb.names.some((n) => n.name.toLowerCase() === clean.toLowerCase())) return 'The name is already in use. Enter a unique name.';
  const old = hit.table.name;
  if (old === clean) return null;
  this.transact('Rename table', (tx) => {
    tx.snapshotWorkbook(this.wb);
    tx.snapshotSheet(hit.sheet);
    hit.sheet.tables = hit.sheet.tables.map((t) => (t === hit.table ? { ...t, name: clean } : t));
    rewriteFormulas(this, tx, (text, table) => {
      if (table.toLowerCase() !== old.toLowerCase()) return null;
      return clean + text.slice(table.length);
    });
  });
  return null;
};

/** Cell formatting that reproduces the table style (Convert to Range keeps the look). */
function bakeStyle(doc: SpreadsheetDocument, tx: Transaction, sheet: Sheet, t: SheetTable): void {
  doc.patchRanges(tx, sheet, [t.ref], (r, c, current) => {
    const d = tableCellDecor(t, r, c, doc.wb.theme);
    if (!d) return null;
    const patch: StylePatch = {};
    if (d.fill && !current.fill) patch.fill = { pattern: 'solid', fg: rgbSpec(d.fill)!, bg: { indexed: 64 } };
    if (d.color) patch.font = { color: rgbSpec(d.color)! };
    if (d.bold) patch.font = { ...(patch.font ?? {}), bold: true };
    const b: Partial<Record<'top' | 'bottom' | 'left' | 'right', BorderLine>> = {};
    for (const side of ['top', 'bottom', 'left', 'right'] as const) if (d[side] && !current.border?.[side]) b[side] = d[side]!;
    if (Object.keys(b).length) patch.border = b;
    return Object.keys(patch).length ? patch : null;
  });
}

/** Structured references to the table become cell references. */
function structuredToRanges(doc: SpreadsheetDocument, tx: Transaction, t: SheetTable, home: Sheet): void {
  doc.adjustAllFormulas(tx, (f, sheet) => {
    if (!f.includes('[') && !f.toLowerCase().includes(t.name.toLowerCase())) return f;
    try {
      return rewriteStructured(f, tokenize(f), (text, table) => {
        const bare = text === table;
        if (table.toLowerCase() !== t.name.toLowerCase()) return null;
        const spec = parseStructured(bare ? `${text}[]` : text, table);
        if (!spec) return null;
        const res = resolveStructured(doc.wb, spec, sheet, 0, 0);
        if (!('range' in res)) return null;
        const prefix = res.sheet === sheet ? '' : `${res.sheet.name.includes(' ') ? `'${res.sheet.name}'` : res.sheet.name}!`;
        return prefix + absRange(res.range);
      });
    } catch {
      return f;
    }
  });
  void home;
}

proto.convertTableToRange = function (this: SpreadsheetDocument, id) {
  const hit = this.tableById(id);
  if (!hit) return;
  const { sheet, table } = hit;
  this.transact('Convert to range', (tx) => {
    tx.snapshotSheet(sheet);
    structuredToRanges(this, tx, table, sheet);
    bakeStyle(this, tx, sheet, table);
    sheet.tables = sheet.tables.filter((t) => t !== table);
    this.applyFilterInTx(sheet);
  });
};

proto.deleteTable = function (this: SpreadsheetDocument, id, clearCells = false) {
  const hit = this.tableById(id);
  if (!hit) return;
  const { sheet, table } = hit;
  this.transact('Delete table', (tx) => {
    tx.snapshotSheet(sheet);
    sheet.tables = sheet.tables.filter((t) => t !== table);
    if (clearCells) for (let r = table.ref.r0; r <= table.ref.r1; r++) for (let c = table.ref.c0; c <= table.ref.c1; c++) sheet.setCell(r, c, undefined);
    this.applyFilterInTx(sheet);
  });
};

function replaceTable(sheet: Sheet, old: SheetTable, next: SheetTable): void {
  sheet.tables = sheet.tables.map((t) => (t === old ? next : t));
}

proto.setTableStyle = function (this: SpreadsheetDocument, id, style) {
  const hit = this.tableById(id);
  if (!hit) return;
  this.transact('Table style', (tx) => {
    tx.snapshotSheet(hit.sheet);
    replaceTable(hit.sheet, hit.table, { ...hit.table, style: { ...hit.table.style, name: style } });
  });
};

proto.setTableOptions = function (this: SpreadsheetDocument, id, patch) {
  const hit = this.tableById(id);
  if (!hit) return;
  const { sheet } = hit;
  this.transact('Table options', (tx) => {
    tx.snapshotSheet(sheet);
    let t: SheetTable = { ...hit.table, style: { ...hit.table.style } };
    const s = t.style;
    if (patch.showRowStripes !== undefined) s.showRowStripes = patch.showRowStripes;
    if (patch.showColumnStripes !== undefined) s.showColumnStripes = patch.showColumnStripes;
    if (patch.showFirstColumn !== undefined) s.showFirstColumn = patch.showFirstColumn;
    if (patch.showLastColumn !== undefined) s.showLastColumn = patch.showLastColumn;
    if (patch.autoExpand !== undefined) t.autoExpand = patch.autoExpand;
    if (patch.autoFilter !== undefined && patch.autoFilter !== t.autoFilter) {
      t.autoFilter = patch.autoFilter && t.headerRow;
      if (!t.autoFilter) t.filters = new Map();
    }
    if (patch.totalsRow !== undefined && patch.totalsRow !== t.totalsRow) {
      if (patch.totalsRow) {
        const below = t.ref.r1 + 1;
        if (!isEmptyRow(sheet, below, t.ref.c0, t.ref.c1)) {
          replaceTable(sheet, hit.table, t);
          this.insertRows(below, 1, sheet);
          t = sheet.tables.find((x) => x.id === t.id)!;
        }
        const columns = t.columns.map((col) => ({ ...col }));
        const first = columns[0];
        const last = columns[columns.length - 1];
        if (first && columns.length > 1 && !first.totalsFunction) first.totalsLabel = first.totalsLabel ?? 'Total';
        if (last && !last.totalsFunction) {
          const data = tableDataRange(t);
          let numeric = false;
          for (let r = data.r0; r <= data.r1 && !numeric; r++) if (typeof sheet.getCell(r, t.ref.c1)?.v === 'number') numeric = true;
          last.totalsFunction = numeric ? 'sum' : 'count';
          last.totalsLabel = undefined;
        }
        t = { ...t, columns, totalsRow: true, ref: { ...t.ref, r1: t.ref.r1 + 1 } };
        replaceTable(sheet, sheet.tables.find((x) => x.id === t.id)!, t);
        writeTotals(this, tx, sheet, t);
      } else {
        const r = t.ref.r1;
        for (let c = t.ref.c0; c <= t.ref.c1; c++) {
          const cell = sheet.getCell(r, c);
          sheet.setCell(r, c, cell?.s ? { v: null, s: cell.s } : undefined);
        }
        t = { ...t, totalsRow: false, ref: { ...t.ref, r1: t.ref.r1 - 1 } };
      }
    }
    if (patch.headerRow !== undefined && patch.headerRow !== t.headerRow) {
      if (!patch.headerRow) {
        for (let c = t.ref.c0; c <= t.ref.c1; c++) {
          const cell = sheet.getCell(t.ref.r0, c);
          sheet.setCell(t.ref.r0, c, cell?.s ? { v: null, s: cell.s } : undefined);
        }
        t = { ...t, headerRow: false, autoFilter: false, filters: new Map(), ref: { ...t.ref, r0: t.ref.r0 + 1 } };
      } else {
        const above = t.ref.r0 - 1;
        if (above >= 0 && isEmptyRow(sheet, above, t.ref.c0, t.ref.c1)) t = { ...t, headerRow: true, autoFilter: true, ref: { ...t.ref, r0: above } };
        else {
          replaceTable(sheet, sheet.tables.find((x) => x.id === t.id)!, t);
          this.insertRows(t.ref.r0, 1, sheet);
          const moved = sheet.tables.find((x) => x.id === t.id)!;
          t = { ...moved, headerRow: true, autoFilter: true, ref: { ...moved.ref, r0: moved.ref.r0 - 1 } };
        }
      }
    }
    replaceTable(sheet, sheet.tables.find((x) => x.id === t.id)!, t);
    writeHeaders(this, tx, sheet, t);
    this.applyFilterInTx(sheet);
  });
};

proto.setTotalsFunction = function (this: SpreadsheetDocument, id, colIndex, fn) {
  const hit = this.tableById(id);
  if (!hit || !hit.table.totalsRow) return;
  const { sheet, table } = hit;
  this.transact('Totals', (tx) => {
    const columns = table.columns.map((c, i) => (i === colIndex ? { ...c, totalsFunction: fn, totalsLabel: fn === 'none' ? undefined : c.totalsLabel } : c));
    tx.snapshotSheet(sheet);
    const next = { ...table, columns };
    replaceTable(sheet, table, next);
    writeTotals(this, tx, sheet, next);
  });
};

proto.setTableFilter = function (this: SpreadsheetDocument, id, col, filter) {
  const hit = this.tableById(id);
  if (!hit) return;
  const { sheet, table } = hit;
  this.transact('Filter', (tx) => {
    tx.snapshotSheet(sheet);
    const filters = new Map(table.filters);
    if (filter && (filter.values || filter.conditions?.length)) filters.set(col, filter);
    else filters.delete(col);
    replaceTable(sheet, table, { ...table, filters });
    this.applyFilterInTx(sheet);
  });
};

proto.clearTableFilters = function (this: SpreadsheetDocument, id) {
  const hit = this.tableById(id);
  if (!hit || !hit.table.filters.size) return;
  this.transact('Clear filters', (tx) => {
    tx.snapshotSheet(hit.sheet);
    replaceTable(hit.sheet, hit.table, { ...hit.table, filters: new Map() });
    this.applyFilterInTx(hit.sheet);
  });
};

proto.sortTable = function (this: SpreadsheetDocument, id, col, ascending) {
  const hit = this.tableById(id);
  if (!hit) return;
  const { sheet, table } = hit;
  const area = { ...table.ref, r1: table.ref.r1 - (table.totalsRow ? 1 : 0) };
  this.transact('Sort', (tx) => {
    this.sortRange(area, col, ascending, table.headerRow, sheet);
    tx.snapshotSheet(sheet);
    this.applyFilterInTx(sheet);
  });
};

// -----------------------------------------------------------------------------------------------
// Typing in and next to tables

function headerEdited(doc: SpreadsheetDocument, tx: Transaction, sheet: Sheet, t: SheetTable, c: number): void {
  const i = c - t.ref.c0;
  const col = t.columns[i];
  const text = doc.displayText(t.ref.r0, c, sheet).trim();
  const others = t.columns.filter((_x, k) => k !== i).map((x) => x.name);
  const name = uniqueColumnName(text || `Column${i + 1}`, others);
  const cell = sheet.getCell(t.ref.r0, c);
  if (cell?.v !== name || cell.f !== undefined) doc.put(tx, sheet, t.ref.r0, c, { v: name, s: cell?.s });
  if (name === col.name) return;
  const old = col.name;
  tx.snapshotSheet(sheet);
  const next = { ...t, columns: t.columns.map((x, k) => (k === i ? { ...x, name } : x)) };
  replaceTable(sheet, t, next);
  // Structured references follow the new column name.
  rewriteFormulas(doc, tx, (text2, table) => {
    const own = table ? table.toLowerCase() === t.name.toLowerCase() : false;
    if (!own && table) return null;
    if (!text2.includes('[')) return null;
    const esc = escapeColumnName(old).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return text2.replace(new RegExp(`\\[(@?)${esc}\\]`, 'gi'), (_m, at: string) => `[${at}${escapeColumnName(name)}]`).replace(new RegExp(`@${esc}(?=\\])`, 'gi'), `@${escapeColumnName(name)}`);
  });
}

/** Excel's calculated columns: a formula typed into an otherwise empty column fills the column. */
function calculatedColumn(doc: SpreadsheetDocument, tx: Transaction, sheet: Sheet, t: SheetTable, r: number, c: number): void {
  const cell = sheet.getCell(r, c);
  if (cell?.f === undefined) return;
  const data = tableDataRange(t);
  if (data.r1 <= data.r0) return;
  for (let rr = data.r0; rr <= data.r1; rr++) {
    if (rr === r) continue;
    const other = sheet.getCell(rr, c);
    if (other && (other.v !== null || other.f !== undefined)) return;
  }
  for (let rr = data.r0; rr <= data.r1; rr++) {
    if (rr === r) continue;
    const other = sheet.getCell(rr, c);
    const f = translateFormula(cell.f, rr - r, 0);
    const next: Cell = { v: null, f, s: other?.s ?? cell.s };
    if (cell.dyn) next.dyn = true;
    doc.put(tx, sheet, rr, c, next);
  }
  const i = c - t.ref.c0;
  tx.snapshotSheet(sheet);
  replaceTable(sheet, t, { ...t, columns: t.columns.map((x, k) => (k === i ? { ...x, calculatedFormula: cell.f } : x)) });
}

function autoExpand(doc: SpreadsheetDocument, tx: Transaction, sheet: Sheet, r: number, c: number): boolean {
  const cell = sheet.getCell(r, c);
  if (!cell || (cell.v === null && cell.f === undefined)) return false;
  for (const t of sheet.tables) {
    if (t.autoExpand === false) continue;
    // A new row right below the data (no totals row).
    if (!t.totalsRow && r === t.ref.r1 + 1 && c >= t.ref.c0 && c <= t.ref.c1) {
      tx.snapshotSheet(sheet);
      const next = { ...t, ref: { ...t.ref, r1: r } };
      replaceTable(sheet, t, next);
      // Calculated columns continue into the new row.
      next.columns.forEach((col, i) => {
        const cc = t.ref.c0 + i;
        if (!col.calculatedFormula || cc === c || sheet.getCell(r, cc)?.f !== undefined) return;
        const above = sheet.getCell(r - 1, cc);
        if (above?.f !== undefined) doc.put(tx, sheet, r, cc, { v: null, f: translateFormula(above.f, 1, 0), s: above.s, dyn: above.dyn });
      });
      return true;
    }
    // A new column right of the table.
    if (c === t.ref.c1 + 1 && r >= t.ref.r0 && r <= t.ref.r1 && !sheet.tables.some((o) => o !== t && rangesIntersect(o.ref, makeRange(t.ref.r0, c, t.ref.r1, c)))) {
      tx.snapshotSheet(sheet);
      const nextId = Math.max(0, ...t.columns.map((x) => x.id)) + 1;
      const headerText = t.headerRow && r === t.ref.r0 ? doc.displayText(r, c, sheet).trim() : '';
      const name = uniqueColumnName(headerText || `Column${nextId}`, t.columns.map((x) => x.name));
      const next = { ...t, ref: { ...t.ref, c1: c }, columns: [...t.columns, { id: nextId, name }] };
      replaceTable(sheet, t, next);
      writeHeaders(doc, tx, sheet, next);
      return true;
    }
  }
  return false;
}

SpreadsheetDocument.inputHooks.push((doc, tx, sheet, r, c) => {
  if (!sheet.tables.length) return;
  const t = sheet.tables.find((x) => r >= x.ref.r0 && r <= x.ref.r1 && c >= x.ref.c0 && c <= x.ref.c1);
  if (t) {
    if (t.headerRow && r === t.ref.r0) headerEdited(doc, tx, sheet, t, c);
    else if (!(t.totalsRow && r === t.ref.r1)) calculatedColumn(doc, tx, sheet, t, r, c);
    return;
  }
  autoExpand(doc, tx, sheet, r, c);
});

export {};
