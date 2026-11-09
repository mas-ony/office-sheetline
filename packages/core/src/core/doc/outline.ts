/**
 * Data > Outline: grouping rows and columns (outline levels 1-7), collapsing and expanding groups,
 * showing an outline level, Auto Outline from summary formulas, and Subtotal.
 */
import { SpreadsheetDocument } from '../document';
import type { Sheet } from '../model/sheet';
import { MAX_OUTLINE_LEVEL, levelOf, outlineGroups, outlined, type OutlineAxis, type OutlineGroup } from '../model/outline';
import type { Transaction } from '../history';
import { colName, makeRange, type Range } from '../refs';
import { collectRefs, parseFormula } from '../formula/parser';
import type { RefParts } from '../formula/tokenizer';

export const SUBTOTAL_FUNCTIONS: { label: string; fn: string; code: number }[] = [
  { label: 'Sum', fn: 'SUM', code: 9 },
  { label: 'Count', fn: 'COUNTA', code: 3 },
  { label: 'Average', fn: 'AVERAGE', code: 1 },
  { label: 'Max', fn: 'MAX', code: 4 },
  { label: 'Min', fn: 'MIN', code: 5 },
  { label: 'Product', fn: 'PRODUCT', code: 6 },
  { label: 'Count Numbers', fn: 'COUNT', code: 2 },
  { label: 'StdDev', fn: 'STDEV', code: 7 },
  { label: 'StdDevp', fn: 'STDEVP', code: 8 },
  { label: 'Var', fn: 'VAR', code: 10 },
  { label: 'Varp', fn: 'VARP', code: 11 },
];

export interface SubtotalOptions {
  /** Column whose changes start a new group. */
  groupBy: number;
  /** SUBTOTAL function number (9 = SUM). */
  code: number;
  /** Columns that get a subtotal. */
  columns: number[];
  /** Remove existing subtotals first. */
  replace: boolean;
  /** Page break between groups. */
  pageBreaks: boolean;
  /** Summary rows below the data (else above). */
  summaryBelow: boolean;
}

export { maxOutlineLevel, outlineGroups, type OutlineAxis, type OutlineGroup } from '../model/outline';

declare module '../document' {
  interface SpreadsheetDocument {
    groupLines(axis: OutlineAxis, from: number, to: number, sheet?: Sheet): string | null;
    ungroupLines(axis: OutlineAxis, from: number, to: number, sheet?: Sheet): boolean;
    clearOutline(axis?: OutlineAxis, sheet?: Sheet): void;
    /** Groups the rows and columns that summary formulas (SUM of the lines above or to the left) cover. */
    autoOutline(sheet?: Sheet): boolean;
    /** Collapses or expands a group (the +/- button). */
    setGroupCollapsed(group: OutlineGroup, collapsed: boolean, sheet?: Sheet): void;
    /** The level buttons: shows lines up to an outline level. */
    showOutlineLevel(axis: OutlineAxis, level: number, sheet?: Sheet): void;
    /** Show Detail / Hide Detail for the group at the active cell. */
    showDetail(show: boolean): boolean;
    setOutlineSettings(settings: { summaryBelow?: boolean; summaryRight?: boolean }, sheet?: Sheet): void;
    applySubtotals(region: Range, opts: SubtotalOptions, sheet?: Sheet): string | null;
    removeSubtotals(region: Range, sheet?: Sheet): void;
  }
}

const proto = SpreadsheetDocument.prototype;

function setInfo(tx: Transaction, sheet: Sheet, axis: OutlineAxis, i: number, patch: { outline?: number; hidden?: boolean; collapsed?: boolean }): void {
  if (axis === 'row') {
    tx.touchRow(sheet, i);
    const info = { ...(sheet.rows.get(i) ?? {}), ...patch };
    if (!info.outline) delete info.outline;
    if (!info.hidden) delete info.hidden;
    if (!info.collapsed) delete info.collapsed;
    sheet.setRowInfo(i, info);
  } else {
    tx.touchCol(sheet, i);
    const info = { ...(sheet.cols.get(i) ?? {}), ...patch };
    if (!info.outline) delete info.outline;
    if (!info.hidden) delete info.hidden;
    if (!info.collapsed) delete info.collapsed;
    sheet.setColInfo(i, info);
  }
}

proto.groupLines = function (this: SpreadsheetDocument, axis, from, to, sheet = this.sheet) {
  for (let i = from; i <= to; i++) if (levelOf(sheet, axis, i) >= MAX_OUTLINE_LEVEL) return 'The outline cannot have more than seven levels.';
  this.transact('Group', (tx) => {
    for (let i = from; i <= to; i++) setInfo(tx, sheet, axis, i, { outline: levelOf(sheet, axis, i) + 1 });
  });
  return null;
};

proto.ungroupLines = function (this: SpreadsheetDocument, axis, from, to, sheet = this.sheet) {
  let changed = false;
  this.transact('Ungroup', (tx) => {
    for (let i = from; i <= to; i++) {
      const level = levelOf(sheet, axis, i);
      if (!level) continue;
      changed = true;
      // Lines leaving the outline become visible again.
      setInfo(tx, sheet, axis, i, level === 1 ? { outline: 0, hidden: false } : { outline: level - 1 });
    }
  });
  return changed;
};

proto.clearOutline = function (this: SpreadsheetDocument, axis, sheet = this.sheet) {
  this.transact('Clear Outline', (tx) => {
    for (const a of axis ? [axis] : (['row', 'col'] as OutlineAxis[])) {
      for (const g of outlineGroups(sheet, a)) if (g.collapsed) setInfo(tx, sheet, a, g.summary, { collapsed: false });
      for (const i of outlined(sheet, a)) setInfo(tx, sheet, a, i, { outline: 0, hidden: false });
    }
  });
};

proto.setGroupCollapsed = function (this: SpreadsheetDocument, group, collapsed, sheet = this.sheet) {
  const { axis } = group;
  // Nested groups that stay collapsed keep their lines hidden when an outer group expands.
  const nested = collapsed ? [] : outlineGroups(sheet, axis).filter((g) => g.level > group.level && g.start >= group.start && g.end <= group.end && g.collapsed);
  this.transact(collapsed ? 'Hide Detail' : 'Show Detail', (tx) => {
    for (let i = group.start; i <= group.end; i++) {
      const keep = nested.some((g) => i >= g.start && i <= g.end);
      setInfo(tx, sheet, axis, i, { hidden: collapsed || keep });
    }
    setInfo(tx, sheet, axis, group.summary, { collapsed });
  });
};

proto.showOutlineLevel = function (this: SpreadsheetDocument, axis, level, sheet = this.sheet) {
  const groups = outlineGroups(sheet, axis);
  this.transact('Outline level', (tx) => {
    for (const i of outlined(sheet, axis)) setInfo(tx, sheet, axis, i, { hidden: levelOf(sheet, axis, i) >= level });
    for (const g of groups) setInfo(tx, sheet, axis, g.summary, { collapsed: g.level >= level });
  });
};

proto.showDetail = function (this: SpreadsheetDocument, show) {
  const sheet = this.sheet;
  const { r, c } = this.selection.active;
  for (const axis of ['row', 'col'] as OutlineAxis[]) {
    const i = axis === 'row' ? r : c;
    const groups = outlineGroups(sheet, axis);
    // The innermost group whose summary line or body holds the active cell.
    const hits = groups.filter((g) => g.summary === i || (i >= g.start && i <= g.end)).sort((a, b) => b.level - a.level);
    const target = show ? hits.find((g) => g.collapsed && g.summary === i) ?? hits.find((g) => g.collapsed) : hits.find((g) => !g.collapsed);
    if (target) {
      this.setGroupCollapsed(target, !show, sheet);
      return true;
    }
  }
  return false;
};

proto.setOutlineSettings = function (this: SpreadsheetDocument, settings, sheet = this.sheet) {
  this.transact('Outline settings', (tx) => {
    tx.snapshotSheet(sheet);
    if (settings.summaryBelow !== undefined) sheet.outlineBelow = settings.summaryBelow;
    if (settings.summaryRight !== undefined) sheet.outlineRight = settings.summaryRight;
  });
};

/** Ranges a summary formula adds up: SUM/SUBTOTAL arguments in the same column above (or row to the left). */
function summarySpans(text: string, r: number, c: number): { axis: OutlineAxis; from: number; to: number }[] {
  if (!/^\s*(SUM|SUBTOTAL)\s*\(/i.test(text)) return [];
  let refs: RefParts[];
  try {
    refs = collectRefs(parseFormula(text)).refs;
  } catch {
    return [];
  }
  const out: { axis: OutlineAxis; from: number; to: number }[] = [];
  for (const ref of refs) {
    if (ref.sheet || ref.kind !== 'range') continue;
    if (ref.c0 === c && ref.c1 === c && ref.r1 < r) out.push({ axis: 'row', from: ref.r0, to: ref.r1 });
    else if (ref.r0 === r && ref.r1 === r && ref.c1 < c) out.push({ axis: 'col', from: ref.c0, to: ref.c1 });
  }
  return out;
}

proto.autoOutline = function (this: SpreadsheetDocument, sheet = this.sheet) {
  const spans = new Map<string, { axis: OutlineAxis; from: number; to: number }>();
  for (const [r, row] of sheet.cells) {
    for (const [c, cell] of row) {
      if (!cell.f) continue;
      for (const s of summarySpans(cell.f, r, c)) {
        // Only blocks right next to their summary line.
        const next = s.axis === 'row' ? r : c;
        if (s.to !== next - 1) continue;
        spans.set(`${s.axis}:${s.from}:${s.to}`, s);
      }
    }
  }
  if (!spans.size) return false;
  // Larger spans first, so nested blocks end up one level deeper.
  const list = [...spans.values()].sort((a, b) => b.to - b.from - (a.to - a.from));
  this.transact('Auto Outline', (tx) => {
    for (const a of ['row', 'col'] as OutlineAxis[]) for (const i of outlined(sheet, a)) setInfo(tx, sheet, a, i, { outline: 0, hidden: false, collapsed: false });
    for (const s of list) {
      for (let i = s.from; i <= s.to; i++) {
        const level = levelOf(sheet, s.axis, i);
        if (level < MAX_OUTLINE_LEVEL) setInfo(tx, sheet, s.axis, i, { outline: level + 1 });
      }
    }
  });
  return true;
};

const SUBTOTAL_RE = /^\s*SUBTOTAL\s*\(/i;

proto.removeSubtotals = function (this: SpreadsheetDocument, region, sheet = this.sheet) {
  const rows: number[] = [];
  for (let r = region.r0; r <= region.r1; r++) {
    const row = sheet.cells.get(r);
    if (row && [...row.values()].some((cell) => cell.f && SUBTOTAL_RE.test(cell.f))) rows.push(r);
  }
  this.transact('Remove Subtotals', () => {
    for (const r of rows.reverse()) this.deleteRows(r, 1, sheet);
    this.clearOutline('row', sheet);
  });
};

proto.applySubtotals = function (this: SpreadsheetDocument, region, opts, sheet = this.sheet) {
  if (sheet.tables.some((t) => t.ref.r0 <= region.r1 && t.ref.r1 >= region.r0 && t.ref.c0 <= region.c1 && t.ref.c1 >= region.c0)) return 'Subtotals are not supported in tables.';
  if (region.r1 - region.r0 < 1) return 'Select a range with a header row and data.';
  if (!opts.columns.length) return 'Select at least one column to add a subtotal to.';
  const fn = SUBTOTAL_FUNCTIONS.find((f) => f.code === opts.code) ?? SUBTOTAL_FUNCTIONS[0];
  let error: string | null = null;
  this.transact('Subtotal', (tx) => {
    let r1 = region.r1;
    if (opts.replace) {
      for (let r = region.r1; r > region.r0; r--) {
        const row = sheet.cells.get(r);
        if (row && [...row.values()].some((cell) => cell.f && SUBTOTAL_RE.test(cell.f))) {
          this.deleteRows(r, 1, sheet);
          r1--;
        }
      }
      for (const i of outlined(sheet, 'row')) setInfo(tx, sheet, 'row', i, { outline: 0, hidden: false, collapsed: false });
    }
    // Group boundaries by the value of the "at each change in" column.
    const first = region.r0 + 1;
    const key = (r: number) => this.displayText(r, opts.groupBy, sheet);
    const groups: { from: number; to: number; label: string }[] = [];
    for (let r = first; r <= r1; r++) {
      const k = key(r);
      const last = groups[groups.length - 1];
      if (last && last.label === k) last.to = r;
      else groups.push({ from: r, to: r, label: k });
    }
    if (!groups.length) {
      error = 'There is no data under the header row.';
      return;
    }
    const labelCol = opts.groupBy;
    const formula = (c: number, from: number, to: number) => `SUBTOTAL(${fn.code},${colName(c)}${from + 1}:${colName(c)}${to + 1})`;
    let offset = 0;
    for (const g of groups) {
      const from = g.from + offset;
      const to = g.to + offset;
      const at = opts.summaryBelow ? to + 1 : from;
      this.insertRows(at, 1, sheet);
      offset++;
      const dataFrom = opts.summaryBelow ? from : from + 1;
      const dataTo = opts.summaryBelow ? to : to + 1;
      this.put(tx, sheet, at, labelCol, { v: `${g.label} ${fn.label === 'Sum' ? 'Total' : fn.label}`, s: sheet.getCell(dataFrom, labelCol)?.s, f: undefined });
      for (const c of opts.columns) this.put(tx, sheet, at, c, { v: null, f: formula(c, dataFrom, dataTo), s: sheet.getCell(dataFrom, c)?.s });
      for (let r = dataFrom; r <= dataTo; r++) setInfo(tx, sheet, 'row', r, { outline: 2 });
      setInfo(tx, sheet, 'row', at, { outline: 1 });
      if (opts.pageBreaks && g !== groups[groups.length - 1]) {
        const breakAt = opts.summaryBelow ? at + 1 : dataTo + 1;
        tx.snapshotSheet(sheet);
        sheet.page.rowBreaks = [...new Set([...(sheet.page.rowBreaks ?? []), breakAt])].sort((a, b) => a - b);
      }
    }
    const end = r1 + offset;
    const grandAt = opts.summaryBelow ? end + 1 : region.r0 + 1;
    this.insertRows(grandAt, 1, sheet);
    const dataFrom = opts.summaryBelow ? first : first + 1;
    const dataTo = opts.summaryBelow ? end : end + 1;
    this.put(tx, sheet, grandAt, labelCol, { v: `Grand ${fn.label === 'Sum' ? 'Total' : fn.label}`, s: sheet.getCell(dataFrom, labelCol)?.s });
    for (const c of opts.columns) this.put(tx, sheet, grandAt, c, { v: null, f: formula(c, dataFrom, dataTo), s: sheet.getCell(dataFrom, c)?.s });
    tx.snapshotSheet(sheet);
    sheet.outlineBelow = opts.summaryBelow;
  });
  if (!error) this.select(makeRange(region.r0, region.c0, region.r0, region.c1));
  return error;
};
