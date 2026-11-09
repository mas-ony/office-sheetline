/** Go To Special: selects cells by kind (formulas, constants, blanks, comments, rules...). */
import { SpreadsheetDocument } from '../document';
import { MAX_COLS, makeRange, rangeContains, rangesIntersect, type Range } from '../refs';
import { filterRegion } from '../filter';
import { CellError } from '../values';
import type { Cell } from '../model/sheet';

export type SpecialKind =
  | 'comments' | 'constants' | 'formulas' | 'blanks' | 'currentRegion' | 'currentArray' | 'objects' | 'rowDifferences'
  | 'columnDifferences' | 'lastCell' | 'visible' | 'condFormats' | 'validation';

export interface SpecialOptions {
  /** For constants and formulas: which value types. */
  numbers?: boolean;
  text?: boolean;
  logicals?: boolean;
  errors?: boolean;
}

declare module '../document' {
  interface SpreadsheetDocument {
    /** Selects the cells of a kind inside the selection (or the used range); returns how many ranges. */
    selectSpecial(kind: SpecialKind, opts?: SpecialOptions): number;
  }
}

/** Merges a set of cells into row runs. */
function runs(cells: { r: number; c: number }[]): Range[] {
  const byRow = new Map<number, number[]>();
  for (const { r, c } of cells) {
    const list = byRow.get(r) ?? [];
    list.push(c);
    byRow.set(r, list);
  }
  const out: Range[] = [];
  for (const [r, cols] of [...byRow.entries()].sort((a, b) => a[0] - b[0])) {
    cols.sort((a, b) => a - b);
    let start = cols[0];
    let prev = cols[0];
    for (const c of cols.slice(1)) {
      if (c === prev + 1) {
        prev = c;
        continue;
      }
      out.push(makeRange(r, start, r, prev));
      start = prev = c;
    }
    out.push(makeRange(r, start, r, prev));
  }
  // Join identical column runs of consecutive rows.
  const joined: Range[] = [];
  for (const rg of out) {
    const last = joined[joined.length - 1];
    if (last && last.c0 === rg.c0 && last.c1 === rg.c1 && last.r1 + 1 === rg.r0) last.r1 = rg.r1;
    else joined.push({ ...rg });
  }
  return joined;
}

function typeMatches(cell: Cell, o: SpecialOptions): boolean {
  const v = cell.v;
  if (typeof v === 'number') return o.numbers !== false;
  if (typeof v === 'string') return o.text !== false;
  if (typeof v === 'boolean') return o.logicals !== false;
  if (v instanceof CellError) return o.errors !== false;
  return false;
}

SpreadsheetDocument.prototype.selectSpecial = function (this: SpreadsheetDocument, kind, opts = {}) {
  const sheet = this.sheet;
  const sel = this.selection;
  const single = sel.ranges.length === 1 && sel.ranges[0].r0 === sel.ranges[0].r1 && sel.ranges[0].c0 === sel.ranges[0].c1;
  const used = sheet.usedRange() ?? makeRange(0, 0);
  const scope: Range[] = single ? [used] : sel.ranges.map((r) => this.bounded(r, sheet) ?? r);
  const inScope = (r: number, c: number) => scope.some((s) => rangeContains(s, r, c));
  let ranges: Range[] = [];
  switch (kind) {
    case 'currentRegion': {
      const reg = filterRegion(sheet, makeRange(sel.active.r, sel.active.c), sel.active);
      ranges = [reg ?? makeRange(sel.active.r, sel.active.c)];
      break;
    }
    case 'lastCell': {
      const ext = sheet.dataExtent();
      ranges = [makeRange(Math.max(0, ext.rows), Math.max(0, ext.cols))];
      break;
    }
    case 'currentArray': {
      const a = this.anchorOf(sel.active.r, sel.active.c) ?? sel.active;
      const spill = this.engine.spillOf(sheet, a.r, a.c);
      ranges = spill ? [spill] : [];
      break;
    }
    case 'comments':
      ranges = runs(sheet.comments.filter((t) => inScope(t.r, t.c)).map((t) => ({ r: t.r, c: t.c })));
      break;
    case 'condFormats':
      ranges = sheet.condFormats.flatMap((cf) => cf.sqref).filter((r) => scope.some((s) => rangesIntersect(r, s)));
      break;
    case 'validation':
      ranges = sheet.validations.flatMap((v) => v.sqref).filter((r) => scope.some((s) => rangesIntersect(r, s)));
      break;
    case 'blanks': {
      const cells: { r: number; c: number }[] = [];
      for (const s of scope) {
        if ((s.r1 - s.r0 + 1) * (s.c1 - s.c0 + 1) > 500000) continue;
        for (let r = s.r0; r <= s.r1; r++) for (let c = s.c0; c <= s.c1; c++) {
          const cell = sheet.getCell(r, c);
          if (!cell || (cell.v === null && cell.f === undefined)) cells.push({ r, c });
        }
      }
      ranges = runs(cells);
      break;
    }
    case 'visible':
      ranges = scope.flatMap((s) => {
        const out: Range[] = [];
        let start = -1;
        for (let r = s.r0; r <= s.r1 + 1; r++) {
          const vis = r <= s.r1 && !sheet.isRowHidden(r);
          if (vis && start < 0) start = r;
          if (!vis && start >= 0) {
            out.push(makeRange(start, s.c0, r - 1, Math.min(s.c1, MAX_COLS - 1)));
            start = -1;
          }
        }
        return out;
      });
      break;
    default: {
      const cells: { r: number; c: number }[] = [];
      for (const s of scope) {
        sheet.forEachCellIn(s, (r, c, cell) => {
          if (kind === 'formulas' && cell.f !== undefined && typeMatches(cell, opts)) cells.push({ r, c });
          else if (kind === 'constants' && cell.f === undefined && cell.spill === undefined && cell.v !== null && typeMatches(cell, opts)) cells.push({ r, c });
        });
      }
      ranges = runs(cells);
    }
  }
  if (!ranges.length) return 0;
  const first = ranges[0];
  this.setSelection({ active: { r: first.r0, c: first.c0 }, anchor: { r: first.r0, c: first.c0 }, ranges });
  return ranges.length;
};

export {};
