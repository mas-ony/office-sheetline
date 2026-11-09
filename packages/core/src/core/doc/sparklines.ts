/** Insert > Sparklines and the Sparkline contextual tab: creating, editing, grouping and clearing. */
import { SpreadsheetDocument } from '../document';
import type { Sheet } from '../model/sheet';
import { cloneSparklineGroup, newSparklineGroup, type Sparkline, type SparklineGroup, type SparklineType } from '../model/sparkline';
import { absRef } from '../chartData';
import type { Range } from '../refs';

declare module '../document' {
  interface SpreadsheetDocument {
    /** Creates a group: one sparkline per row (or column) of `data`, placed in `location`. */
    addSparklines(type: SparklineType, data: { sheet: Sheet; range: Range }, location: Range, sheet?: Sheet): string | null;
    sparklineAt(r: number, c: number, sheet?: Sheet): { group: SparklineGroup; sparkline: Sparkline } | null;
    updateSparklineGroup(id: string, patch: Partial<SparklineGroup>, sheet?: Sheet): void;
    /** Removes the sparklines in the ranges (or whole groups touching them). */
    clearSparklines(ranges: Range[], wholeGroups: boolean, sheet?: Sheet): void;
    /** Merges the groups of the sparklines in the ranges into one (formatting of the active one). */
    groupSparklines(ranges: Range[], sheet?: Sheet): void;
    ungroupSparklines(ranges: Range[], sheet?: Sheet): void;
  }
}

const proto = SpreadsheetDocument.prototype;

const inRanges = (ranges: Range[], r: number, c: number) => ranges.some((x) => r >= x.r0 && r <= x.r1 && c >= x.c0 && c <= x.c1);

proto.addSparklines = function (this: SpreadsheetDocument, type, data, location, sheet = this.sheet) {
  const dr = data.range;
  const rows = dr.r1 - dr.r0 + 1;
  const cols = dr.c1 - dr.c0 + 1;
  const cells: { r: number; c: number }[] = [];
  for (let r = location.r0; r <= location.r1; r++) for (let c = location.c0; c <= location.c1; c++) cells.push({ r, c });
  if (location.r0 !== location.r1 && location.c0 !== location.c1) return 'Location reference is not valid because the cells are not all in the same column or row. Select cells that are all in a single row or column.';
  let sparklines: Sparkline[];
  if (cells.length === 1) {
    if (rows > 1 && cols > 1) return 'Data Range must be a single row or a single column when the location is a single cell.';
    sparklines = [{ ...cells[0], ref: absRef(data.sheet, dr) }];
  } else if (location.c0 === location.c1 && rows === cells.length) {
    // A column of sparklines: one per data row.
    sparklines = cells.map((cell, i) => ({ ...cell, ref: absRef(data.sheet, { r0: dr.r0 + i, r1: dr.r0 + i, c0: dr.c0, c1: dr.c1 }) }));
  } else if (location.r0 === location.r1 && cols === cells.length) {
    sparklines = cells.map((cell, i) => ({ ...cell, ref: absRef(data.sheet, { r0: dr.r0, r1: dr.r1, c0: dr.c0 + i, c1: dr.c0 + i }) }));
  } else if (rows === cells.length) {
    sparklines = cells.map((cell, i) => ({ ...cell, ref: absRef(data.sheet, { r0: dr.r0 + i, r1: dr.r0 + i, c0: dr.c0, c1: dr.c1 }) }));
  } else if (cols === cells.length) {
    sparklines = cells.map((cell, i) => ({ ...cell, ref: absRef(data.sheet, { r0: dr.r0, r1: dr.r1, c0: dr.c0 + i, c1: dr.c0 + i }) }));
  } else return 'Location reference is not valid because the cells do not match the number of rows or columns in the data range.';
  const group = newSparklineGroup(type, sparklines);
  if (type === 'stacked') group.negative = true;
  this.transact('Insert sparklines', (tx) => {
    tx.snapshotSheet(sheet);
    // A cell holds one sparkline: new ones replace older ones.
    const taken = new Set(sparklines.map((s) => `${s.r}:${s.c}`));
    sheet.sparklineGroups = [...sheet.sparklineGroups.map((g) => ({ ...g, sparklines: g.sparklines.filter((s) => !taken.has(`${s.r}:${s.c}`)) })).filter((g) => g.sparklines.length), group];
  });
  return null;
};

proto.sparklineAt = function (this: SpreadsheetDocument, r, c, sheet = this.sheet) {
  for (const group of sheet.sparklineGroups) {
    const sparkline = group.sparklines.find((s) => s.r === r && s.c === c);
    if (sparkline) return { group, sparkline };
  }
  return null;
};

proto.updateSparklineGroup = function (this: SpreadsheetDocument, id, patch, sheet = this.sheet) {
  if (!sheet.sparklineGroups.some((g) => g.id === id)) return;
  this.transact('Sparkline', (tx) => {
    tx.snapshotSheet(sheet);
    sheet.sparklineGroups = sheet.sparklineGroups.map((g) => (g.id === id ? { ...cloneSparklineGroup(g), ...patch, colors: { ...g.colors, ...(patch.colors ?? {}) } } : g));
  });
};

proto.clearSparklines = function (this: SpreadsheetDocument, ranges, wholeGroups, sheet = this.sheet) {
  this.transact('Clear sparklines', (tx) => {
    tx.snapshotSheet(sheet);
    sheet.sparklineGroups = sheet.sparklineGroups
      .filter((g) => !(wholeGroups && g.sparklines.some((s) => inRanges(ranges, s.r, s.c))))
      .map((g) => ({ ...g, sparklines: g.sparklines.filter((s) => !inRanges(ranges, s.r, s.c)) }))
      .filter((g) => g.sparklines.length);
  });
};

proto.groupSparklines = function (this: SpreadsheetDocument, ranges, sheet = this.sheet) {
  const touched = sheet.sparklineGroups.filter((g) => g.sparklines.some((s) => inRanges(ranges, s.r, s.c)));
  if (touched.length < 1) return;
  const { r, c } = this.selection.active;
  const lead = touched.find((g) => g.sparklines.some((s) => s.r === r && s.c === c)) ?? touched[0];
  const moved = touched.flatMap((g) => g.sparklines.filter((s) => inRanges(ranges, s.r, s.c)));
  this.transact('Group sparklines', (tx) => {
    tx.snapshotSheet(sheet);
    const rest = sheet.sparklineGroups.map((g) => ({ ...g, sparklines: g.sparklines.filter((s) => !inRanges(ranges, s.r, s.c)) })).filter((g) => g.sparklines.length);
    sheet.sparklineGroups = [...rest, { ...cloneSparklineGroup(lead), id: newSparklineGroup(lead.type, []).id, sparklines: moved.map((s) => ({ ...s })) }];
  });
};

proto.ungroupSparklines = function (this: SpreadsheetDocument, ranges, sheet = this.sheet) {
  this.transact('Ungroup sparklines', (tx) => {
    tx.snapshotSheet(sheet);
    sheet.sparklineGroups = sheet.sparklineGroups.flatMap((g) => {
      if (!g.sparklines.some((s) => inRanges(ranges, s.r, s.c))) return [g];
      return g.sparklines.map((s) => ({ ...cloneSparklineGroup(g), id: newSparklineGroup(g.type, []).id, sparklines: [{ ...s }] }));
    });
  });
};
