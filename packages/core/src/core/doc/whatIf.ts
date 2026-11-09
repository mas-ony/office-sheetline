/**
 * Data > What-If Analysis: Goal Seek, Data Table (one or two input cells) and the Scenario Manager
 * (named sets of input values, Show, Summary).
 */
import { SpreadsheetDocument } from '../document';
import type { Cell, Sheet } from '../model/sheet';
import type { SheetScenarios } from '../model/scenario';
import { cellName, type Range } from '../refs';
import { parseInput } from '../format/input';
import type { Scalar } from '../values';

export type { Scenario, SheetScenarios } from '../model/scenario';

export interface GoalSeekResult {
  found: boolean;
  /** Value found for the changing cell. */
  value: number;
  /** Value the formula cell reached. */
  result: number;
  iterations: number;
}

declare module '../document' {
  interface SpreadsheetDocument {
    /** Finds the input value that makes a formula reach a target (the input cell is set when found). */
    goalSeek(set: { r: number; c: number }, target: number, by: { r: number; c: number }, sheet?: Sheet): GoalSeekResult | string;
    /** Fills a data table (results are computed once, as values). */
    dataTable(range: Range, rowInput: { r: number; c: number } | null, colInput: { r: number; c: number } | null, sheet?: Sheet): string | null;
    setScenarios(next: SheetScenarios | undefined, sheet?: Sheet): void;
    showScenario(index: number, sheet?: Sheet): void;
    /** Adds a "Scenario Summary" sheet; returns its index. */
    scenarioSummary(resultCells: { r: number; c: number }[], sheet?: Sheet): number;
  }
}

const proto = SpreadsheetDocument.prototype;

/** Evaluates `read` with temporary values in input cells, restoring them afterwards (no undo record). */
function withInputs<T>(doc: SpreadsheetDocument, sheet: Sheet, inputs: { r: number; c: number; v: Scalar }[], read: () => T): T {
  const saved: { r: number; c: number; cell: Cell | undefined }[] = inputs.map((i) => ({ r: i.r, c: i.c, cell: sheet.getCell(i.r, i.c) }));
  try {
    for (const i of inputs) {
      const old = sheet.getCell(i.r, i.c);
      sheet.setCell(i.r, i.c, { v: i.v, s: old?.s });
      doc.engine.cellChanged(sheet, i.r, i.c);
    }
    doc.engine.recalc();
    return read();
  } finally {
    for (const s of saved) {
      sheet.setCell(s.r, s.c, s.cell);
      doc.engine.cellChanged(sheet, s.r, s.c);
    }
    doc.engine.recalc();
  }
}

function numberOf(v: Scalar): number | null {
  return typeof v === 'number' && Number.isFinite(v) ? v : null;
}

proto.goalSeek = function (this: SpreadsheetDocument, set, target, by, sheet = this.sheet) {
  const formula = sheet.getCell(set.r, set.c);
  if (!formula?.f) return 'The Set cell must contain a formula.';
  const input = sheet.getCell(by.r, by.c);
  if (input?.f) return 'The By changing cell must contain a value, not a formula.';
  const f = (x: number): number | null => withInputs(this, sheet, [{ ...by, v: x }], () => numberOf(this.engine.cellValue(sheet, set.r, set.c)));
  const tol = 0.001;
  const maxIter = 100;
  let x0 = numberOf(input?.v ?? 0) ?? 0;
  let y0 = f(x0);
  if (y0 === null) return 'Goal Seek cannot find a solution: the formula does not return a number.';
  let best = { x: x0, y: y0 };
  let iterations = 0;
  if (Math.abs(y0 - target) > tol) {
    // Secant steps from two starting points; fall back to bisection once the root is bracketed.
    let x1 = x0 === 0 ? 1 : x0 * 1.01;
    let y1 = f(x1);
    let bracket: [number, number, number, number] | null = null;
    for (iterations = 1; iterations <= maxIter && y1 !== null; iterations++) {
      if (Math.abs(y1 - target) < Math.abs(best.y - target)) best = { x: x1, y: y1 };
      if (Math.abs(y1 - target) <= tol) break;
      if ((y0 - target) * (y1 - target) < 0) bracket = [x0, y0, x1, y1];
      let x2: number;
      if (bracket && iterations > 20) {
        const [a, fa, b]: [number, number, number, number] = bracket;
        x2 = (a + b) / 2;
        const y2 = f(x2);
        if (y2 === null) break;
        if ((fa - target) * (y2 - target) < 0) bracket = [a, fa, x2, y2];
        else bracket = [x2, y2, b, bracket[3]];
        x0 = x1;
        y0 = y1;
        x1 = x2;
        y1 = y2;
        continue;
      }
      const slope = (y1 - y0) / (x1 - x0);
      x2 = slope && Number.isFinite(slope) ? x1 + (target - y1) / slope : x1 + (x1 - x0 || 1);
      if (!Number.isFinite(x2)) break;
      x0 = x1;
      y0 = y1;
      x1 = x2;
      y1 = f(x1);
    }
  }
  const found = Math.abs(best.y - target) <= Math.max(tol, Math.abs(target) * 1e-9);
  // Like Excel, the cells show the best value reached; Cancel in the status dialog undoes it.
  this.transact('Goal Seek', (tx) => this.put(tx, sheet, by.r, by.c, { v: best.x, s: input?.s }));
  return { found, value: best.x, result: best.y, iterations };
};

proto.dataTable = function (this: SpreadsheetDocument, range, rowInput, colInput, sheet = this.sheet) {
  if (!rowInput && !colInput) return 'Enter a row input cell, a column input cell, or both.';
  if (range.r1 <= range.r0 && range.c1 <= range.c0) return 'Select the whole data table (values and formulas).';
  const valueOf = (r: number, c: number): Scalar => this.engine.cellValue(sheet, r, c);
  const results: { r: number; c: number; v: Scalar }[] = [];
  if (rowInput && colInput) {
    // Formula in the corner, row input values across the top, column input values down the side.
    for (let r = range.r0 + 1; r <= range.r1; r++) {
      for (let c = range.c0 + 1; c <= range.c1; c++) {
        const v = withInputs(this, sheet, [{ ...rowInput, v: valueOf(range.r0, c) }, { ...colInput, v: valueOf(r, range.c0) }], () => valueOf(range.r0, range.c0));
        results.push({ r, c, v });
      }
    }
  } else if (colInput) {
    // Input values in the first column, formulas across the first row.
    for (let r = range.r0 + 1; r <= range.r1; r++) {
      for (let c = range.c0 + 1; c <= range.c1; c++) {
        const v = withInputs(this, sheet, [{ ...colInput, v: valueOf(r, range.c0) }], () => valueOf(range.r0, c));
        results.push({ r, c, v });
      }
    }
  } else if (rowInput) {
    // Input values in the first row, formulas down the first column.
    for (let c = range.c0 + 1; c <= range.c1; c++) {
      for (let r = range.r0 + 1; r <= range.r1; r++) {
        const v = withInputs(this, sheet, [{ ...rowInput, v: valueOf(range.r0, c) }], () => valueOf(r, range.c0));
        results.push({ r, c, v });
      }
    }
  }
  this.transact('Data Table', (tx) => {
    for (const x of results) this.put(tx, sheet, x.r, x.c, { v: x.v, s: sheet.getCell(x.r, x.c)?.s });
  });
  return null;
};

proto.setScenarios = function (this: SpreadsheetDocument, next, sheet = this.sheet) {
  this.transact('Scenarios', (tx) => {
    tx.snapshotSheet(sheet);
    sheet.scenarios = next && next.list.length ? { list: next.list.map((s) => ({ ...s, cells: s.cells.map((c) => ({ ...c })), values: [...s.values] })), current: next.current } : undefined;
  });
};

function scenarioValue(doc: SpreadsheetDocument, text: string): Scalar {
  const parsed = parseInput(text, { textFormat: false, date1904: doc.wb.date1904 });
  return parsed.formula !== undefined ? text : parsed.value;
}

proto.showScenario = function (this: SpreadsheetDocument, index, sheet = this.sheet) {
  const sc = sheet.scenarios?.list[index];
  if (!sc) return;
  this.transact(`Show scenario ${sc.name}`, (tx) => {
    sc.cells.forEach((cell, i) => {
      const old = sheet.getCell(cell.r, cell.c);
      this.put(tx, sheet, cell.r, cell.c, { v: scenarioValue(this, sc.values[i] ?? ''), s: old?.s });
    });
    tx.snapshotSheet(sheet);
    if (sheet.scenarios) sheet.scenarios = { ...sheet.scenarios, current: index };
  });
};

proto.scenarioSummary = function (this: SpreadsheetDocument, resultCells, sheet = this.sheet) {
  const list = sheet.scenarios?.list ?? [];
  const changing = [...new Map(list.flatMap((s) => s.cells).map((c) => [`${c.r}:${c.c}`, c])).values()];
  const current = changing.map((c) => this.engine.cellValue(sheet, c.r, c.c));
  const currentResults = resultCells.map((c) => this.engine.cellValue(sheet, c.r, c.c));
  const perScenario = list.map((sc) =>
    withInputs(
      this,
      sheet,
      sc.cells.map((cell, i) => ({ ...cell, v: scenarioValue(this, sc.values[i] ?? '') })),
      () => ({ changing: changing.map((c) => this.engine.cellValue(sheet, c.r, c.c)), results: resultCells.map((c) => this.engine.cellValue(sheet, c.r, c.c)) }),
    ),
  );
  let index = -1;
  this.transact('Scenario Summary', () => {
    const names = new Set(this.wb.sheets.map((s) => s.name.toLowerCase()));
    let name = 'Scenario Summary';
    for (let k = 2; names.has(name.toLowerCase()); k++) name = `Scenario Summary ${k}`;
    index = this.wb.sheets.indexOf(sheet);
    this.addSheet(name, index);
    const out = this.wb.sheets[index];
    const rows: Scalar[][] = [];
    rows.push([null, 'Scenario Summary']);
    rows.push([null, null, 'Current Values:', ...list.map((s) => s.name)]);
    rows.push([null, 'Changing Cells:']);
    changing.forEach((c, i) => rows.push([null, cellName(c.r, c.c, true, true), current[i], ...perScenario.map((p) => p.changing[i])]));
    if (resultCells.length) {
      rows.push([null, 'Result Cells:']);
      resultCells.forEach((c, i) => rows.push([null, cellName(c.r, c.c, true, true), currentResults[i], ...perScenario.map((p) => p.results[i])]));
    }
    rows.push([null, 'Notes: Current Values column represents values of changing cells at']);
    rows.push([null, 'time Scenario Summary Report was created. Changing cells for each']);
    rows.push([null, 'scenario are highlighted in gray.']);
    const bold = this.wb.styles.patch(0, { font: { bold: true } });
    const header = this.wb.styles.patch(0, { font: { bold: true, color: { theme: 0 } }, fill: { pattern: 'solid', fg: { theme: 4, tint: -0.249977111117893 }, bg: { indexed: 64 } } });
    rows.forEach((row, r) => {
      row.forEach((v, c) => {
        if (v === null) return;
        const s = r === 0 ? bold : r === 1 && c >= 2 ? header : undefined;
        out.setCell(r + 1, c, { v, s });
      });
    });
    out.setColInfo(1, { w: 18, customWidth: true });
    out.setColInfo(2, { w: 14, customWidth: true });
  });
  return index;
};
