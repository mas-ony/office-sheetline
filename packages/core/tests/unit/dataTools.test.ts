import { describe, expect, it } from 'vitest';
import '../../src/core/doc';
import { SpreadsheetDocument } from '../../src/core/document';
import { outlineGroups } from '../../src/core/doc/outline';
import { CUSTOM_LISTS, splitText, type TextToColumnsOptions } from '../../src/core/doc/dataTools';
import { parseRange } from '../../src/core/refs';
import { readXlsx } from '../../src/core/xlsx/import';

function v(d: SpreadsheetDocument, ref: string) {
  const r = parseRange(ref)!;
  return d.sheet.getCell(r.r0, r.c0)?.v ?? null;
}

function f(d: SpreadsheetDocument, ref: string) {
  const r = parseRange(ref)!;
  return d.sheet.getCell(r.r0, r.c0)?.f;
}

function column(d: SpreadsheetDocument, ref: string) {
  const r = parseRange(ref)!;
  const out = [];
  for (let row = r.r0; row <= r.r1; row++) out.push(d.sheet.getCell(row, r.c0)?.v ?? null);
  return out;
}

const hidden = (d: SpreadsheetDocument, r: number) => !!d.sheet.rows.get(r)?.hidden;

function salesDoc() {
  const d = SpreadsheetDocument.blank();
  d.setValues(0, 0, [
    ['Region', 'Rep', 'Sales'],
    ['East', 'Ann', 10],
    ['East', 'Bob', 20],
    ['West', 'Cid', 5],
    ['West', 'Dee', 15],
    ['West', 'Eve', 25],
  ]);
  return d;
}

describe('outline groups', () => {
  it('groups, collapses and expands rows', () => {
    const d = SpreadsheetDocument.blank();
    expect(d.groupLines('row', 1, 3)).toBeNull();
    expect(d.groupLines('row', 2, 3)).toBeNull();
    const groups = outlineGroups(d.sheet, 'row');
    expect(groups.map((g) => [g.level, g.start, g.end, g.summary])).toEqual([[1, 1, 3, 4], [2, 2, 3, 4]]);
    d.setGroupCollapsed(groups[0], true);
    expect([1, 2, 3].map((r) => hidden(d, r))).toEqual([true, true, true]);
    expect(d.sheet.rows.get(4)?.collapsed).toBe(true);
    // Level buttons: level 2 shows the outer group and hides the inner one.
    d.showOutlineLevel('row', 2);
    expect([1, 2, 3].map((r) => hidden(d, r))).toEqual([false, true, true]);
    d.undo();
    expect([1, 2, 3].map((r) => hidden(d, r))).toEqual([true, true, true]);
    d.showOutlineLevel('row', 3);
    expect([1, 2, 3].map((r) => hidden(d, r))).toEqual([false, false, false]);
  });

  it('keeps nested collapsed groups hidden when the outer group expands', () => {
    const d = SpreadsheetDocument.blank();
    d.groupLines('col', 1, 4);
    d.groupLines('col', 2, 3);
    const [outer, inner] = outlineGroups(d.sheet, 'col');
    d.setGroupCollapsed(inner, true);
    d.setGroupCollapsed(outlineGroups(d.sheet, 'col')[0], true);
    d.setGroupCollapsed({ ...outer, collapsed: true }, false);
    expect([1, 2, 3, 4].map((c) => !!d.sheet.cols.get(c)?.hidden)).toEqual([false, true, true, false]);
  });

  it('refuses an eighth level and ungroups one level at a time', () => {
    const d = SpreadsheetDocument.blank();
    for (let i = 0; i < 7; i++) expect(d.groupLines('row', 1, 2)).toBeNull();
    expect(d.groupLines('row', 1, 2)).toMatch(/seven levels/);
    expect(d.ungroupLines('row', 1, 1)).toBe(true);
    expect(d.sheet.rows.get(1)?.outline).toBe(6);
    d.clearOutline();
    expect(outlineGroups(d.sheet, 'row')).toEqual([]);
  });

  it('builds an outline from summary formulas', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[1], [2], [3]]);
    d.setInput(3, 0, '=SUM(A1:A3)');
    expect(d.autoOutline()).toBe(true);
    expect(outlineGroups(d.sheet, 'row').map((g) => [g.start, g.end, g.summary])).toEqual([[0, 2, 3]]);
  });

  it('places summary lines above when the setting says so and saves the outline', () => {
    const d = SpreadsheetDocument.blank();
    d.groupLines('row', 2, 4);
    d.setOutlineSettings({ summaryBelow: false });
    expect(outlineGroups(d.sheet, 'row')[0].summary).toBe(1);
    d.setGroupCollapsed(outlineGroups(d.sheet, 'row')[0], true);
    const back = readXlsx(d.save()).sheets[0];
    expect(back.outlineBelow).toBe(false);
    expect([2, 3, 4].map((r) => back.rows.get(r)?.outline)).toEqual([1, 1, 1]);
    expect([2, 3, 4].map((r) => !!back.rows.get(r)?.hidden)).toEqual([true, true, true]);
    expect(back.rows.get(1)?.collapsed).toBe(true);
  });
});

describe('subtotals', () => {
  it('inserts a SUBTOTAL row per group and a grand total, then removes them', () => {
    const d = salesDoc();
    const region = parseRange('A1:C6')!;
    expect(d.applySubtotals(region, { groupBy: 0, code: 9, columns: [2], replace: true, pageBreaks: false, summaryBelow: true })).toBeNull();
    expect(column(d, 'A1:A9')).toEqual(['Region', 'East', 'East', 'East Total', 'West', 'West', 'West', 'West Total', 'Grand Total']);
    expect(f(d, 'C4')).toBe('SUBTOTAL(9,C2:C3)');
    expect(v(d, 'C4')).toBe(30);
    expect(v(d, 'C8')).toBe(45);
    // SUBTOTAL skips the other subtotals.
    expect(f(d, 'C9')).toBe('SUBTOTAL(9,C2:C8)');
    expect(v(d, 'C9')).toBe(75);
    expect([1, 3, 8].map((r) => d.sheet.rows.get(r)?.outline ?? 0)).toEqual([2, 1, 0]);
    d.removeSubtotals(parseRange('A1:C9')!);
    expect(column(d, 'A1:A7')).toEqual(['Region', 'East', 'East', 'West', 'West', 'West', null]);
    expect(outlineGroups(d.sheet, 'row')).toEqual([]);
  });

  it('counts with another function and adds page breaks between groups', () => {
    const d = salesDoc();
    d.applySubtotals(parseRange('A1:C6')!, { groupBy: 0, code: 3, columns: [1], replace: false, pageBreaks: true, summaryBelow: true });
    expect(v(d, 'A4')).toBe('East Count');
    expect(v(d, 'B4')).toBe(2);
    expect(v(d, 'A9')).toBe('Grand Count');
    expect(d.sheet.page.rowBreaks).toEqual([4]);
  });
});

describe('sorting', () => {
  it('sorts by several levels and by a custom list', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['Day', 'N'], ['Wed', 2], ['Mon', 3], ['Wed', 1], ['Sun', 9]]);
    d.sortLevels(parseRange('A1:B5')!, [
      { index: 0, order: 'custom', sortOn: 'values', list: CUSTOM_LISTS[0] },
      { index: 1, order: 'asc', sortOn: 'values' },
    ], { hasHeader: true });
    expect(column(d, 'A1:A5')).toEqual(['Day', 'Sun', 'Mon', 'Wed', 'Wed']);
    expect(column(d, 'B2:B5')).toEqual([9, 3, 1, 2]);
  });

  it('puts a fill colour on top', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['a'], ['b'], ['c']]);
    d.format({ fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } } }, [parseRange('A3')!]);
    d.sortLevels(parseRange('A1:A3')!, [{ index: 0, order: 'asc', sortOn: 'cellColor', color: '#FFFF00' }], { hasHeader: false });
    expect(column(d, 'A1:A3')).toEqual(['c', 'a', 'b']);
  });

  it('sorts columns left to right', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[3, 1, 2], ['c', 'a', 'b']]);
    d.sortLevels(parseRange('A1:C2')!, [{ index: 0, order: 'desc', sortOn: 'values' }], { hasHeader: false, byColumns: true });
    expect([v(d, 'A2'), v(d, 'B2'), v(d, 'C2')]).toEqual(['c', 'b', 'a']);
  });
});

const delimited = (patch: Partial<TextToColumnsOptions> = {}): TextToColumnsOptions => ({
  type: 'delimited',
  delimiters: { tab: false, semicolon: false, comma: true, space: false, other: '' },
  consecutive: false,
  qualifier: '"',
  breaks: [],
  ...patch,
});

describe('text to columns', () => {
  it('splits on delimiters with a text qualifier and converts values', () => {
    expect(splitText('a,"b,c",3', delimited())).toEqual(['a', 'b,c', '3']);
    expect(splitText('a  b', delimited({ delimiters: { tab: false, semicolon: false, comma: false, space: true, other: '' }, consecutive: true }))).toEqual(['a', 'b']);
    expect(splitText('abcdef', delimited({ type: 'fixed', breaks: [2, 4] }))).toEqual(['ab', 'cd', 'ef']);
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['Ann,7,2024-01-15'], ['Bob,0012,x']]);
    expect(d.textToColumns(parseRange('A1:A2')!, delimited({ formats: ['general', 'text', 'general'] }))).toBeNull();
    expect([v(d, 'A1'), v(d, 'B1'), v(d, 'C1')]).toEqual(['Ann', '7', 45306]);
    expect(v(d, 'B2')).toBe('0012');
  });

  it('skips columns, writes to a destination and reports overwrites', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['a;b;c', null, 'keep']]);
    const opts = delimited({ delimiters: { tab: false, semicolon: true, comma: false, space: false, other: '' }, formats: ['general', 'skip', 'general'] });
    expect(d.textToColumnsOverwrites(parseRange('A1')!, opts)).toBe(false);
    expect(d.textToColumnsOverwrites(parseRange('A1')!, { ...opts, formats: undefined })).toBe(true);
    d.textToColumns(parseRange('A1')!, { ...opts, destination: { r: 2, c: 0 } });
    expect([v(d, 'A3'), v(d, 'B3'), v(d, 'C3')]).toEqual(['a', 'c', null]);
    expect(v(d, 'A1')).toBe('a;b;c');
  });
});

describe('remove duplicates, insert and delete cells, series', () => {
  it('removes duplicate rows by the chosen columns', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['K', 'V'], ['a', 1], ['A', 2], ['b', 1], ['a', 1]]);
    expect(d.removeDuplicates(parseRange('A1:B5')!, [0], true)).toEqual({ removed: 2, remaining: 2 });
    expect(column(d, 'A1:A5')).toEqual(['K', 'a', 'b', null, null]);
    d.undo();
    expect(d.removeDuplicates(parseRange('A1:B5')!, [0, 1], true)).toEqual({ removed: 1, remaining: 3 });
  });

  it('shifts cells down or left and adjusts references to them', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[1, 10], [2, 20], [3, 30]]);
    d.setInput(0, 3, '=SUM(A1:A3)');
    d.setInput(1, 3, '=B2');
    expect(d.insertCells(parseRange('A2')!, 'down')).toBeNull();
    expect(column(d, 'A1:A4')).toEqual([1, null, 2, 3]);
    expect(column(d, 'B1:B3')).toEqual([10, 20, 30]);
    expect(f(d, 'D1')).toBe('SUM(A1:A4)');
    expect(v(d, 'D1')).toBe(6);
    expect(d.deleteCells(parseRange('A1:A2')!, 'left')).toBeNull();
    expect([v(d, 'A1'), v(d, 'A2'), v(d, 'A3')]).toEqual([10, 20, 2]);
    // D1:D2 moved left with their rows; the reference to B2 follows the cell it pointed to.
    expect(f(d, 'C2')).toBe('A2');
    expect(v(d, 'C2')).toBe(20);
  });

  it('refuses to split a merged cell when shifting', () => {
    const d = SpreadsheetDocument.blank();
    d.merge([parseRange('A2:B2')!]);
    expect(d.insertCells(parseRange('A1')!, 'down')).toMatch(/merged cells/);
  });

  it('fills linear, growth and date series', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[2], [null], [null], [null]]);
    d.fillSeriesRange(parseRange('A1:A4')!, { by: 'columns', type: 'linear', unit: 'day', step: 3, trend: false });
    expect(column(d, 'A1:A4')).toEqual([2, 5, 8, 11]);
    d.setValues(0, 1, [[1]]);
    d.fillSeriesRange(parseRange('B1:B6')!, { by: 'columns', type: 'growth', unit: 'day', step: 2, stop: 10, trend: false });
    expect(column(d, 'B1:B5')).toEqual([1, 2, 4, 8, null]);
    d.setInput(0, 2, '2024-01-31');
    d.fillSeriesRange(parseRange('C1:C3')!, { by: 'columns', type: 'date', unit: 'month', step: 1, trend: false });
    expect([1, 2].map((r) => d.displayText(r, 2))).toEqual(['2024-02-29', '2024-03-31']);
  });

  it('finds cells that break their validation rule', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[5], [50], ['x']]);
    d.addValidation({ sqref: [parseRange('A1:A3')!], type: 'whole', operator: 'between', formula1: '1', formula2: '10', allowBlank: true, showErrorMessage: true, errorStyle: 'stop' });
    expect(d.invalidCells()).toEqual([{ r: 1, c: 0 }, { r: 2, c: 0 }]);
  });
});

describe('what-if analysis', () => {
  it('seeks a goal by changing an input cell', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[100], [0.05]]);
    d.setInput(2, 0, '=A1*(1+A2)^10');
    const result = d.goalSeek({ r: 2, c: 0 }, 200, { r: 1, c: 0 });
    expect(typeof result).toBe('object');
    if (typeof result === 'string') return;
    expect(result.found).toBe(true);
    expect(result.value).toBeCloseTo(0.071773, 5);
    expect(v(d, 'A3')).toBeCloseTo(200, 2);
    expect(d.goalSeek({ r: 0, c: 0 }, 1, { r: 1, c: 0 })).toMatch(/must contain a formula/);
  });

  it('fills one- and two-input data tables', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[2], [3]]);
    d.setInput(0, 4, '=A1*A2');
    // Column input: values down D2:D4, formula in E1.
    d.setValues(1, 3, [[1], [2], [3]]);
    expect(d.dataTable(parseRange('D1:E4')!, null, { r: 0, c: 0 })).toBeNull();
    expect(column(d, 'E2:E4')).toEqual([3, 6, 9]);
    expect(v(d, 'A1')).toBe(2);
    // Two inputs: formula in the corner.
    d.setInput(5, 0, '=A1*A2');
    d.setValues(5, 1, [[10, 20]]);
    d.setValues(6, 0, [[1], [2]]);
    d.dataTable(parseRange('A6:C8')!, { r: 0, c: 0 }, { r: 1, c: 0 });
    expect([v(d, 'B7'), v(d, 'C7'), v(d, 'B8'), v(d, 'C8')]).toEqual([10, 20, 20, 40]);
  });

  it('shows scenarios, summarises them and saves them in the file', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[10], [2]]);
    d.setInput(2, 0, '=A1*A2');
    d.setScenarios({ list: [
      { name: 'Low', cells: [{ r: 0, c: 0 }, { r: 1, c: 0 }], values: ['1', '1'] },
      { name: 'High', cells: [{ r: 0, c: 0 }, { r: 1, c: 0 }], values: ['100', '3'], comment: 'Best case' },
    ] });
    d.showScenario(1);
    expect(v(d, 'A3')).toBe(300);
    expect(d.sheet.scenarios?.current).toBe(1);
    const back = readXlsx(d.save()).sheets[0].scenarios!;
    expect(back.list.map((s) => [s.name, s.values, s.comment])).toEqual([['Low', ['1', '1'], undefined], ['High', ['100', '3'], 'Best case']]);
    expect(back.list[0].cells).toEqual([{ r: 0, c: 0 }, { r: 1, c: 0 }]);
    const index = d.scenarioSummary([{ r: 2, c: 0 }]);
    const summary = d.wb.sheets[index];
    expect(summary.name).toBe('Scenario Summary');
    const texts: unknown[] = [];
    for (const row of summary.cells.values()) for (const cell of row.values()) texts.push(cell.v);
    expect(texts).toEqual(expect.arrayContaining(['Low', 'High', 1, 300]));
  });
});
