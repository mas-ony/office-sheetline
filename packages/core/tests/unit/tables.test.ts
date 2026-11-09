import { describe, expect, it } from 'vitest';
import '../../src/core/doc';
import { SpreadsheetDocument } from '../../src/core/document';
import { makeRange } from '../../src/core/refs';
import { tableCellDecor } from '../../src/core/tableStyles';

function sample(): SpreadsheetDocument {
  const doc = SpreadsheetDocument.blank();
  doc.setValues(0, 0, [
    ['Region', 'Q1', 'Q2'],
    ['North', 120, 135],
    ['South', 98, 110],
    ['East', 143, 128],
  ]);
  return doc;
}

describe('tables', () => {
  it('creates a table with column names from the header row', () => {
    const doc = sample();
    const t = doc.createTable(makeRange(0, 0, 3, 2));
    expect(typeof t).toBe('object');
    const table = doc.sheet.tables[0];
    expect(table.name).toBe('Table1');
    expect(table.columns.map((c) => c.name)).toEqual(['Region', 'Q1', 'Q2']);
    expect(table.style.name).toBe('TableStyleMedium2');
    expect(tableCellDecor(table, 0, 0, doc.wb.theme)?.bold).toBe(true);
  });

  it('evaluates structured references and keeps them when the header changes', () => {
    const doc = sample();
    doc.createTable(makeRange(0, 0, 3, 2));
    doc.setInput(5, 1, '=SUM(Table1[Q1])');
    doc.setInput(5, 2, '=COUNTA(Table1[[#Headers],[Region]:[Q2]])');
    doc.setInput(6, 1, '=SUM(Table1)');
    expect(doc.cell(5, 1)?.v).toBe(361);
    expect(doc.cell(5, 2)?.v).toBe(3);
    expect(doc.cell(6, 1)?.v).toBe(361 + 373);
    doc.setInput(0, 1, 'Jan');
    expect(doc.sheet.tables[0].columns[1].name).toBe('Jan');
    expect(doc.cell(5, 1)?.f).toBe('SUM(Table1[Jan])');
    expect(doc.cell(5, 1)?.v).toBe(361);
  });

  it('uses [@Col] inside the table and fills calculated columns', () => {
    const doc = sample();
    doc.createTable(makeRange(0, 0, 3, 2));
    doc.setInput(0, 3, 'Total');
    expect(doc.sheet.tables[0].ref.c1).toBe(3);
    doc.setInput(1, 3, '=[@Q1]+[@Q2]');
    expect(doc.cell(1, 3)?.v).toBe(255);
    expect(doc.cell(3, 3)?.v).toBe(271);
    // Typing below the table extends it.
    doc.setInput(4, 0, 'West');
    expect(doc.sheet.tables[0].ref.r1).toBe(4);
  });

  it('adds a totals row with SUBTOTAL formulas', () => {
    const doc = sample();
    doc.createTable(makeRange(0, 0, 3, 2));
    doc.setTableOptions(1, { totalsRow: true });
    const t = doc.sheet.tables[0];
    expect(t.ref.r1).toBe(4);
    expect(doc.cell(4, 0)?.v).toBe('Total');
    expect(doc.cell(4, 2)?.f).toBe('SUBTOTAL(109,Table1[Q2])');
    expect(doc.cell(4, 2)?.v).toBe(373);
    doc.undo();
    expect(doc.sheet.tables[0].totalsRow).toBe(false);
  });

  it('round-trips tables through XLSX', () => {
    const doc = sample();
    doc.createTable(makeRange(0, 0, 3, 2), { style: 'TableStyleLight9' });
    doc.setTableOptions(1, { totalsRow: true, showColumnStripes: true });
    doc.setInput(1, 3, '=[@Q1]*2');
    const bytes = doc.save();
    const back = SpreadsheetDocument.fromXlsx(bytes);
    const t = back.sheet.tables[0];
    expect(t.name).toBe('Table1');
    expect(t.style.name).toBe('TableStyleLight9');
    expect(t.style.showColumnStripes).toBe(true);
    expect(t.totalsRow).toBe(true);
    expect(t.columns[2].totalsFunction).toBe('sum');
    expect(back.cell(4, 2)?.v).toBe(373);
  });

  it('converts a table to a range with plain references', () => {
    const doc = sample();
    doc.createTable(makeRange(0, 0, 3, 2));
    doc.setInput(5, 1, '=SUM(Table1[Q1])');
    doc.convertTableToRange(1);
    expect(doc.sheet.tables.length).toBe(0);
    expect(doc.cell(5, 1)?.f).toBe('SUM($B$2:$B$4)');
    expect(doc.styleAt(0, 0).font.bold).toBe(true);
  });
});
