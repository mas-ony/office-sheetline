/** Page Layout commands: page setup, print area, print titles and manual page breaks. */
import { SpreadsheetDocument } from '../document';
import type { Sheet } from '../model/sheet';
import type { PageSetup } from '../model/page';
import { cellName, colName, parseRange, quoteSheetName, splitSheetRef, type Range } from '../refs';

declare module '../document' {
  interface SpreadsheetDocument {
    setPageSetup(patch: Partial<PageSetup>, sheet?: Sheet): void;
    printArea(sheet?: Sheet): Range | null;
    setPrintArea(range: Range | null, sheet?: Sheet): void;
    /** Rows and columns repeated on every printed page (inclusive index pairs). */
    printTitles(sheet?: Sheet): { rows: [number, number] | null; cols: [number, number] | null };
    setPrintTitles(rows: [number, number] | null, cols: [number, number] | null, sheet?: Sheet): void;
    /** Inserts a page break above and/or left of the active cell. */
    insertPageBreak(sheet?: Sheet): void;
    removePageBreak(sheet?: Sheet): void;
    resetPageBreaks(sheet?: Sheet): void;
  }
}

const proto = SpreadsheetDocument.prototype;
const AREA = '_xlnm.Print_Area';
const TITLES = '_xlnm.Print_Titles';

function nameFor(doc: SpreadsheetDocument, sheet: Sheet, name: string) {
  const index = doc.wb.sheets.indexOf(sheet);
  return doc.wb.names.find((n) => n.name.toLowerCase() === name.toLowerCase() && n.localSheet === index);
}

function setName(doc: SpreadsheetDocument, sheet: Sheet, name: string, formula: string | null): void {
  const index = doc.wb.sheets.indexOf(sheet);
  doc.wb.names = doc.wb.names.filter((n) => !(n.name.toLowerCase() === name.toLowerCase() && n.localSheet === index));
  if (formula) doc.wb.names.push({ name, formula, localSheet: index });
}

proto.setPageSetup = function (this: SpreadsheetDocument, patch, sheet = this.sheet) {
  this.transact('Page setup', (tx) => {
    tx.snapshotSheet(sheet);
    sheet.page = { ...sheet.page, ...patch, margins: { ...sheet.page.margins, ...(patch.margins ?? {}) }, headerFooter: { ...sheet.page.headerFooter, ...(patch.headerFooter ?? {}) } };
  });
};

proto.printArea = function (this: SpreadsheetDocument, sheet = this.sheet) {
  const n = nameFor(this, sheet, AREA);
  if (!n) return null;
  return parseRange(splitSheetRef(n.formula.split(',')[0]).ref.replace(/\$/g, ''));
};

proto.setPrintArea = function (this: SpreadsheetDocument, range, sheet = this.sheet) {
  this.transact(range ? 'Set print area' : 'Clear print area', (tx) => {
    tx.snapshotWorkbook(this.wb);
    const ref = range ? `${quoteSheetName(sheet.name)}!${cellName(range.r0, range.c0, true, true)}:${cellName(range.r1, range.c1, true, true)}` : null;
    setName(this, sheet, AREA, ref);
  });
};

proto.printTitles = function (this: SpreadsheetDocument, sheet = this.sheet) {
  const out: { rows: [number, number] | null; cols: [number, number] | null } = { rows: null, cols: null };
  const n = nameFor(this, sheet, TITLES);
  if (!n) return out;
  for (const part of n.formula.split(',')) {
    const ref = splitSheetRef(part.trim()).ref.replace(/\$/g, '');
    const rows = /^(\d+):(\d+)$/.exec(ref);
    const cols = /^([A-Za-z]+):([A-Za-z]+)$/.exec(ref);
    if (rows) out.rows = [Number(rows[1]) - 1, Number(rows[2]) - 1];
    else if (cols) {
      const r = parseRange(`${cols[1]}1:${cols[2]}1`);
      if (r) out.cols = [r.c0, r.c1];
    }
  }
  return out;
};

proto.setPrintTitles = function (this: SpreadsheetDocument, rows, cols, sheet = this.sheet) {
  this.transact('Print titles', (tx) => {
    tx.snapshotWorkbook(this.wb);
    const q = quoteSheetName(sheet.name);
    const parts: string[] = [];
    if (cols) parts.push(`${q}!$${colName(cols[0])}:$${colName(cols[1])}`);
    if (rows) parts.push(`${q}!$${rows[0] + 1}:$${rows[1] + 1}`);
    setName(this, sheet, TITLES, parts.length ? parts.join(',') : null);
  });
};

proto.insertPageBreak = function (this: SpreadsheetDocument, sheet = this.sheet) {
  const { r, c } = this.selectionOf(sheet).active;
  const p = sheet.page;
  const rowBreaks = r > 0 && !p.rowBreaks.includes(r) ? [...p.rowBreaks, r] : p.rowBreaks;
  const colBreaks = c > 0 && !p.colBreaks.includes(c) ? [...p.colBreaks, c] : p.colBreaks;
  this.setPageSetup({ rowBreaks, colBreaks }, sheet);
};

proto.removePageBreak = function (this: SpreadsheetDocument, sheet = this.sheet) {
  const { r, c } = this.selectionOf(sheet).active;
  const p = sheet.page;
  this.setPageSetup({ rowBreaks: p.rowBreaks.filter((x) => x !== r), colBreaks: p.colBreaks.filter((x) => x !== c) }, sheet);
};

proto.resetPageBreaks = function (this: SpreadsheetDocument, sheet = this.sheet) {
  this.setPageSetup({ rowBreaks: [], colBreaks: [] }, sheet);
};

export {};
