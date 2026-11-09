/** Print options of a sheet from its Excel page setup (File > Print and PDF export). */
import type { SpreadsheetDocument } from '../core/document';
import type { Sheet } from '../core/model/sheet';
import { paperOf } from '../core/model/page';
import { DEFAULT_PRINT_OPTIONS, type PrintOptions } from './pages';

export interface PrintChoice {
  scope: PrintOptions['scope'];
  pageFrom?: number;
  pageTo?: number;
}

export function printOptionsFor(doc: SpreadsheetDocument, sheet: Sheet, choice: PrintChoice = { scope: 'sheet' }): PrintOptions {
  const p = sheet.page;
  const paper = paperOf(p.paperSize);
  const titles = doc.printTitles(sheet);
  const inch = 72;
  return {
    ...DEFAULT_PRINT_OPTIONS,
    scope: choice.scope,
    orientation: p.orientation,
    paperPt: { w: paper.w * inch, h: paper.h * inch },
    marginsPt: {
      t: p.margins.top * inch,
      b: p.margins.bottom * inch,
      l: p.margins.left * inch,
      r: p.margins.right * inch,
      header: p.margins.header * inch,
      footer: p.margins.footer * inch,
    },
    scale: p.fitToPage ? 'fitTo' : 'percent',
    scalePct: p.scale,
    fitWide: p.fitToWidth,
    fitTall: p.fitToHeight,
    order: p.pageOrder === 'overThenDown' ? 'leftToRight' : 'topToBottom',
    hAlign: p.centerH ? 'center' : 'left',
    vAlign: p.centerV ? 'center' : 'top',
    gridlines: p.gridLines,
    headings: p.headings,
    notes: false,
    titleRows: titles.rows,
    titleCols: titles.cols,
    repeatRows: false,
    repeatCols: false,
    rowBreaks: p.rowBreaks,
    colBreaks: p.colBreaks,
    headerFooter: p.headerFooter,
    blackAndWhite: p.blackAndWhite,
    firstPageNumber: p.firstPageNumber,
    pageFrom: choice.pageFrom,
    pageTo: choice.pageTo,
    objects: true,
  };
}
