/** References as dialogs show and read them. */
import { resolveRef } from '../core/chartData';
import type { SpreadsheetDocument } from '../core/document';
import type { Sheet } from '../core/model/sheet';
import { cellName, quoteSheetName, type Range } from '../core/refs';

/** Text of a range as a dialog shows it: absolute, sheet-qualified when on another sheet. */
export function rangeText(r: Range, sheet: Sheet, home?: Sheet): string {
  const a = cellName(r.r0, r.c0, true, true);
  const b = cellName(r.r1, r.c1, true, true);
  const ref = a === b ? a : `${a}:${b}`;
  return home && home !== sheet ? `${quoteSheetName(sheet.name)}!${ref}` : ref;
}

/** Selection of the active sheet as range text (several areas joined with commas). */
export function selectionText(doc: SpreadsheetDocument, home?: Sheet): string {
  return doc.selection.ranges.map((r) => rangeText(doc.bounded(r, doc.sheet) ?? r, doc.sheet, home)).join(',');
}

/** Resolves a typed reference ("A1:B4", "Sheet2!$A$1", a defined name). */
export function parseRangeText(doc: SpreadsheetDocument, text: string, home?: Sheet): { sheet: Sheet; range: Range } | null {
  const t = text.trim().replace(/^=/, '');
  if (!t) return null;
  return resolveRef(doc, t, home ?? doc.sheet);
}
