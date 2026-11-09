/**
 * What the data dialogs of both interfaces share: the sort levels' lines, the validation choices
 * and how a rule replaces the selection's previous ones, and the scenarios' changing cells.
 */
import type { SpreadsheetDocument } from '../core/document';
import type { DataValidation, Sheet, ValidationOperator } from '../core/model/sheet';
import type { Scenario } from '../core/model/scenario';
import { cellName, colName, type Range, rangesIntersect } from '../core/refs';
import { parseRangeText } from './rangeText';

// -----------------------------------------------------------------------------------------------
// Sort

/** The columns a sort level can use (the rows for a left to right sort). */
export function sortLines(region: Range, byColumns: boolean): number[] {
  return byColumns ? Array.from({ length: region.r1 - region.r0 + 1 }, (_, i) => region.r0 + i) : Array.from({ length: region.c1 - region.c0 + 1 }, (_, i) => region.c0 + i);
}

/** The name of a sort line: its header, or Row 1 / Column A. */
export function sortLineName(doc: SpreadsheetDocument, region: Range, byColumns: boolean, headers: boolean, i: number): string {
  if (byColumns) return headers ? doc.displayText(i, region.c0) || `Row ${i + 1}` : `Row ${i + 1}`;
  return headers ? doc.displayText(region.r0, i) || `Column ${colName(i)}` : `Column ${colName(i)}`;
}

/** Whether a sort line holds numbers: their orders read Smallest to Largest instead of A to Z. */
export function sortLineNumeric(doc: SpreadsheetDocument, region: Range, byColumns: boolean, headers: boolean, i: number): boolean {
  const n = byColumns ? region.c1 - region.c0 + 1 : region.r1 - region.r0 + 1;
  for (let k = headers && !byColumns ? 1 : 0; k < n; k++) {
    const v = byColumns ? doc.sheet.getCell(i, region.c0 + k)?.v : doc.sheet.getCell(region.r0 + k, i)?.v;
    if (v !== null && v !== undefined && v !== '') return typeof v === 'number';
  }
  return false;
}

// -----------------------------------------------------------------------------------------------
// Data Validation

export const VALIDATION_ALLOW_OPTIONS: { value: string; label: string }[] = [
  { value: 'none', label: 'Any value' },
  { value: 'whole', label: 'Whole number' },
  { value: 'decimal', label: 'Decimal' },
  { value: 'list', label: 'List' },
  { value: 'date', label: 'Date' },
  { value: 'time', label: 'Time' },
  { value: 'textLength', label: 'Text length' },
  { value: 'custom', label: 'Custom' },
];

export const VALIDATION_OPERATOR_OPTIONS: { value: ValidationOperator; label: string }[] = [
  { value: 'between', label: 'between' },
  { value: 'notBetween', label: 'not between' },
  { value: 'equal', label: 'equal to' },
  { value: 'notEqual', label: 'not equal to' },
  { value: 'greaterThan', label: 'greater than' },
  { value: 'lessThan', label: 'less than' },
  { value: 'greaterThanOrEqual', label: 'greater than or equal to' },
  { value: 'lessThanOrEqual', label: 'less than or equal to' },
];

/** Parts of `a` outside `b`. */
function subtract(a: Range, b: Range): Range[] {
  if (!rangesIntersect(a, b)) return [a];
  const out: Range[] = [];
  if (a.r0 < b.r0) out.push({ ...a, r1: b.r0 - 1 });
  if (a.r1 > b.r1) out.push({ ...a, r0: b.r1 + 1 });
  const r0 = Math.max(a.r0, b.r0);
  const r1 = Math.min(a.r1, b.r1);
  if (a.c0 < b.c0) out.push({ r0, r1, c0: a.c0, c1: b.c0 - 1 });
  if (a.c1 > b.c1) out.push({ r0, r1, c0: b.c1 + 1, c1: a.c1 });
  return out;
}

/** The formula of a validation value box: formulas as typed, dates and times as serial numbers. */
export function validationFormula(doc: SpreadsheetDocument, r: number, c: number, type: string, text: string): string {
  const t = text.trim();
  if (t.startsWith('=')) return t.slice(1);
  if (type === 'date' || type === 'time') {
    const v = doc.parsedInput(r, c, t);
    return typeof v === 'number' ? String(v) : t;
  }
  return t;
}

/** Data Validation's OK: the selection's cells leave their previous rules and take `rule` (none when null). */
export function setSelectionValidation(doc: SpreadsheetDocument, sel: Range[], rule: DataValidation | null): void {
  doc.transact('Data Validation', () => {
    // The selection's cells leave their previous rules.
    for (let i = doc.sheet.validations.length - 1; i >= 0; i--) {
      const v = doc.sheet.validations[i];
      if (!v.sqref.some((x) => sel.some((s) => rangesIntersect(x, s)))) continue;
      let rest = v.sqref;
      for (const s of sel) rest = rest.flatMap((x) => subtract(x, s));
      if (rest.length) doc.updateValidation(i, { ...v, sqref: rest });
      else doc.removeValidation(i);
    }
    if (rule) doc.addValidation(rule);
  });
}

// -----------------------------------------------------------------------------------------------
// Scenarios

/** The changing cells of a scenario as text ($A$1,$B$2). */
export function scenarioCellsText(s: Scenario): string {
  return s.cells.map((c) => cellName(c.r, c.c, true, true)).join(',');
}

/** The changing cells typed for a scenario: references on its sheet, up to 32 cells (null otherwise). */
export function parseScenarioCells(doc: SpreadsheetDocument, sheet: Sheet, text: string): { r: number; c: number }[] | null {
  const out: { r: number; c: number }[] = [];
  for (const part of text.split(',')) {
    const hit = parseRangeText(doc, part);
    if (!hit || hit.sheet !== sheet) return null;
    for (let r = hit.range.r0; r <= hit.range.r1; r++) for (let c = hit.range.c0; c <= hit.range.c1; c++) out.push({ r, c });
  }
  return out.length && out.length <= 32 ? out : null;
}

/** The cells of the references in a box (A1:B2,C3), skipping what is not a reference. */
export function cellsOfRangeText(doc: SpreadsheetDocument, text: string): { r: number; c: number }[] {
  const cells: { r: number; c: number }[] = [];
  for (const part of text.split(',')) {
    const hit = parseRangeText(doc, part);
    if (!hit) continue;
    for (let r = hit.range.r0; r <= hit.range.r1; r++) for (let c = hit.range.c0; c <= hit.range.c1; c++) cells.push({ r, c });
  }
  return cells;
}
