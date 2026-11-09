/** The status bar's summary of the selection and its zoom slider. */
import type { SpreadsheetDocument } from '../core/document';
import { formatValue } from '../core/format/display';

export type StatusStat = 'average' | 'count' | 'numCount' | 'min' | 'max' | 'sum';

export const STATUS_STAT_LABELS: Record<StatusStat, string> = { average: 'Average', count: 'Count', numCount: 'Numerical Count', min: 'Min', max: 'Max', sum: 'Sum' };

/** Average, count, sum... of the selected cells (hidden rows of filters skipped); null under two values. */
export function selectionStats(doc: SpreadsheetDocument): Record<StatusStat, string | null> | null {
  const sheet = doc.sheet;
  let count = 0;
  let numCount = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  let fmt = 'General';
  let visited = 0;
  for (const range of doc.selection.ranges) {
    const b = doc.bounded(range, sheet);
    if (!b) continue;
    sheet.forEachCellIn(b, (r, _c, cell) => {
      if (visited++ > 2_000_000) return;
      if (sheet.filteredRows.has(r)) return;
      if (cell.v === null || cell.v === '') return;
      count++;
      if (typeof cell.v === 'number') {
        if (!numCount) fmt = doc.styleAt(r, _c).numFmt;
        numCount++;
        sum += cell.v;
        min = Math.min(min, cell.v);
        max = Math.max(max, cell.v);
      }
    });
  }
  if (count < 2 && !numCount) return null;
  const f = (v: number) => formatValue(v, /%|0|#/.test(fmt) && !/[dmyhs]/i.test(fmt) ? fmt : 'General', doc.wb.date1904).text;
  return {
    average: numCount ? f(sum / numCount) : null,
    count: String(count),
    numCount: String(numCount),
    min: numCount ? f(min) : null,
    max: numCount ? f(max) : null,
    sum: numCount ? f(sum) : null,
  };
}

/** Slider position (0..100) for a zoom between 10% and 400% with 100% in the middle. */
export function zoomToSlider(z: number): number {
  return z <= 1 ? ((z - 0.1) / 0.9) * 50 : 50 + ((z - 1) / 3) * 50;
}

export function sliderToZoom(v: number): number {
  const z = v <= 50 ? 0.1 + (v / 50) * 0.9 : 1 + ((v - 50) / 50) * 3;
  return Math.round(z * 100) / 100;
}
