/** Drawing sparklines (line, column, win/loss) inside their cells. */
import type { SpreadsheetDocument } from '../core/document';
import type { Sheet } from '../core/model/sheet';
import { resolveColor, type ColorSpec, type Theme } from '../core/model/color';
import type { SparklineGroup } from '../core/model/sparkline';
import { readNumbers } from '../core/chartData';
import type { Rect } from './viewport';
import type { Paint2D } from './decor';

const hex = (spec: ColorSpec, theme: Theme, fallback: string) => `#${resolveColor(spec, theme, fallback)}`;

/** Values of a sparkline with empty cells handled as the group asks. */
export function sparklineValues(doc: SpreadsheetDocument, group: SparklineGroup, ref: string, sheet: Sheet): (number | null)[] {
  const values = readNumbers(doc, ref, sheet);
  const list = group.displayEmptyCellsAs === 'zero' ? values.map((v) => v ?? 0) : values;
  return group.rightToLeft ? [...list].reverse() : list;
}

function bounds(values: (number | null)[]): { min: number; max: number } | null {
  let min = Infinity;
  let max = -Infinity;
  for (const v of values) {
    if (v === null) continue;
    min = Math.min(min, v);
    max = Math.max(max, v);
  }
  return Number.isFinite(min) ? { min, max } : null;
}

/**
 * Axis bounds of a group's sparklines: a function from a sparkline's values to the minimum and
 * maximum it is drawn with (shared group bounds are computed once).
 */
export function sparklineScaler(doc: SpreadsheetDocument, sheet: Sheet, g: SparklineGroup): (values: (number | null)[]) => { min: number; max: number } | null {
  let groupBounds: { min: number; max: number } | null | undefined;
  return (values) => {
    const own = bounds(values);
    if (!own) return null;
    if ((g.minAxisType === 'group' || g.maxAxisType === 'group') && groupBounds === undefined) {
      groupBounds = bounds(g.sparklines.flatMap((s) => sparklineValues(doc, g, s.ref, sheet)));
    }
    let min = g.minAxisType === 'custom' && g.manualMin !== undefined ? g.manualMin : g.minAxisType === 'group' && groupBounds ? groupBounds.min : own.min;
    let max = g.maxAxisType === 'custom' && g.manualMax !== undefined ? g.manualMax : g.maxAxisType === 'group' && groupBounds ? groupBounds.max : own.max;
    if (g.type === 'column') {
      // Columns grow from zero when all values have one sign.
      if (g.minAxisType === 'individual' && min > 0) min = 0;
      if (g.maxAxisType === 'individual' && max < 0) max = 0;
    }
    if (max === min) {
      max += 1;
      min -= g.type === 'line' ? 1 : 0;
    }
    return { min, max };
  };
}

/** Draws the sparklines of the visible cells; `rectOf` gives a cell's screen rectangle. */
export function drawSparklines(ctx: Paint2D, doc: SpreadsheetDocument, sheet: Sheet, visible: { r0: number; r1: number; c0: number; c1: number }, rectOf: (r: number, c: number) => Rect, zoom: number): void {
  const theme = doc.wb.theme;
  for (const g of sheet.sparklineGroups) {
    const shown = g.sparklines.filter((s) => s.r >= visible.r0 && s.r <= visible.r1 && s.c >= visible.c0 && s.c <= visible.c1);
    if (!shown.length) continue;
    const scale = sparklineScaler(doc, sheet, g);
    for (const s of shown) {
      const rect = rectOf(s.r, s.c);
      if (rect.w < 6 || rect.h < 6) continue;
      const values = sparklineValues(doc, g, s.ref, sheet);
      const range = scale(values);
      if (range) paintSparkline(ctx, g, values, rect, range.min, range.max, theme, zoom);
    }
  }
}

/** Draws one sparkline in a cell rectangle with its axis bounds. */
export function paintSparkline(ctx: Paint2D, g: SparklineGroup, values: (number | null)[], rect: Rect, min: number, max: number, theme: Theme, zoom: number): void {
  const pad = Math.max(2, 3 * zoom);
  const x0 = rect.x + pad;
  const y0 = rect.y + pad;
  const w = rect.w - 2 * pad;
  const h = rect.h - 2 * pad;
  const n = values.length;
  if (n === 0 || w <= 0 || h <= 0) return;
  const Y = (v: number) => y0 + h - ((Math.min(max, Math.max(min, v)) - min) / (max - min)) * h;
  const real = values.map((v, i) => ({ v, i })).filter((p): p is { v: number; i: number } => p.v !== null);
  const hiV = Math.max(...real.map((p) => p.v));
  const loV = Math.min(...real.map((p) => p.v));
  const firstI = real[0]?.i;
  const lastI = real[real.length - 1]?.i;
  const pointColor = (v: number, i: number): string | null => {
    if (g.high && v === hiV) return hex(g.colors.colorHigh, theme, '5B9BD5');
    if (g.low && v === loV) return hex(g.colors.colorLow, theme, '5B9BD5');
    if (g.first && i === firstI) return hex(g.colors.colorFirst, theme, '9DC3E6');
    if (g.last && i === lastI) return hex(g.colors.colorLast, theme, '9DC3E6');
    if (g.negative && v < 0) return hex(g.colors.colorNegative, theme, 'ED7D31');
    return null;
  };
  const series = hex(g.colors.colorSeries, theme, '2E75B6');
  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.w, rect.h);
  ctx.clip();
  if (g.displayXAxis && min < 0 && max > 0) {
    ctx.strokeStyle = hex(g.colors.colorAxis, theme, '000000');
    ctx.lineWidth = 1;
    ctx.beginPath();
    const ay = Math.round(Y(0)) + 0.5;
    ctx.moveTo(x0, ay);
    ctx.lineTo(x0 + w, ay);
    ctx.stroke();
  }
  if (g.type === 'line') {
    const X = (i: number) => (n === 1 ? x0 + w / 2 : x0 + (i / (n - 1)) * w);
    ctx.strokeStyle = series;
    ctx.lineWidth = Math.max(0.75, (g.lineWeight * 4) / 3) * zoom;
    ctx.lineJoin = 'round';
    ctx.beginPath();
    let pen = false;
    values.forEach((v, i) => {
      if (v === null) {
        if (g.displayEmptyCellsAs === 'gap') pen = false;
        return;
      }
      if (pen) ctx.lineTo(X(i), Y(v));
      else ctx.moveTo(X(i), Y(v));
      pen = true;
    });
    ctx.stroke();
    const r = Math.max(1.5, 2 * zoom);
    for (const p of real) {
      const special = pointColor(p.v, p.i);
      if (!special && !g.markers) continue;
      ctx.fillStyle = special ?? hex(g.colors.colorMarkers, theme, '2E75B6');
      ctx.beginPath();
      ctx.arc(X(p.i), Y(p.v), r, 0, Math.PI * 2);
      ctx.fill();
    }
  } else {
    const band = w / n;
    const bw = Math.max(1, band * 0.75);
    for (const p of real) {
      const x = x0 + band * p.i + (band - bw) / 2;
      let top: number;
      let bottom: number;
      if (g.type === 'stacked') {
        // Win/loss: equal bars above or below the middle.
        const mid = y0 + h / 2;
        if (p.v === 0) continue;
        top = p.v > 0 ? y0 : mid;
        bottom = p.v > 0 ? mid : y0 + h;
      } else {
        const zero = Y(Math.max(min, Math.min(max, 0)));
        const y = Y(p.v);
        top = Math.min(zero, y);
        bottom = Math.max(zero, y);
        if (bottom - top < 1) bottom = top + 1;
      }
      ctx.fillStyle = pointColor(p.v, p.i) ?? series;
      ctx.fillRect(x, top, bw, bottom - top);
    }
  }
  ctx.restore();
}
