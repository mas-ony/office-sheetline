/**
 * Outline symbols in the gutters left of the row headers and above the column headers: level
 * buttons (1, 2, 3...), group brackets and the +/- buttons on summary lines.
 */
import { maxOutlineLevel, outlineGroups, type OutlineAxis, type OutlineGroup } from '../core/model/outline';
import type { Sheet } from '../core/model/sheet';
import type { Rect, ViewTransform } from './viewport';

/** Width of one level column of the gutter (CSS pixels). */
export const OUTLINE_STEP = 16;

/** Gutter sizes for a sheet's outline (0 without groups). */
export function outlineGutters(sheet: Sheet): { w: number; h: number } {
  const rows = maxOutlineLevel(sheet, 'row');
  const cols = maxOutlineLevel(sheet, 'col');
  return { w: rows ? (rows + 1) * OUTLINE_STEP + 4 : 0, h: cols ? (cols + 1) * OUTLINE_STEP + 4 : 0 };
}

export interface OutlineButton {
  rect: Rect;
  /** A level button, or the +/- button of a group. */
  level?: { axis: OutlineAxis; level: number };
  group?: OutlineGroup;
}

export interface OutlineLayout {
  buttons: OutlineButton[];
  /** Bracket lines of expanded groups: [x1, y1, x2, y2] with the end tick. */
  lines: { axis: OutlineAxis; a: number; b: number; at: number; tick: number }[];
}

const BOX = 12;

/** Positions of the outline symbols for the current view. */
export function outlineLayout(vt: ViewTransform): OutlineLayout {
  const { sheet } = vt;
  const gw = vt.v.gutterW ?? 0;
  const gh = vt.v.gutterH ?? 0;
  const { headerW, headerH } = vt.v;
  const out: OutlineLayout = { buttons: [], lines: [] };
  const center = (k: number) => 2 + (k - 1) * OUTLINE_STEP + OUTLINE_STEP / 2;
  if (gw) {
    const max = maxOutlineLevel(sheet, 'row');
    const y = gh + (headerH - gh) / 2;
    for (let k = 1; k <= max + 1; k++) out.buttons.push({ rect: { x: center(k) - BOX / 2, y: y - BOX / 2, w: BOX, h: BOX }, level: { axis: 'row', level: k } });
    for (const g of outlineGroups(sheet, 'row')) {
      const x = center(g.level);
      const s = vt.cellRect(g.summary, 0);
      if (s.h > 0) out.buttons.push({ rect: { x: x - BOX / 2, y: s.y + s.h / 2 - BOX / 2, w: BOX, h: BOX }, group: g });
      if (g.collapsed) continue;
      const first = vt.cellRect(g.start, 0);
      const last = vt.cellRect(g.end, 0);
      const below = g.summary > g.end;
      out.lines.push({ axis: 'row', a: first.y + 2, b: last.y + last.h - 2, at: x, tick: below ? first.y + 2 : last.y + last.h - 2 });
    }
  }
  if (gh) {
    const max = maxOutlineLevel(sheet, 'col');
    const x = gw + (headerW - gw) / 2;
    for (let k = 1; k <= max + 1; k++) out.buttons.push({ rect: { x: x - BOX / 2, y: center(k) - BOX / 2, w: BOX, h: BOX }, level: { axis: 'col', level: k } });
    for (const g of outlineGroups(sheet, 'col')) {
      const y = center(g.level);
      const s = vt.cellRect(0, g.summary);
      if (s.w > 0) out.buttons.push({ rect: { x: s.x + s.w / 2 - BOX / 2, y: y - BOX / 2, w: BOX, h: BOX }, group: g });
      if (g.collapsed) continue;
      const first = vt.cellRect(0, g.start);
      const last = vt.cellRect(0, g.end);
      const right = g.summary > g.end;
      out.lines.push({ axis: 'col', a: first.x + 2, b: last.x + last.w - 2, at: y, tick: right ? first.x + 2 : last.x + last.w - 2 });
    }
  }
  return out;
}

/** The outline button under a point of the gutters. */
export function outlineHit(vt: ViewTransform, x: number, y: number): OutlineButton | null {
  const gw = vt.v.gutterW ?? 0;
  const gh = vt.v.gutterH ?? 0;
  if (!(x < gw || y < gh)) return null;
  const pad = 3;
  for (const b of outlineLayout(vt).buttons) {
    if (x >= b.rect.x - pad && x <= b.rect.x + b.rect.w + pad && y >= b.rect.y - pad && y <= b.rect.y + b.rect.h + pad) return b;
  }
  return null;
}

/** Draws the gutters (called after the headers). */
export function drawOutline(ctx: CanvasRenderingContext2D, vt: ViewTransform, colors: { gutter: string; line: string; text: string; button: string }): void {
  const gw = vt.v.gutterW ?? 0;
  const gh = vt.v.gutterH ?? 0;
  if (!gw && !gh) return;
  const { width, height, headerW, headerH } = vt.v;
  ctx.save();
  ctx.fillStyle = colors.gutter;
  if (gw) ctx.fillRect(0, 0, gw, height);
  if (gh) ctx.fillRect(0, 0, width, gh);
  const layout = outlineLayout(vt);
  ctx.strokeStyle = colors.line;
  ctx.lineWidth = 1;
  // Brackets, clipped to the header strips they follow.
  for (const l of layout.lines) {
    ctx.save();
    ctx.beginPath();
    if (l.axis === 'row') ctx.rect(0, headerH, gw, height - headerH);
    else ctx.rect(headerW, 0, width - headerW, gh);
    ctx.clip();
    ctx.beginPath();
    if (l.axis === 'row') {
      const x = Math.round(l.at) + 0.5;
      ctx.moveTo(x, l.a);
      ctx.lineTo(x, l.b);
      ctx.moveTo(x, Math.round(l.tick) + 0.5);
      ctx.lineTo(x + 5, Math.round(l.tick) + 0.5);
    } else {
      const y = Math.round(l.at) + 0.5;
      ctx.moveTo(l.a, y);
      ctx.lineTo(l.b, y);
      ctx.moveTo(Math.round(l.tick) + 0.5, y);
      ctx.lineTo(Math.round(l.tick) + 0.5, y + 5);
    }
    ctx.stroke();
    ctx.restore();
  }
  ctx.font = '10px Arial, "Arimo", sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  for (const b of layout.buttons) {
    if (b.group) {
      const inRows = b.group.axis === 'row';
      if (inRows ? b.rect.y + b.rect.h < headerH : b.rect.x + b.rect.w < headerW) continue;
    }
    const x = Math.round(b.rect.x) + 0.5;
    const y = Math.round(b.rect.y) + 0.5;
    ctx.fillStyle = colors.button;
    ctx.fillRect(x, y, b.rect.w, b.rect.h);
    ctx.strokeStyle = colors.line;
    ctx.strokeRect(x, y, b.rect.w, b.rect.h);
    ctx.fillStyle = colors.text;
    ctx.strokeStyle = colors.text;
    const cx = x + b.rect.w / 2;
    const cy = y + b.rect.h / 2;
    if (b.level) ctx.fillText(String(b.level.level), cx, cy + 0.5);
    else if (b.group) {
      ctx.beginPath();
      ctx.moveTo(cx - 3, cy);
      ctx.lineTo(cx + 3, cy);
      if (b.group.collapsed) {
        ctx.moveTo(cx, cy - 3);
        ctx.lineTo(cx, cy + 3);
      }
      ctx.stroke();
    }
  }
  // Separators between the gutters and the headers.
  ctx.strokeStyle = colors.line;
  ctx.beginPath();
  if (gw) {
    ctx.moveTo(gw - 0.5, 0);
    ctx.lineTo(gw - 0.5, height);
  }
  if (gh) {
    ctx.moveTo(0, gh - 0.5);
    ctx.lineTo(width, gh - 0.5);
  }
  ctx.stroke();
  ctx.restore();
}
