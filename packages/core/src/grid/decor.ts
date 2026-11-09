/**
 * Drawing extras on top of cell formatting: table styles (below direct formatting), conditional
 * format icons and data bars. Shared by the grid and the printed pages.
 */
import type { SpreadsheetDocument } from '../core/document';
import type { Sheet } from '../core/model/sheet';
import type { CellStyle } from '../core/model/style';
import { tableCellDecor, type CellDecor } from '../core/tableStyles';
import type { CfIcon, CfResult } from './condFormat';
import type { Rect } from './viewport';

/**
 * The part of the canvas API that cell decorations (icons, data bars, sparklines) paint with. The
 * printed pages record the same calls as SVG.
 */
export type Paint2D = Pick<
  CanvasRenderingContext2D,
  | 'save' | 'restore' | 'translate' | 'rotate' | 'beginPath' | 'moveTo' | 'lineTo' | 'closePath' | 'arc' | 'rect'
  | 'fill' | 'stroke' | 'fillRect' | 'strokeRect' | 'clip' | 'setLineDash' | 'createLinearGradient'
  | 'fillStyle' | 'strokeStyle' | 'lineWidth' | 'lineCap' | 'lineJoin'
>;

/** Table style formatting of the cells of a sheet, or null when the sheet has no styled tables. */
export function tableDecorFor(doc: SpreadsheetDocument, sheet: Sheet): ((r: number, c: number) => CellDecor | null) | null {
  const tables = sheet.tables.filter((t) => t.style.name);
  if (!tables.length) return null;
  const theme = doc.wb.theme;
  return (r, c) => {
    for (const t of tables) {
      if (r >= t.ref.r0 && r <= t.ref.r1 && c >= t.ref.c0 && c <= t.ref.c1) return tableCellDecor(t, r, c, theme);
    }
    return null;
  };
}

/** Whether a cell's font colour was set explicitly (the default text colour lets table styles show). */
export function explicitFontColor(style: CellStyle): boolean {
  const c = style.font.color;
  if (!c) return false;
  if ('auto' in c) return false;
  if ('theme' in c && c.theme === 1 && !c.tint) return false;
  if ('indexed' in c && (c.indexed === 64 || c.indexed === 8)) return false;
  return true;
}

const GREEN = '#21A350';
const YELLOW = '#F4B400';
const RED = '#D93B3B';
const GRAY = '#808080';
const BLACK = '#262626';

type IconSpec =
  | { k: 'arrow'; dir: 'up' | 'down' | 'right' | 'upRight' | 'downRight'; color: string }
  | { k: 'circle'; color: string; rim?: boolean }
  | { k: 'triangle'; dir: 'up' | 'down'; color: string }
  | { k: 'dash'; color: string }
  | { k: 'diamond'; color: string }
  | { k: 'symbol'; sym: 'check' | 'bang' | 'cross'; color: string; circled: boolean }
  | { k: 'flag'; color: string }
  | { k: 'star'; fill: 0 | 0.5 | 1 }
  | { k: 'bars'; filled: number; total: number }
  | { k: 'quarters'; filled: number }
  | { k: 'boxes'; filled: number };

/** Icons of each set, lowest value first. */
export function iconSpec(icon: CfIcon): IconSpec {
  const { set, index: i } = icon;
  switch (set) {
    case '3Arrows':
      return { k: 'arrow', dir: (['down', 'right', 'up'] as const)[i], color: [RED, YELLOW, GREEN][i] };
    case '3ArrowsGray':
      return { k: 'arrow', dir: (['down', 'right', 'up'] as const)[i], color: GRAY };
    case '4Arrows':
      return { k: 'arrow', dir: (['down', 'downRight', 'upRight', 'up'] as const)[i], color: [RED, YELLOW, YELLOW, GREEN][i] };
    case '4ArrowsGray':
      return { k: 'arrow', dir: (['down', 'downRight', 'upRight', 'up'] as const)[i], color: GRAY };
    case '5Arrows':
      return { k: 'arrow', dir: (['down', 'downRight', 'right', 'upRight', 'up'] as const)[i], color: [RED, YELLOW, YELLOW, YELLOW, GREEN][i] };
    case '5ArrowsGray':
      return { k: 'arrow', dir: (['down', 'downRight', 'right', 'upRight', 'up'] as const)[i], color: GRAY };
    case '3Triangles':
      return i === 1 ? { k: 'dash', color: YELLOW } : { k: 'triangle', dir: i === 2 ? 'up' : 'down', color: i === 2 ? GREEN : RED };
    case '3TrafficLights1':
      return { k: 'circle', color: [RED, YELLOW, GREEN][i] };
    case '3TrafficLights2':
      return { k: 'circle', color: [RED, YELLOW, GREEN][i], rim: true };
    case '4TrafficLights':
      return { k: 'circle', color: [BLACK, RED, YELLOW, GREEN][i] };
    case '4RedToBlack':
      return { k: 'circle', color: [BLACK, GRAY, '#F2A7A7', RED][i] };
    case '3Signs':
      return i === 0 ? { k: 'diamond', color: RED } : i === 1 ? { k: 'triangle', dir: 'up', color: YELLOW } : { k: 'circle', color: GREEN };
    case '3Symbols':
      return { k: 'symbol', sym: (['cross', 'bang', 'check'] as const)[i], color: [RED, YELLOW, GREEN][i], circled: true };
    case '3Symbols2':
      return { k: 'symbol', sym: (['cross', 'bang', 'check'] as const)[i], color: [RED, YELLOW, GREEN][i], circled: false };
    case '3Flags':
      return { k: 'flag', color: [RED, YELLOW, GREEN][i] };
    case '3Stars':
      return { k: 'star', fill: ([0, 0.5, 1] as const)[i] };
    case '4Rating':
      return { k: 'bars', filled: i + 1, total: 4 };
    case '5Rating':
      return { k: 'bars', filled: i, total: 4 };
    case '5Quarters':
      return { k: 'quarters', filled: i };
    case '5Boxes':
      return { k: 'boxes', filled: i };
  }
}

/** Draws a conditional format icon in a square of size s at (x, y). */
export function drawIcon(ctx: Paint2D, icon: CfIcon, x: number, y: number, s: number): void {
  const spec = iconSpec(icon);
  const cx = x + s / 2;
  const cy = y + s / 2;
  ctx.save();
  ctx.lineJoin = 'round';
  switch (spec.k) {
    case 'arrow': {
      ctx.translate(cx, cy);
      const angle = { up: -90, upRight: -45, right: 0, downRight: 45, down: 90 }[spec.dir];
      ctx.rotate((angle * Math.PI) / 180);
      ctx.fillStyle = spec.color;
      const L = s * 0.45;
      ctx.beginPath();
      ctx.moveTo(L, 0);
      ctx.lineTo(0, -L * 0.85);
      ctx.lineTo(0, -L * 0.35);
      ctx.lineTo(-L, -L * 0.35);
      ctx.lineTo(-L, L * 0.35);
      ctx.lineTo(0, L * 0.35);
      ctx.lineTo(0, L * 0.85);
      ctx.closePath();
      ctx.fill();
      break;
    }
    case 'circle':
      ctx.fillStyle = spec.color;
      ctx.beginPath();
      ctx.arc(cx, cy, s * 0.4, 0, Math.PI * 2);
      ctx.fill();
      if (spec.rim) {
        ctx.strokeStyle = '#404040';
        ctx.lineWidth = Math.max(1, s * 0.08);
        ctx.stroke();
      }
      break;
    case 'triangle':
      ctx.fillStyle = spec.color;
      ctx.beginPath();
      if (spec.dir === 'up') {
        ctx.moveTo(cx, y + s * 0.12);
        ctx.lineTo(x + s * 0.9, y + s * 0.85);
        ctx.lineTo(x + s * 0.1, y + s * 0.85);
      } else {
        ctx.moveTo(cx, y + s * 0.88);
        ctx.lineTo(x + s * 0.9, y + s * 0.15);
        ctx.lineTo(x + s * 0.1, y + s * 0.15);
      }
      ctx.closePath();
      ctx.fill();
      break;
    case 'dash':
      ctx.fillStyle = spec.color;
      ctx.fillRect(x + s * 0.15, cy - s * 0.12, s * 0.7, s * 0.24);
      break;
    case 'diamond':
      ctx.fillStyle = spec.color;
      ctx.beginPath();
      ctx.moveTo(cx, y + s * 0.08);
      ctx.lineTo(x + s * 0.92, cy);
      ctx.lineTo(cx, y + s * 0.92);
      ctx.lineTo(x + s * 0.08, cy);
      ctx.closePath();
      ctx.fill();
      break;
    case 'symbol': {
      const fg = spec.circled ? '#ffffff' : spec.color;
      if (spec.circled) {
        ctx.fillStyle = spec.color;
        ctx.beginPath();
        ctx.arc(cx, cy, s * 0.45, 0, Math.PI * 2);
        ctx.fill();
      }
      ctx.strokeStyle = fg;
      ctx.fillStyle = fg;
      ctx.lineWidth = Math.max(1.5, s * 0.14);
      ctx.lineCap = 'round';
      ctx.beginPath();
      const k = spec.circled ? 0.22 : 0.32;
      if (spec.sym === 'check') {
        ctx.moveTo(cx - s * k, cy);
        ctx.lineTo(cx - s * k * 0.2, cy + s * k * 0.8);
        ctx.lineTo(cx + s * k, cy - s * k * 0.8);
        ctx.stroke();
      } else if (spec.sym === 'cross') {
        ctx.moveTo(cx - s * k, cy - s * k);
        ctx.lineTo(cx + s * k, cy + s * k);
        ctx.moveTo(cx + s * k, cy - s * k);
        ctx.lineTo(cx - s * k, cy + s * k);
        ctx.stroke();
      } else {
        ctx.moveTo(cx, cy - s * k);
        ctx.lineTo(cx, cy + s * k * 0.3);
        ctx.stroke();
        ctx.beginPath();
        ctx.arc(cx, cy + s * k * 0.85, ctx.lineWidth * 0.55, 0, Math.PI * 2);
        ctx.fill();
      }
      break;
    }
    case 'flag':
      ctx.strokeStyle = '#404040';
      ctx.lineWidth = Math.max(1, s * 0.08);
      ctx.beginPath();
      ctx.moveTo(x + s * 0.25, y + s * 0.1);
      ctx.lineTo(x + s * 0.25, y + s * 0.92);
      ctx.stroke();
      ctx.fillStyle = spec.color;
      ctx.beginPath();
      ctx.moveTo(x + s * 0.28, y + s * 0.12);
      ctx.lineTo(x + s * 0.88, y + s * 0.32);
      ctx.lineTo(x + s * 0.28, y + s * 0.55);
      ctx.closePath();
      ctx.fill();
      break;
    case 'star': {
      const pts: [number, number][] = [];
      for (let k = 0; k < 10; k++) {
        const rr = k % 2 ? s * 0.2 : s * 0.46;
        const a = (Math.PI / 5) * k - Math.PI / 2;
        pts.push([cx + rr * Math.cos(a), cy + rr * Math.sin(a)]);
      }
      const path = () => {
        ctx.beginPath();
        pts.forEach(([px, py], k) => (k ? ctx.lineTo(px, py) : ctx.moveTo(px, py)));
        ctx.closePath();
      };
      ctx.fillStyle = '#ffffff';
      path();
      ctx.fill();
      if (spec.fill) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(x, y, spec.fill === 1 ? s : s / 2, s);
        ctx.clip();
        ctx.fillStyle = '#F2B632';
        path();
        ctx.fill();
        ctx.restore();
      }
      ctx.strokeStyle = '#C79100';
      ctx.lineWidth = 1;
      path();
      ctx.stroke();
      break;
    }
    case 'bars':
      for (let k = 0; k < spec.total; k++) {
        const bh = (s * 0.8 * (k + 1)) / spec.total;
        ctx.fillStyle = k < spec.filled ? '#2B6CB0' : '#C9D3DE';
        ctx.fillRect(x + s * 0.08 + (k * s * 0.84) / spec.total, y + s * 0.9 - bh, (s * 0.84) / spec.total - 1, bh);
      }
      break;
    case 'quarters':
      ctx.fillStyle = '#ffffff';
      ctx.strokeStyle = '#2B2B2B';
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.arc(cx, cy, s * 0.42, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
      if (spec.filled) {
        ctx.fillStyle = '#2B2B2B';
        ctx.beginPath();
        ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, s * 0.42, -Math.PI / 2, -Math.PI / 2 + (Math.PI / 2) * spec.filled);
        ctx.closePath();
        ctx.fill();
      }
      break;
    case 'boxes': {
      const b = s * 0.4;
      [[0, 0], [1, 0], [0, 1], [1, 1]].forEach(([bx, by], k) => {
        ctx.fillStyle = k < spec.filled ? '#2B6CB0' : '#C9D3DE';
        ctx.fillRect(x + s * 0.08 + bx * (b + 1.5), y + s * 0.08 + by * (b + 1.5), b, b);
      });
      break;
    }
  }
  ctx.restore();
}

export type { CellDecor };

/** Half-pixel aligned position of a 1-pixel line. */
function crisp(x: number, dpr: number): number {
  return (Math.round(x * dpr) + 0.5) / dpr;
}

/** Draws a conditional format data bar in a cell rectangle. */
export function drawDataBar(ctx: Paint2D, rect: Rect, bar: NonNullable<CfResult['bar']>, dpr = 1): void {
  const pad = 2;
  const inner = rect.w - 2 * pad;
  const h = rect.h - 2 * pad;
  const ratio = Math.min(1, Math.max(0, bar.ratio));
  let x0 = rect.x + pad;
  const w = inner * ratio;
  if (bar.axis !== undefined) {
    const ax = rect.x + pad + inner * bar.axis;
    if (bar.negative) x0 = ax - w;
    else x0 = ax;
    ctx.save();
    ctx.strokeStyle = '#000000';
    ctx.setLineDash([2, 2]);
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(crisp(ax, dpr), rect.y + pad);
    ctx.lineTo(crisp(ax, dpr), rect.y + pad + h);
    ctx.stroke();
    ctx.restore();
  }
  if (w <= 0) return;
  if (bar.gradient) {
    const g = bar.negative ? ctx.createLinearGradient(x0 + w, 0, x0, 0) : ctx.createLinearGradient(x0, 0, x0 + w, 0);
    g.addColorStop(0, bar.color);
    g.addColorStop(1, '#ffffff');
    ctx.fillStyle = g;
  } else ctx.fillStyle = bar.color;
  ctx.fillRect(x0, rect.y + pad, w, h);
  if (bar.border) {
    ctx.strokeStyle = bar.border;
    ctx.lineWidth = 1;
    ctx.strokeRect(Math.round(x0) + 0.5, Math.round(rect.y + pad) + 0.5, Math.max(0, Math.round(w) - 1), Math.max(0, Math.round(h) - 1));
  }
}
