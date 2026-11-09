/** Floating objects: inserting shapes and text boxes, arranging, sizing and the selected object. */
import { SheetLayout } from '../core/layout';
import { DEFAULT_SHAPE_FILL, DEFAULT_SHAPE_LINE, type SheetChart, type SheetShape } from '../core/model/drawing';
import type { SheetImage } from '../core/model/sheet';
import { LINE_PRESETS, type GalleryItem } from '../grid/shapes';
import type { UIContextValue } from './context';
import { commit } from './commands';

type Ctx = UIContextValue;

/** Where new objects go: at the active cell, or near the top-left of the view when it is off screen. */
function placement(ctx: Ctx, w: number, h: number): { x: number; y: number; width: number; height: number } {
  const sheet = ctx.doc.sheet;
  const layout = SheetLayout.of(sheet);
  const { r, c } = ctx.doc.selection.active;
  return { x: layout.colX(c), y: layout.rowY(r), width: w, height: h };
}

export function insertShape(ctx: Ctx, it: GalleryItem): SheetShape | null {
  commit(ctx);
  const blocked = ctx.doc.blockedBy('objects');
  if (blocked) {
    ctx.notify(blocked);
    return null;
  }
  const line = LINE_PRESETS.has(it.preset);
  const w = ctx.phone ? Math.min(240, Math.round(window.innerWidth * 0.6)) : 240;
  const h = line ? Math.round(w / 3) : Math.round((w * 2) / 3);
  const n = ctx.doc.sheet.shapes.length + 1;
  const shape = ctx.doc.addShape(
    {
      name: `${it.label.replace(/\b\w/g, (m) => m.toUpperCase())} ${n}`,
      preset: it.preset,
      fill: line ? null : DEFAULT_SHAPE_FILL,
      line: { color: DEFAULT_SHAPE_LINE, width: line ? 1.5 : 1 },
      fontColor: '#FFFFFF',
      arrowStart: it.arrowStart || undefined,
      arrowEnd: it.arrowEnd || undefined,
    },
    placement(ctx, w, h),
  );
  if (shape) {
    ctx.grid?.scrollIntoView(shape.from.r, shape.from.c);
    ctx.selectObject(shape.id);
    ctx.showTab('shape-format');
  }
  return shape;
}

export function insertTextBox(ctx: Ctx): SheetShape | null {
  commit(ctx);
  const blocked = ctx.doc.blockedBy('objects');
  if (blocked) {
    ctx.notify(blocked);
    return null;
  }
  const n = ctx.doc.sheet.shapes.length + 1;
  const shape = ctx.doc.addShape(
    { name: `TextBox ${n}`, preset: 'rect', fill: '#FFFFFF', line: { color: '#BFBFBF', width: 1 }, fontColor: '#000000', text: '', textBox: true },
    placement(ctx, 180, 60),
  );
  if (shape) {
    ctx.selectObject(shape.id);
    ctx.openDialog({ kind: 'shape-text', id: shape.id });
  }
  return shape;
}

export type SelectedObject = { kind: 'chart'; obj: SheetChart } | { kind: 'shape'; obj: SheetShape } | { kind: 'image'; obj: SheetImage };

export function selected(ctx: Ctx): SelectedObject | null {
  return ctx.selectedObject ? ctx.doc.objectById(ctx.selectedObject) : null;
}

/** Pixel rectangle of an object at 100% zoom. */
export function objectRect(ctx: Ctx, id: string): { x: number; y: number; w: number; h: number } | null {
  const hit = ctx.doc.objectById(id);
  if (!hit) return null;
  const layout = SheetLayout.of(ctx.doc.sheet);
  if (hit.kind === 'image') return layout.imageRect(hit.obj);
  return layout.anchorRect(hit.obj.from, hit.obj.to);
}

export function setObjectSize(ctx: Ctx, id: string, w: number, h: number): void {
  const r = objectRect(ctx, id);
  if (!r || !(w > 0) || !(h > 0)) return;
  ctx.doc.setObjectRect(id, { x: r.x, y: r.y, width: Math.round(w), height: Math.round(h) });
}

/** Inches (two decimals) for a length in pixels at 96 dpi, as shown in the size fields. */
export function pxToInches(px: number): string {
  return String(Math.round((px / 96) * 100) / 100);
}

/** Align selected object to the sheet's visible area edges (single selection: relative to the view). */
export function alignObject(ctx: Ctx, how: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'): void {
  const id = ctx.selectedObject;
  if (!id || !ctx.grid) return;
  const r = objectRect(ctx, id);
  if (!r) return;
  const vt = ctx.grid.transform();
  const z = vt.zoom;
  const left = vt.sheetX(vt.mainLeft);
  const top = vt.sheetY(vt.mainTop);
  const right = left + (vt.v.width - vt.mainLeft) / z;
  const bottom = top + (vt.v.height - vt.mainTop) / z;
  let { x, y } = r;
  if (how === 'left') x = left;
  if (how === 'center') x = (left + right - r.w) / 2;
  if (how === 'right') x = right - r.w;
  if (how === 'top') y = top;
  if (how === 'middle') y = (top + bottom - r.h) / 2;
  if (how === 'bottom') y = bottom - r.h;
  ctx.doc.setObjectRect(id, { x: Math.max(0, x), y: Math.max(0, y), width: r.w, height: r.h });
}

export function rotateShape(ctx: Ctx, how: 'right' | 'left' | 'flipV' | 'flipH'): void {
  const sel = selected(ctx);
  if (sel?.kind !== 'shape') return;
  const s = sel.obj;
  if (how === 'right') ctx.doc.updateShape(s.id, { rotation: (((s.rotation ?? 0) + 90) % 360) || undefined });
  else if (how === 'left') ctx.doc.updateShape(s.id, { rotation: (((s.rotation ?? 0) + 270) % 360) || undefined });
  else if (how === 'flipV') ctx.doc.updateShape(s.id, { flipV: !s.flipV || undefined });
  else ctx.doc.updateShape(s.id, { flipH: !s.flipH || undefined });
}

/** Shape style presets of the Shape Styles gallery: theme colour fills with darker outlines. */
export const SHAPE_STYLES: { label: string; fill: string | null; line: string; font: string }[] = [
  { label: 'Colored Outline - Blue', fill: '#FFFFFF', line: '#5B9BD5', font: '#000000' },
  { label: 'Colored Outline - Orange', fill: '#FFFFFF', line: '#ED7D31', font: '#000000' },
  { label: 'Colored Outline - Gray', fill: '#FFFFFF', line: '#A5A5A5', font: '#000000' },
  { label: 'Colored Outline - Gold', fill: '#FFFFFF', line: '#FFC000', font: '#000000' },
  { label: 'Colored Outline - Green', fill: '#FFFFFF', line: '#70AD47', font: '#000000' },
  { label: 'Colored Fill - Blue', fill: '#5B9BD5', line: '#41719C', font: '#FFFFFF' },
  { label: 'Colored Fill - Orange', fill: '#ED7D31', line: '#AE5A21', font: '#FFFFFF' },
  { label: 'Colored Fill - Gray', fill: '#A5A5A5', line: '#787878', font: '#FFFFFF' },
  { label: 'Colored Fill - Gold', fill: '#FFC000', line: '#BC8C00', font: '#FFFFFF' },
  { label: 'Colored Fill - Green', fill: '#70AD47', line: '#507E32', font: '#FFFFFF' },
  { label: 'Light 1 Fill - Blue', fill: '#DEEBF7', line: '#5B9BD5', font: '#000000' },
  { label: 'Light 1 Fill - Orange', fill: '#FBE5D6', line: '#ED7D31', font: '#000000' },
  { label: 'Light 1 Fill - Gray', fill: '#EDEDED', line: '#A5A5A5', font: '#000000' },
  { label: 'Light 1 Fill - Gold', fill: '#FFF2CC', line: '#FFC000', font: '#000000' },
  { label: 'Light 1 Fill - Green', fill: '#E2F0D9', line: '#70AD47', font: '#000000' },
  { label: 'Intense Effect - Dark', fill: '#404040', line: '#262626', font: '#FFFFFF' },
];
