/**
 * Editor commands shared by the ribbon, the context menus, the command palette and keyboard
 * shortcuts.
 */
import { MAX_COLS, MAX_ROWS, isFullColumns, isFullRows, makeRange, rangesIntersect, type Range } from '../core/refs';
import { filterRegion } from '../core/filter';
import type { StylePatch, UnderlineType } from '../core/model/style';
import type { ColorSpec } from '../core/model/color';
import { formatInfo } from '../core/format/display';
import type { UIContextValue } from './context';
import { copySelection, pasteClipboard, type PasteMode, type PasteOptions } from './clipboardActions';

type Ctx = UIContextValue;

export function commit(ctx: Ctx): void {
  if (ctx.grid?.isEditing()) ctx.grid.commitEdit();
}

// -----------------------------------------------------------------------------------------------
// Undo, clipboard

export function undo(ctx: Ctx): void {
  if (ctx.grid?.isEditing()) {
    ctx.grid.cancelEdit();
    return;
  }
  ctx.doc.undo();
}

export function redo(ctx: Ctx): void {
  commit(ctx);
  ctx.doc.redo();
}

export function cut(ctx: Ctx): void {
  copySelection(ctx.doc, ctx.grid, true);
}

export function copy(ctx: Ctx): void {
  copySelection(ctx.doc, ctx.grid, false);
}

export async function paste(ctx: Ctx, mode: PasteMode | PasteOptions = 'all'): Promise<void> {
  commit(ctx);
  const ok = await pasteClipboard(ctx.doc, mode);
  if (!ok) ctx.notify(window.isSecureContext ? 'Allow clipboard access, or press Ctrl+V to paste.' : 'Press Ctrl+V to paste from other apps (this page cannot read the clipboard directly).');
  ctx.grid?.focusIdle();
}

// -----------------------------------------------------------------------------------------------
// Font

export function format(ctx: Ctx, patch: StylePatch, label = 'Format'): void {
  commit(ctx);
  ctx.doc.format(patch, undefined, undefined, label);
}

export function toggleFont(ctx: Ctx, prop: 'bold' | 'italic' | 'underline' | 'strike'): void {
  commit(ctx);
  ctx.doc.toggleFont(prop);
}

export function toggleUnderline(ctx: Ctx, kind: UnderlineType): void {
  const current = ctx.doc.styleAt(ctx.doc.selection.active.r, ctx.doc.selection.active.c).font.underline;
  format(ctx, { font: { underline: current === kind ? null : kind } }, 'Underline');
}

export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

export const FONT_NAMES = [
  'Arial', 'Arial Black', 'Calibri', 'Calibri Light', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas',
  'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Impact', 'Lucida Console',
  'Lucida Sans Unicode', 'Microsoft Sans Serif', 'Palatino Linotype', 'Segoe UI', 'Tahoma', 'Times New Roman',
  'Trebuchet MS', 'Verdana',
];

export function setFontName(ctx: Ctx, name: string): void {
  format(ctx, { font: { name, scheme: null } }, 'Font');
}

export function setFontSize(ctx: Ctx, size: number): void {
  if (!Number.isFinite(size) || size < 1 || size > 409) {
    ctx.notify('The font size must be a number between 1 and 409.');
    return;
  }
  format(ctx, { font: { size: Math.round(size * 2) / 2 } }, 'Font size');
}

export function growFont(ctx: Ctx, dir: 1 | -1): void {
  const cur = ctx.state.format.fontSize;
  const next = dir > 0 ? FONT_SIZES.find((s) => s > cur) ?? Math.min(409, cur + 10) : [...FONT_SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
  setFontSize(ctx, next);
}

export function setFontColor(ctx: Ctx, color: ColorSpec | null): void {
  format(ctx, { font: { color: color ?? { theme: 1 } } }, 'Font color');
}

export function setFillColor(ctx: Ctx, color: ColorSpec | null): void {
  if (color && 'rgb' in color && color.rgb === 'FFFFFF00') {
    // The fill button before any pick: Excel's default yellow.
    format(ctx, { fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } } }, 'Fill color');
    return;
  }
  format(ctx, { fill: color ? { pattern: 'solid', fg: color, bg: { indexed: 64 } } : null }, 'Fill color');
}

// -----------------------------------------------------------------------------------------------
// Borders

export type BorderPreset =
  | 'bottom' | 'top' | 'left' | 'right' | 'none' | 'all' | 'outside' | 'thick-box' | 'inside' | 'horizontal' | 'vertical'
  | 'bottom-double' | 'thick-bottom' | 'top-bottom' | 'top-thick-bottom' | 'top-double-bottom';

export function applyBorder(ctx: Ctx, preset: BorderPreset): void {
  commit(ctx);
  const { doc } = ctx;
  doc.transact('Borders', () => {
    switch (preset) {
      case 'thick-box':
        doc.setBorder('outside', { style: 'medium' });
        break;
      case 'bottom-double':
        doc.setBorder('bottom', { style: 'double' });
        break;
      case 'thick-bottom':
        doc.setBorder('bottom', { style: 'medium' });
        break;
      case 'top-bottom':
        doc.setBorder('top');
        doc.setBorder('bottom');
        break;
      case 'top-thick-bottom':
        doc.setBorder('top');
        doc.setBorder('bottom', { style: 'medium' });
        break;
      case 'top-double-bottom':
        doc.setBorder('top');
        doc.setBorder('bottom', { style: 'double' });
        break;
      default:
        doc.setBorder(preset, preset === 'none' ? null : { style: 'thin' });
    }
  });
}

// -----------------------------------------------------------------------------------------------
// Alignment

export function setHAlign(ctx: Ctx, h: 'left' | 'center' | 'right' | 'justify'): void {
  const cur = ctx.state.format.hAlign;
  format(ctx, { align: { horizontal: cur === h ? null : h } }, 'Align');
}

export function setVAlign(ctx: Ctx, v: 'top' | 'center' | 'bottom'): void {
  format(ctx, { align: { vertical: v === 'bottom' ? null : v } }, 'Align');
}

export type Orientation = 'ccw' | 'cw' | 'vertical' | 'up' | 'down' | 'none';

/** Text orientation (XLSX textRotation: 1..90 up, 91..180 down, 255 stacked). */
export function setOrientation(ctx: Ctx, o: Orientation): void {
  const cur = ctx.state.format.rotation;
  const target = { ccw: 45, cw: 135, vertical: 255, up: 90, down: 180, none: 0 }[o];
  format(ctx, { align: { rotation: cur === target || !target ? null : target } }, 'Orientation');
}

export function indent(ctx: Ctx, delta: 1 | -1): void {
  const { doc } = ctx;
  const style = doc.styleAt(doc.selection.active.r, doc.selection.active.c);
  const cur = style.align?.indent ?? 0;
  const next = Math.max(0, Math.min(250, cur + delta));
  const h = style.align?.horizontal;
  format(ctx, { align: { indent: next || null, horizontal: next && (!h || h === 'general' || h === 'center') ? 'left' : h ?? null } }, 'Indent');
}

export function toggleWrap(ctx: Ctx): void {
  format(ctx, { align: { wrap: ctx.state.format.wrap === 'wrap' ? null : true } }, 'Wrap text');
}

export async function mergeCells(ctx: Ctx, mode: 'center' | 'across' | 'merge' | 'unmerge'): Promise<void> {
  commit(ctx);
  const { doc } = ctx;
  const sheet = doc.sheet;
  const range = doc.range;
  if (mode === 'unmerge') {
    doc.unmerge();
    return;
  }
  if (mode === 'center' && ctx.state.format.merged) {
    doc.unmerge();
    return;
  }
  if (range.r0 === range.r1 && range.c0 === range.c1) return;
  if (sheet.tables?.some((t) => rangesIntersect(t.ref, range))) {
    ctx.notify('Cells in a table cannot be merged.');
    return;
  }
  let filled = 0;
  sheet.forEachCellIn(range, (_r, _c, cell) => {
    if (cell.v !== null && cell.v !== '') filled++;
  });
  if (filled > 1 && !(await ctx.confirm('Merging cells only keeps the upper-left value and discards other values.', { title: 'Merge Cells' }))) return;
  doc.transact('Merge', () => {
    doc.merge(undefined, undefined, mode === 'across' ? 'across' : 'all');
    if (mode === 'center') doc.format({ align: { horizontal: 'center' } });
  });
}

// -----------------------------------------------------------------------------------------------
// Number formats

export const NUMBER_FORMATS: { label: string; code: string }[] = [
  { label: 'General', code: 'General' },
  { label: 'Number', code: '0.00' },
  { label: 'Currency', code: '"$"#,##0.00' },
  { label: 'Accounting', code: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' },
  { label: 'Short Date', code: 'm/d/yyyy' },
  { label: 'Long Date', code: 'dddd, mmmm d, yyyy' },
  { label: 'Time', code: 'h:mm:ss AM/PM' },
  { label: 'Percentage', code: '0.00%' },
  { label: 'Fraction', code: '# ?/?' },
  { label: 'Scientific', code: '0.00E+00' },
  { label: 'Text', code: '@' },
];

export const COMMA_STYLE = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';

/** Category shown in the number format box. */
export function numberFormatLabel(code: string): string {
  const exact = NUMBER_FORMATS.find((f) => f.code === code);
  if (exact) return exact.label;
  if (code === 'General') return 'General';
  if (code === '@') return 'Text';
  const info = formatInfo(code);
  if (!info) return 'Custom';
  if (/\*/.test(code) && /_\(/.test(code)) return 'Accounting';
  switch (info.type) {
    case 'currency':
      return 'Currency';
    case 'date':
      return /dddd|mmmm/.test(code) ? 'Long Date' : 'Short Date';
    case 'time':
      return 'Time';
    case 'datetime':
      return 'Custom';
    case 'percent':
      return 'Percentage';
    case 'fraction':
      return 'Fraction';
    case 'scientific':
      return 'Scientific';
    case 'number':
    case 'grouped':
      return 'Number';
    case 'text':
      return 'Text';
    default:
      return 'Custom';
  }
}

export function setNumberFormat(ctx: Ctx, code: string): void {
  commit(ctx);
  ctx.doc.setNumberFormat(code);
}

// -----------------------------------------------------------------------------------------------
// Cells, rows, columns, sheets

export function insertRows(ctx: Ctx): void {
  commit(ctx);
  ctx.doc.insertRowsAtSelection('above');
}

export function insertCols(ctx: Ctx): void {
  commit(ctx);
  ctx.doc.insertColsAtSelection('left');
}

export function deleteRows(ctx: Ctx): void {
  commit(ctx);
  if (isFullColumns(ctx.doc.range)) {
    ctx.notify('Select rows to delete, not whole columns.');
    return;
  }
  ctx.doc.deleteSelectedRows();
}

export function deleteCols(ctx: Ctx): void {
  commit(ctx);
  if (isFullRows(ctx.doc.range)) {
    ctx.notify('Select columns to delete, not whole rows.');
    return;
  }
  ctx.doc.deleteSelectedCols();
}

function selectedLines(ctx: Ctx, axis: 'row' | 'col'): number[] {
  const out = new Set<number>();
  for (const r of ctx.doc.selection.ranges) {
    const b = axis === 'row' ? [r.r0, Math.min(r.r1, r.r0 + 100000)] : [r.c0, r.c1];
    for (let i = b[0]; i <= b[1]; i++) out.add(i);
  }
  return [...out];
}

export function autoFit(ctx: Ctx, axis: 'row' | 'col'): void {
  commit(ctx);
  const lines = selectedLines(ctx, axis);
  if (axis === 'row') ctx.doc.autoFitRows(lines.slice(0, 5000));
  else ctx.doc.autoFitColumns(lines);
}

export function setHidden(ctx: Ctx, axis: 'row' | 'col', hidden: boolean): void {
  commit(ctx);
  const sheet = ctx.doc.sheet;
  let lines = selectedLines(ctx, axis);
  if (!hidden) {
    // Unhiding a selection also reveals hidden lines between its edges.
    lines = lines.filter((i) => (axis === 'row' ? sheet.isRowHidden(i) : sheet.isColHidden(i)));
    if (!lines.length) {
      const range = ctx.doc.range;
      const from = axis === 'row' ? range.r0 : range.c0;
      if (from > 0 && (axis === 'row' ? sheet.isRowHidden(from - 1) : sheet.isColHidden(from - 1))) lines = [from - 1];
    }
  }
  if (lines.length > 20000) lines = lines.slice(0, 20000);
  ctx.doc.setHidden(axis, lines, hidden);
}

export function newWorksheet(ctx: Ctx): void {
  commit(ctx);
  ctx.doc.addSheet(undefined, ctx.doc.activeIndex + 1);
}

export async function renameWorksheet(ctx: Ctx, index = ctx.doc.activeIndex): Promise<void> {
  const sheet = ctx.doc.wb.sheets[index];
  if (!sheet) return;
  let value = sheet.name;
  for (;;) {
    const next = await ctx.prompt('Sheet name', value, { title: 'Rename Sheet', ok: 'OK' });
    if (next === null || next === sheet.name) return;
    const error = ctx.doc.renameSheet(index, next);
    if (!error) return;
    ctx.notify(error);
    value = next;
  }
}

export async function deleteWorksheet(ctx: Ctx, index = ctx.doc.activeIndex): Promise<void> {
  if (ctx.doc.wb.visibleSheets().length <= 1) {
    ctx.notify('A workbook must contain at least one visible worksheet.');
    return;
  }
  const sheet = ctx.doc.wb.sheets[index];
  const empty = !sheet.cells.size && !sheet.charts.length && !sheet.images.length && !sheet.shapes.length;
  if (!empty && !(await ctx.confirm(`The sheet “${sheet.name}” will be permanently deleted. Do you want to continue?`, { title: 'Delete Sheet', ok: 'Delete', danger: true }))) return;
  ctx.doc.removeSheet(index);
}

// -----------------------------------------------------------------------------------------------
// Editing

export function autoSum(ctx: Ctx, fn = 'SUM'): void {
  commit(ctx);
  ctx.grid?.autoSum(fn);
}

export function fillDirection(ctx: Ctx, dir: 'down' | 'right' | 'up' | 'left'): void {
  commit(ctx);
  const { doc } = ctx;
  const range = doc.range;
  if (dir === 'down' || dir === 'up') {
    if (range.r0 === range.r1) {
      if (dir === 'down' && range.r0 > 0) {
        doc.select(makeRange(range.r0 - 1, range.c0, range.r0 - 1, range.c1));
        doc.fill(range, true);
        doc.select(range);
      }
      return;
    }
  } else if (range.c0 === range.c1) {
    if (dir === 'right' && range.c0 > 0) {
      doc.select(makeRange(range.r0, range.c0 - 1, range.r1, range.c0 - 1));
      doc.fill(range, true);
      doc.select(range);
    }
    return;
  }
  const source = dir === 'down' ? makeRange(range.r0, range.c0, range.r0, range.c1)
    : dir === 'up' ? makeRange(range.r1, range.c0, range.r1, range.c1)
      : dir === 'right' ? makeRange(range.r0, range.c0, range.r1, range.c0)
        : makeRange(range.r0, range.c1, range.r1, range.c1);
  doc.select(source);
  doc.fill(range, true);
  doc.select(range);
}

export function clear(ctx: Ctx, what: 'contents' | 'formats' | 'all' | 'comments' | 'links'): void {
  commit(ctx);
  ctx.doc.clear(what);
}

/** Whether the first row of a block looks like a header (text above numbers or dates). */
export function looksLikeHeader(ctx: Ctx, region: Range, col: number): boolean {
  const { doc } = ctx;
  if (doc.sheet.autoFilter && rangesIntersect(doc.sheet.autoFilter.ref, region)) return true;
  if (doc.sheet.tables?.some((t) => t.headerRow && t.ref.r0 === region.r0 && rangesIntersect(t.ref, region))) return true;
  if (region.r1 <= region.r0) return false;
  const first = doc.cell(region.r0, col)?.v;
  const second = doc.cell(region.r0 + 1, col)?.v;
  if (typeof first === 'string' && (typeof second === 'number' || typeof second === 'boolean')) return true;
  let allText = true;
  for (let c = region.c0; c <= region.c1; c++) {
    const v = doc.cell(region.r0, c)?.v;
    if (v !== null && v !== undefined && typeof v !== 'string') allText = false;
  }
  if (!allText) return false;
  for (let c = region.c0; c <= region.c1; c++) if (typeof doc.cell(region.r0 + 1, c)?.v === 'number') return true;
  return false;
}

/** The block a sort or filter applies to: the selection, or the data region around the active cell. */
export function dataRegion(ctx: Ctx): Range | null {
  const { doc } = ctx;
  const sel = doc.selection;
  const range = doc.range;
  const table = doc.sheet.tables?.find((t) => rangesIntersect(t.ref, makeRange(sel.active.r, sel.active.c)));
  if (table && range.r0 === range.r1 && range.c0 === range.c1) return table.ref;
  if (!(range.r0 === range.r1 && range.c0 === range.c1) && !isFullColumns(range) && !isFullRows(range)) return range;
  return filterRegion(doc.sheet, makeRange(sel.active.r, sel.active.c), sel.active);
}

export function sortSelection(ctx: Ctx, ascending: boolean): void {
  commit(ctx);
  const area = dataRegion(ctx);
  if (!area) {
    ctx.notify('Select a cell in the data to sort.');
    return;
  }
  const col = Math.max(area.c0, Math.min(area.c1, ctx.doc.selection.active.c));
  ctx.doc.sortRange(area, col, ascending, looksLikeHeader(ctx, area, col));
}

export function toggleFilter(ctx: Ctx): void {
  commit(ctx);
  const { doc } = ctx;
  const sel = doc.selection.active;
  const table = doc.sheet.tables?.find((t) => rangesIntersect(t.ref, makeRange(sel.r, sel.c)));
  if (table) {
    doc.setTableOptions(table.id, { autoFilter: !table.autoFilter });
    return;
  }
  const on = doc.toggleAutoFilter();
  if (!on && !doc.sheet.autoFilter && !ctx.state.hasFilter) ctx.notify('Select a cell inside your data first.');
}

export function clearFilter(ctx: Ctx): void {
  commit(ctx);
  ctx.doc.clearFilters();
  for (const t of ctx.doc.sheet.tables ?? []) if (t.filters.size) ctx.doc.clearTableFilters(t.id);
}

export function reapplyFilter(ctx: Ctx): void {
  commit(ctx);
  ctx.doc.reapplyFilters();
}

// -----------------------------------------------------------------------------------------------
// View

export type FreezeKind = 'panes' | 'top-row' | 'first-col' | 'none';

export function freeze(ctx: Ctx, kind: FreezeKind): void {
  commit(ctx);
  const { doc } = ctx;
  const sheet = doc.sheet;
  if (kind === 'none') doc.setFreeze(0, 0);
  else if (kind === 'top-row') doc.setFreeze(1, 0);
  else if (kind === 'first-col') doc.setFreeze(0, 1);
  else {
    if (sheet.freeze.rows || sheet.freeze.cols) {
      doc.setFreeze(0, 0);
      return;
    }
    const { r, c } = doc.selection.active;
    if (r === 0 && c === 0) {
      ctx.notify('Select the cell below and to the right of the rows and columns to freeze.');
      return;
    }
    doc.setFreeze(r, c);
  }
}

export function zoomTo(ctx: Ctx, zoom: number): void {
  const z = Math.max(0.1, Math.min(4, zoom));
  if (ctx.grid) ctx.grid.setZoom(z);
  else ctx.doc.setZoom(z);
}

export function zoomToSelection(ctx: Ctx): void {
  const g = ctx.grid;
  if (!g) return;
  const range = ctx.doc.bounded(ctx.doc.range, ctx.doc.sheet) ?? ctx.doc.range;
  const vt = g.transform();
  const rect = vt.rangeRect(range);
  const z0 = vt.zoom;
  const w = rect.w / z0;
  const h = rect.h / z0;
  const avail = g.root.getBoundingClientRect();
  const z = Math.min((avail.width - 60) / Math.max(1, w), (avail.height - 40) / Math.max(1, h));
  zoomTo(ctx, Math.max(0.1, Math.min(4, z)));
  g.scrollIntoView(range.r0, range.c0);
}

// -----------------------------------------------------------------------------------------------
// Comments and notes

export function newComment(ctx: Ctx): void {
  const { r, c } = ctx.doc.selection.active;
  const m = ctx.doc.sheet.mergeAt(r, c);
  ctx.openComments(m?.r0 ?? r, m?.c0 ?? c, true);
}

export function editNote(ctx: Ctx): void {
  const { r, c } = ctx.doc.selection.active;
  const m = ctx.doc.sheet.mergeAt(r, c);
  ctx.openNote(m?.r0 ?? r, m?.c0 ?? c);
}

export function deleteNote(ctx: Ctx): void {
  const { r, c } = ctx.doc.selection.active;
  ctx.doc.setNote(r, c, '');
}

/** Moves the selection to the next or previous comment (Review > Next / Previous). */
export function stepComment(ctx: Ctx, dir: 1 | -1): void {
  const { doc } = ctx;
  const sheets = doc.wb.sheets.map((s, i) => ({ s, i })).filter(({ s }) => s.state === 'visible');
  const all: { sheet: number; r: number; c: number }[] = [];
  for (const { s, i } of sheets) for (const t of [...s.comments].filter((t) => t.kind === 'threaded').sort((a, b) => a.r - b.r || a.c - b.c)) all.push({ sheet: i, r: t.r, c: t.c });
  if (!all.length) return;
  const cur = { sheet: doc.activeIndex, ...doc.selection.active };
  const key = (x: { sheet: number; r: number; c: number }) => x.sheet * 1e12 + x.r * 1e5 + x.c;
  const k = key(cur);
  const next = dir > 0 ? all.find((x) => key(x) > k) ?? all[0] : [...all].reverse().find((x) => key(x) < k) ?? all[all.length - 1];
  doc.setActiveSheet(next.sheet);
  doc.selectCell(next.r, next.c);
  ctx.grid?.scrollIntoView(next.r, next.c);
  ctx.openComments(next.r, next.c);
}

// -----------------------------------------------------------------------------------------------
// Pictures

export function uploadImage(ctx: Ctx): void {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.jpg,.jpeg,.png,.gif,.webp,.bmp,.svg,image/*';
  input.onchange = () => {
    const file = input.files?.[0];
    if (file) void ctx.grid?.insertImageFile(file, file.name);
  };
  input.click();
}

export function selectAll(ctx: Ctx): void {
  ctx.doc.selectAll();
}

export { MAX_COLS, MAX_ROWS };
