/** Context menus of cells, row and column headers, the corner, sheet tabs and floating objects. */
import type { MenuItem } from './menuItems';
import type { UIContextValue } from './context';
import * as cmd from './commands';
import { STANDARD_COLORS } from './palette';
import { resolveColor, rgbSpec } from '../core/model/color';
import type { IconName } from './icons';
import type { PasteOptions } from '../core/document';

type Ctx = UIContextValue;

const PASTE_OPTIONS: { title: string; icon: IconName; opts: PasteOptions }[] = [
  { title: 'Paste', icon: 'content_paste', opts: { what: 'all' } },
  { title: 'Values', icon: '123', opts: { what: 'values' } },
  { title: 'Formulas', icon: 'functions', opts: { what: 'formulas' } },
  { title: 'Transpose', icon: 'swap_horiz', opts: { what: 'all', transpose: true } },
  { title: 'Formatting', icon: 'format_paint', opts: { what: 'formats' } },
  { title: 'Paste Link', icon: 'link', opts: { what: 'all', link: true } },
];

function pasteItems(ctx: Ctx): MenuItem[] {
  return [
    { title: 'Cut', icon: 'content_cut', shortcut: 'Ctrl+X', testId: 'ctx-cut', onClick: () => cmd.cut(ctx) },
    { title: 'Copy', icon: 'content_copy', shortcut: 'Ctrl+C', testId: 'ctx-copy', onClick: () => cmd.copy(ctx) },
    {
      type: 'gallery',
      title: 'Paste Options:',
      columns: 6,
      gallery: PASTE_OPTIONS.map((p, i) => ({
        title: p.title,
        testId: `ctx-paste-${i}`,
        preview: { kind: 'paste', badge: i ? p.icon : undefined },
        onClick: () => void cmd.paste(ctx, p.opts),
      })),
    },
    { title: 'Paste Special...', disabled: !ctx.doc.hasClip(), testId: 'ctx-paste-special', onClick: () => ctx.openDialog({ kind: 'paste-special' }) },
  ];
}

export function cellMenu(ctx: Ctx): MenuItem[] {
  const { doc } = ctx;
  const { r, c } = doc.selection.active;
  const m = doc.sheet.mergeAt(r, c);
  const thread = doc.sheet.commentAt(m?.r0 ?? r, m?.c0 ?? c);
  const table = doc.tableAt(r, c);
  const link = doc.sheet.hyperlinkAt(r, c);
  const items: MenuItem[] = [
    ...pasteItems(ctx),
    { type: 'line' },
    table
      ? {
          title: 'Insert',
          testId: 'ctx-insert',
          submenu: [
            { title: 'Table Rows Above', onClick: () => cmd.insertRows(ctx) },
            { title: 'Table Columns to the Left', onClick: () => cmd.insertCols(ctx) },
          ],
        }
      : { title: 'Insert...', testId: 'ctx-insert', onClick: () => ctx.openDialog({ kind: 'insert-cells' }) },
    table
      ? {
          title: 'Delete',
          testId: 'ctx-delete',
          submenu: [
            { title: 'Table Rows', onClick: () => cmd.deleteRows(ctx) },
            { title: 'Table Columns', onClick: () => cmd.deleteCols(ctx) },
          ],
        }
      : { title: 'Delete...', testId: 'ctx-delete', onClick: () => ctx.openDialog({ kind: 'delete-cells' }) },
    { title: 'Clear Contents', testId: 'ctx-clear', onClick: () => cmd.clear(ctx, 'contents') },
    { type: 'line' },
    {
      title: 'Filter',
      icon: 'filter_alt',
      submenu: [
        { title: 'Filter by Selected Cell\'s Value', onClick: () => filterBySelection(ctx) },
        { title: 'Clear Filter', onClick: () => cmd.clearFilter(ctx) },
        { title: 'Reapply', onClick: () => cmd.reapplyFilter(ctx) },
      ],
    },
    {
      title: 'Sort',
      icon: 'sort',
      testId: 'ctx-sort',
      submenu: [
        { title: 'Sort A to Z', icon: 'sort_by_alpha', onClick: () => cmd.sortSelection(ctx, true) },
        { title: 'Sort Z to A', onClick: () => cmd.sortSelection(ctx, false) },
        { title: 'Custom Sort...', onClick: () => ctx.openDialog({ kind: 'sort' }) },
      ],
    },
  ];
  if (table) {
    items.push({
      title: 'Table',
      icon: 'table_view',
      testId: 'ctx-table',
      submenu: [
        { title: 'Totals Row', checked: table.totalsRow, onClick: () => doc.setTableOptions(table.id, { totalsRow: !table.totalsRow }) },
        { title: 'Convert to Range', onClick: () => doc.convertTableToRange(table.id) },
        { title: 'Resize Table...', onClick: () => ctx.openDialog({ kind: 'resize-table', tableId: table.id }) },
      ],
    });
  }
  items.push({ type: 'line' });
  if (thread?.kind === 'threaded') {
    items.push({ title: 'Reply to Comment', icon: 'comment', testId: 'ctx-comment', onClick: () => ctx.openComments(thread.r, thread.c) });
    items.push({ title: 'Delete Comment', icon: 'delete', onClick: () => doc.deleteComment(thread.comments[0].id) });
  } else items.push({ title: 'New Threaded Comment', icon: 'add_comment', testId: 'ctx-comment', disabled: !!thread, onClick: () => cmd.newComment(ctx) });
  if (thread?.kind === 'note') {
    items.push({ title: 'Edit Note', icon: 'sticky_note_2', testId: 'ctx-note', onClick: () => cmd.editNote(ctx) });
    items.push({ title: 'Delete Note', onClick: () => cmd.deleteNote(ctx) });
  } else if (!thread) items.push({ title: 'New Note', icon: 'sticky_note_2', testId: 'ctx-note', onClick: () => cmd.editNote(ctx) });
  items.push(
    { type: 'line' },
    { title: 'Format Cells...', icon: 'format_cells', shortcut: 'Ctrl+1', testId: 'ctx-format-cells', onClick: () => ctx.openDialog({ kind: 'format-cells' }) },
    { title: 'Define Name...', testId: 'ctx-define-name', onClick: () => ctx.openDialog({ kind: 'new-name' }) },
    link
      ? { title: 'Edit Hyperlink...', icon: 'link', testId: 'ctx-link', onClick: () => ctx.openDialog({ kind: 'hyperlink' }) }
      : { title: 'Link...', icon: 'link', shortcut: 'Ctrl+K', testId: 'ctx-link', onClick: () => ctx.openDialog({ kind: 'hyperlink' }) },
  );
  if (link) items.push({ title: 'Remove Hyperlink', icon: 'link_off', onClick: () => doc.removeLink() });
  return items;
}

/** Filters the column of the active cell by the active cell's value (auto filter or table). */
function filterBySelection(ctx: Ctx): void {
  const { doc } = ctx;
  const { r, c } = doc.selection.active;
  const text = doc.displayText(r, c);
  const table = doc.tableAt(r, c);
  if (table && table.headerRow) {
    if (!table.autoFilter) doc.setTableOptions(table.id, { autoFilter: true });
    doc.setTableFilter(table.id, c, { values: [text], blanks: text === '' });
    return;
  }
  if (!doc.sheet.autoFilter) doc.toggleAutoFilter();
  if (doc.sheet.autoFilter) doc.setColumnFilter(c, { values: [text], blanks: text === '' });
}

export function rowMenu(ctx: Ctx): MenuItem[] {
  return [
    ...pasteItems(ctx),
    { type: 'line' },
    { title: 'Insert', testId: 'ctx-insert-rows', onClick: () => cmd.insertRows(ctx) },
    { title: 'Delete', testId: 'ctx-delete-rows', onClick: () => cmd.deleteRows(ctx) },
    { title: 'Clear Contents', onClick: () => cmd.clear(ctx, 'contents') },
    { type: 'line' },
    { title: 'Format Cells...', icon: 'format_cells', onClick: () => ctx.openDialog({ kind: 'format-cells' }) },
    { title: 'Row Height...', icon: 'height', testId: 'ctx-row-height', onClick: () => ctx.openDialog({ kind: 'size', what: 'row-height' }) },
    { title: 'Hide', testId: 'ctx-hide-rows', onClick: () => cmd.setHidden(ctx, 'row', true) },
    { title: 'Unhide', testId: 'ctx-unhide-rows', onClick: () => cmd.setHidden(ctx, 'row', false) },
  ];
}

export function columnMenu(ctx: Ctx): MenuItem[] {
  return [
    ...pasteItems(ctx),
    { type: 'line' },
    { title: 'Insert', testId: 'ctx-insert-cols', onClick: () => cmd.insertCols(ctx) },
    { title: 'Delete', testId: 'ctx-delete-cols', onClick: () => cmd.deleteCols(ctx) },
    { title: 'Clear Contents', onClick: () => cmd.clear(ctx, 'contents') },
    { type: 'line' },
    { title: 'Format Cells...', icon: 'format_cells', onClick: () => ctx.openDialog({ kind: 'format-cells' }) },
    { title: 'Column Width...', icon: 'width_normal', testId: 'ctx-col-width', onClick: () => ctx.openDialog({ kind: 'size', what: 'col-width' }) },
    { title: 'Hide', testId: 'ctx-hide-cols', onClick: () => cmd.setHidden(ctx, 'col', true) },
    { title: 'Unhide', testId: 'ctx-unhide-cols', onClick: () => cmd.setHidden(ctx, 'col', false) },
  ];
}

export function cornerMenu(ctx: Ctx): MenuItem[] {
  return [
    ...pasteItems(ctx),
    { type: 'line' },
    { title: 'Clear Contents', onClick: () => cmd.clear(ctx, 'contents') },
    { title: 'Format Cells...', icon: 'format_cells', onClick: () => ctx.openDialog({ kind: 'format-cells' }) },
    { title: 'Row Height...', onClick: () => ctx.openDialog({ kind: 'size', what: 'row-height' }) },
    { title: 'Column Width...', onClick: () => ctx.openDialog({ kind: 'size', what: 'col-width' }) },
    { title: 'Define Name...', onClick: () => ctx.openDialog({ kind: 'new-name' }) },
  ];
}

export function tabMenu(ctx: Ctx, index: number): MenuItem[] {
  const { doc } = ctx;
  const sheet = doc.wb.sheets[index];
  const locked = !!doc.wb.protection?.structure;
  const theme = doc.wb.theme;
  const themeColors = [4, 5, 6, 7, 8, 9, 0, 1, 2, 3].map((t) => ({ spec: { theme: t }, css: `#${resolveColor({ theme: t }, theme, '000000')}` }));
  return [
    { title: 'Insert', icon: 'post_add', disabled: locked, testId: 'tab-insert', onClick: () => doc.addSheet(undefined, index + 1) },
    { title: 'Delete', icon: 'delete', disabled: locked, testId: 'tab-delete', onClick: () => void cmd.deleteWorksheet(ctx, index) },
    { title: 'Rename', icon: 'drive_file_rename_outline', disabled: locked, testId: 'tab-rename', onClick: () => void cmd.renameWorksheet(ctx, index) },
    { title: 'Move or Copy...', disabled: locked, testId: 'tab-move-copy', onClick: () => ctx.openDialog({ kind: 'move-copy-sheet', index }) },
    { type: 'line' },
    sheet.protection
      ? { title: 'Unprotect Sheet...', icon: 'lock_open', onClick: () => { doc.setActiveSheet(index); ctx.openDialog({ kind: 'unprotect-sheet' }); } }
      : { title: 'Protect Sheet...', icon: 'lock', testId: 'tab-protect', onClick: () => { doc.setActiveSheet(index); ctx.openDialog({ kind: 'protect-sheet' }); } },
    {
      title: 'Tab Color',
      icon: 'palette',
      testId: 'tab-color',
      submenu: [
        { title: 'No Color', onClick: () => doc.setTabColor(index, null) },
        {
          type: 'gallery',
          title: 'Theme Colors',
          columns: 10,
          gallery: themeColors.map((x) => ({ title: x.css, preview: { kind: 'color', css: x.css }, onClick: () => doc.setTabColorSpec(index, x.spec) })),
        },
        {
          type: 'gallery',
          title: 'Standard Colors',
          columns: 10,
          gallery: STANDARD_COLORS.map((x) => ({ title: x.name, preview: { kind: 'color', css: x.hex }, onClick: () => doc.setTabColorSpec(index, rgbSpec(x.hex)!) })),
        },
      ],
    },
    { type: 'line' },
    { title: 'Hide', disabled: locked || doc.wb.visibleSheets().length <= 1, testId: 'tab-hide', onClick: () => doc.setSheetHidden(index, true) },
    { title: 'Unhide...', disabled: locked || !doc.wb.sheets.some((s) => s.state === 'hidden'), testId: 'tab-unhide', onClick: () => ctx.openDialog({ kind: 'unhide-sheet' }) },
  ];
}

export function objectMenu(ctx: Ctx, id: string): MenuItem[] {
  const { doc } = ctx;
  const hit = doc.objectById(id);
  const items: MenuItem[] = [
    { title: 'Delete', icon: 'delete', testId: 'obj-delete', onClick: () => { doc.deleteObject(id); ctx.selectObject(null); } },
    { type: 'line' },
    { title: 'Bring to Front', icon: 'flip_to_front', testId: 'obj-front', onClick: () => doc.reorderObject(id, 'front') },
    { title: 'Bring Forward', onClick: () => doc.reorderObject(id, 'forward') },
    { title: 'Send to Back', icon: 'flip_to_back', testId: 'obj-back', onClick: () => doc.reorderObject(id, 'back') },
    { title: 'Send Backward', onClick: () => doc.reorderObject(id, 'backward') },
    { type: 'line' },
  ];
  if (hit?.kind === 'chart') {
    items.push(
      { title: 'Change Chart Type...', icon: 'insert_chart', testId: 'obj-chart-type', onClick: () => ctx.openDialog({ kind: 'insert-chart', changeId: id }) },
      { title: 'Select Data...', icon: 'dataset', testId: 'obj-select-data', onClick: () => ctx.openDialog({ kind: 'select-data', chartId: id }) },
      { title: 'Move Chart...', onClick: () => ctx.openDialog({ kind: 'move-chart', chartId: id }) },
      { title: 'Format Chart Area...', icon: 'tune', testId: 'obj-format', onClick: () => ctx.setPanel('format') },
    );
  } else if (hit?.kind === 'shape') {
    items.push({ title: 'Format Shape...', icon: 'tune', testId: 'obj-format', onClick: () => ctx.setPanel('format') });
  } else items.push({ title: 'Format Picture...', icon: 'tune', testId: 'obj-format', onClick: () => ctx.setPanel('format') });
  items.push({ title: 'Edit Alt Text...', icon: 'accessibility', onClick: () => ctx.openDialog({ kind: 'alt-text', id }) });
  return items;
}
