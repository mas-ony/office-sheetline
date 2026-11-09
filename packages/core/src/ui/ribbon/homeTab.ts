/** Home tab: Undo, Clipboard, Font, Alignment, Number, Styles, Cells, Editing. */
import type { MenuItem } from '../menuItems';
import * as cmd from '../commands';
import { FONT_NAMES, FONT_SIZES, NUMBER_FORMATS, numberFormatLabel } from '../commands';
import { TINT, type Command, type Control, type Ctx, type Tab } from './types';
import { borderMenu, condFormatMenu } from './galleries';
import { isFullColumns, isFullRows } from '../../core/refs';

const fmt = (ctx: Ctx) => ctx.state.format;

function historyMenu(ctx: Ctx, which: 'undo' | 'redo'): MenuItem[] {
  const labels = ctx.doc.history.labels()[which].slice(0, 20);
  return labels.map((label, i) => ({
    title: label,
    onClick: () => {
      for (let k = 0; k <= i; k++) (which === 'undo' ? cmd.undo : cmd.redo)(ctx);
    },
  }));
}

export const C = {
  undo: { id: 'undo', label: 'Undo', icon: 'undo', shortcut: 'Ctrl+Z', run: cmd.undo, menu: (ctx) => historyMenu(ctx, 'undo'), disabled: (ctx) => !ctx.state.canUndo && !ctx.edit.active } as Command,
  redo: { id: 'redo', label: 'Redo', icon: 'redo', shortcut: 'Ctrl+Y', run: cmd.redo, menu: (ctx) => historyMenu(ctx, 'redo'), disabled: (ctx) => !ctx.state.canRedo } as Command,
  paste: { id: 'paste', label: 'Paste', icon: 'content_paste', tint: TINT.orange, shortcut: 'Ctrl+V', run: (ctx) => void cmd.paste(ctx), panel: { kind: 'paste' } } as Command,
  cut: { id: 'cut', label: 'Cut', icon: 'content_cut', shortcut: 'Ctrl+X', run: cmd.cut } as Command,
  copy: { id: 'copy', label: 'Copy', icon: 'content_copy', shortcut: 'Ctrl+C', run: cmd.copy } as Command,
  painter: {
    id: 'format-painter',
    label: 'Format Painter',
    tip: 'Format Painter (double-click to keep it on)',
    icon: 'format_paint',
    tint: TINT.orange,
    checked: (ctx) => !!ctx.painter,
    run: (ctx) => {
      if (ctx.painter) {
        ctx.setPainter(null);
        return;
      }
      ctx.doc.copy();
      ctx.grid?.markClip();
      ctx.setPainter('once');
    },
  } as Command,
  growFont: { id: 'font-grow', label: 'Increase Font Size', icon: 'text_increase', run: (ctx) => cmd.growFont(ctx, 1) } as Command,
  shrinkFont: { id: 'font-shrink', label: 'Decrease Font Size', icon: 'text_decrease', run: (ctx) => cmd.growFont(ctx, -1) } as Command,
  bold: { id: 'bold', label: 'Bold', icon: 'format_bold', shortcut: 'Ctrl+B', checked: (ctx) => fmt(ctx).bold, run: (ctx) => cmd.toggleFont(ctx, 'bold') } as Command,
  italic: { id: 'italic', label: 'Italic', icon: 'format_italic', shortcut: 'Ctrl+I', checked: (ctx) => fmt(ctx).italic, run: (ctx) => cmd.toggleFont(ctx, 'italic') } as Command,
  underline: { id: 'underline', label: 'Underline', icon: 'format_underlined', shortcut: 'Ctrl+U', checked: (ctx) => fmt(ctx).underline && fmt(ctx).underlineKind !== 'double', run: (ctx) => cmd.toggleUnderline(ctx, 'single') } as Command,
  doubleUnderline: { id: 'double-underline', label: 'Double Underline', icon: 'format_underlined', checked: (ctx) => fmt(ctx).underlineKind === 'double', run: (ctx) => cmd.toggleUnderline(ctx, 'double') } as Command,
  strike: { id: 'strikethrough', label: 'Strikethrough', icon: 'format_strikethrough', shortcut: 'Ctrl+5', checked: (ctx) => fmt(ctx).strike, run: (ctx) => cmd.toggleFont(ctx, 'strike') } as Command,
  borders: { id: 'borders', label: 'Borders', icon: 'border_bottom', run: (ctx) => cmd.applyBorder(ctx, 'bottom'), menu: borderMenu } as Command,
  alignTop: { id: 'align-top', label: 'Top Align', icon: 'vertical_align_top', checked: (ctx) => fmt(ctx).vAlign === 'top', run: (ctx) => cmd.setVAlign(ctx, 'top') } as Command,
  alignMiddle: { id: 'align-middle', label: 'Middle Align', icon: 'vertical_align_center', checked: (ctx) => fmt(ctx).vAlign === 'center', run: (ctx) => cmd.setVAlign(ctx, 'center') } as Command,
  alignBottom: { id: 'align-bottom', label: 'Bottom Align', icon: 'vertical_align_bottom', checked: (ctx) => fmt(ctx).vAlign === 'bottom', run: (ctx) => cmd.setVAlign(ctx, 'bottom') } as Command,
  orientation: {
    id: 'orientation',
    label: 'Orientation',
    icon: 'text_rotation_angleup',
    menu: (ctx) => [
      { title: 'Angle Counterclockwise', icon: 'text_rotation_angleup', checked: fmt(ctx).rotation === 45, onClick: () => cmd.setOrientation(ctx, 'ccw') },
      { title: 'Angle Clockwise', icon: 'text_rotation_angledown', checked: fmt(ctx).rotation === 135, onClick: () => cmd.setOrientation(ctx, 'cw') },
      { title: 'Vertical Text', icon: 'text_rotate_vertical', checked: fmt(ctx).rotation === 255, onClick: () => cmd.setOrientation(ctx, 'vertical') },
      { title: 'Rotate Text Up', icon: 'text_rotate_up', checked: fmt(ctx).rotation === 90, onClick: () => cmd.setOrientation(ctx, 'up') },
      { title: 'Rotate Text Down', icon: 'text_rotation_down', checked: fmt(ctx).rotation === 180, onClick: () => cmd.setOrientation(ctx, 'down') },
      { type: 'line' },
      { title: 'Format Cell Alignment', icon: 'format_align_left', onClick: () => ctx.openDialog({ kind: 'format-cells', tab: 'alignment' }) },
    ],
  } as Command,
  alignLeft: { id: 'align-left', label: 'Align Left', icon: 'format_align_left', checked: (ctx) => fmt(ctx).hAlign === 'left', run: (ctx) => cmd.setHAlign(ctx, 'left') } as Command,
  alignCenter: { id: 'align-center', label: 'Center', icon: 'format_align_center', checked: (ctx) => fmt(ctx).hAlign === 'center', run: (ctx) => cmd.setHAlign(ctx, 'center') } as Command,
  alignRight: { id: 'align-right', label: 'Align Right', icon: 'format_align_right', checked: (ctx) => fmt(ctx).hAlign === 'right', run: (ctx) => cmd.setHAlign(ctx, 'right') } as Command,
  indentDec: { id: 'indent-decrease', label: 'Decrease Indent', icon: 'format_indent_decrease', run: (ctx) => cmd.indent(ctx, -1) } as Command,
  indentInc: { id: 'indent-increase', label: 'Increase Indent', icon: 'format_indent_increase', run: (ctx) => cmd.indent(ctx, 1) } as Command,
  wrap: { id: 'wrap-text', label: 'Wrap Text', icon: 'wrap_text', checked: (ctx) => fmt(ctx).wrap === 'wrap', run: cmd.toggleWrap } as Command,
  merge: {
    id: 'merge',
    label: 'Merge & Center',
    icon: 'merge_center',
    checked: (ctx) => fmt(ctx).merged,
    run: (ctx) => void cmd.mergeCells(ctx, 'center'),
    menu: (ctx) => [
      { title: 'Merge & Center', icon: 'merge_center', testId: 'merge-center', onClick: () => void cmd.mergeCells(ctx, 'center') },
      { title: 'Merge Across', icon: 'call_merge', testId: 'merge-across', onClick: () => void cmd.mergeCells(ctx, 'across') },
      { title: 'Merge Cells', icon: 'merge', testId: 'merge-cells', onClick: () => void cmd.mergeCells(ctx, 'merge') },
      { title: 'Unmerge Cells', icon: 'call_split', testId: 'unmerge', onClick: () => void cmd.mergeCells(ctx, 'unmerge') },
    ],
  } as Command,
  percent: { id: 'percent-style', label: 'Percent Style', icon: 'percent', shortcut: 'Ctrl+Shift+%', run: (ctx) => cmd.setNumberFormat(ctx, '0%') } as Command,
  comma: { id: 'comma-style', label: 'Comma Style', icon: 'comma_style', run: (ctx) => cmd.setNumberFormat(ctx, cmd.COMMA_STYLE) } as Command,
  accounting: {
    id: 'accounting',
    label: 'Accounting Number Format',
    icon: 'attach_money',
    run: (ctx) => cmd.setNumberFormat(ctx, NUMBER_FORMATS[3].code),
    menu: (ctx) => [
      { title: '$ English (United States)', onClick: () => cmd.setNumberFormat(ctx, NUMBER_FORMATS[3].code) },
      { title: '£ English (United Kingdom)', onClick: () => cmd.setNumberFormat(ctx, '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-') },
      { title: '€ Euro (€ 123)', onClick: () => cmd.setNumberFormat(ctx, '_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "-"??_-;_-@_-') },
      { title: '¥ Chinese (PRC)', onClick: () => cmd.setNumberFormat(ctx, '_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* -#,##0.00_ ;_ [$¥-804]* "-"??_ ;_ @_ ') },
      { title: 'Rp Indonesian', onClick: () => cmd.setNumberFormat(ctx, '_-[$Rp-421]* #,##0.00_-;-[$Rp-421]* #,##0.00_-;_-[$Rp-421]* "-"??_-;_-@_-') },
      { type: 'line' },
      { title: 'More Accounting Formats...', onClick: () => ctx.openDialog({ kind: 'format-cells', tab: 'number' }) },
    ],
  } as Command,
  decInc: { id: 'decimal-increase', label: 'Increase Decimal', icon: 'decimal_increase', run: (ctx) => { cmd.commit(ctx); ctx.doc.adjustDecimals(1); } } as Command,
  decDec: { id: 'decimal-decrease', label: 'Decrease Decimal', icon: 'decimal_decrease', run: (ctx) => { cmd.commit(ctx); ctx.doc.adjustDecimals(-1); } } as Command,
  condFormat: { id: 'conditional-format', label: 'Conditional Format', icon: 'cond_format', keywords: 'conditional formatting rules', menu: condFormatMenu } as Command,
  formatTable: {
    id: 'format-table',
    label: 'Format Table',
    keywords: 'format as table styles',
    icon: 'format_table',
    panel: {
      kind: 'table-styles',
      pick: (ctx, name) => {
        const t = ctx.doc.tableAt(ctx.doc.selection.active.r, ctx.doc.selection.active.c);
        if (t) ctx.doc.setTableStyle(t.id, name);
        else ctx.openDialog({ kind: 'create-table', style: name });
      },
    },
  } as Command,
  cellStyles: { id: 'cell-styles', label: 'Cell Styles', icon: 'cell_styles', panel: { kind: 'cell-styles' } } as Command,
  insert: {
    id: 'insert-cells',
    label: 'Insert',
    icon: 'insert_cells',
    run: (ctx) => {
      const r = ctx.doc.range;
      if (isFullRows(r)) cmd.insertRows(ctx);
      else if (isFullColumns(r)) cmd.insertCols(ctx);
      else ctx.openDialog({ kind: 'insert-cells' });
    },
    menu: (ctx) => [
      { title: 'Insert Cells...', icon: 'insert_cells', testId: 'insert-cells-dialog', onClick: () => ctx.openDialog({ kind: 'insert-cells' }) },
      { title: 'Insert Sheet Rows', icon: 'table_rows', testId: 'insert-rows', onClick: () => cmd.insertRows(ctx) },
      { title: 'Insert Sheet Columns', icon: 'view_column', testId: 'insert-cols', onClick: () => cmd.insertCols(ctx) },
      { type: 'line' },
      { title: 'Insert Sheet', icon: 'sheet_tab', testId: 'insert-sheet', onClick: () => cmd.newWorksheet(ctx) },
    ],
  } as Command,
  delete: {
    id: 'delete-cells',
    label: 'Delete',
    icon: 'delete_cells',
    run: (ctx) => {
      const r = ctx.doc.range;
      if (isFullRows(r)) cmd.deleteRows(ctx);
      else if (isFullColumns(r)) cmd.deleteCols(ctx);
      else ctx.openDialog({ kind: 'delete-cells' });
    },
    menu: (ctx) => [
      { title: 'Delete Cells...', icon: 'delete_cells', testId: 'delete-cells-dialog', onClick: () => ctx.openDialog({ kind: 'delete-cells' }) },
      { title: 'Delete Sheet Rows', icon: 'table_rows', testId: 'delete-rows', onClick: () => cmd.deleteRows(ctx) },
      { title: 'Delete Sheet Columns', icon: 'view_column', testId: 'delete-cols', onClick: () => cmd.deleteCols(ctx) },
      { type: 'line' },
      { title: 'Delete Sheet', icon: 'delete', testId: 'delete-sheet', onClick: () => void cmd.deleteWorksheet(ctx) },
    ],
  } as Command,
  format: {
    id: 'format-cells-menu',
    label: 'Format',
    icon: 'format_cells',
    menu: (ctx) => [
      { type: 'header', title: 'Cell Size' },
      { title: 'Row Height...', icon: 'height', onClick: () => ctx.openDialog({ kind: 'size', what: 'row-height' }) },
      { title: 'AutoFit Row Height', onClick: () => cmd.autoFit(ctx, 'row') },
      { title: 'Default Height...', onClick: () => ctx.openDialog({ kind: 'size', what: 'default-height' }) },
      { title: 'Column Width...', icon: 'width_normal', onClick: () => ctx.openDialog({ kind: 'size', what: 'col-width' }) },
      { title: 'AutoFit Column Width', onClick: () => cmd.autoFit(ctx, 'col') },
      { title: 'Default Width...', onClick: () => ctx.openDialog({ kind: 'size', what: 'default-width' }) },
      { type: 'header', title: 'Visibility' },
      { title: 'Hide Rows', onClick: () => cmd.setHidden(ctx, 'row', true) },
      { title: 'Hide Columns', onClick: () => cmd.setHidden(ctx, 'col', true) },
      { title: 'Hide Sheet', onClick: () => { if (ctx.doc.wb.visibleSheets().length > 1) ctx.doc.setSheetHidden(ctx.doc.activeIndex, true); else ctx.notify('A workbook must contain at least one visible worksheet.'); } },
      { title: 'Unhide Rows', onClick: () => cmd.setHidden(ctx, 'row', false) },
      { title: 'Unhide Columns', onClick: () => cmd.setHidden(ctx, 'col', false) },
      { title: 'Unhide Sheet...', disabled: !ctx.doc.wb.sheets.some((s) => s.state === 'hidden'), onClick: () => ctx.openDialog({ kind: 'unhide-sheet' }) },
      { type: 'header', title: 'Organize Sheets' },
      { title: 'Rename Sheet', onClick: () => void cmd.renameWorksheet(ctx) },
      { title: 'Move or Copy Sheet...', onClick: () => ctx.openDialog({ kind: 'move-copy-sheet', index: ctx.doc.activeIndex }) },
      { type: 'header', title: 'Protection' },
      ctx.doc.sheet.protection
        ? { title: 'Unprotect Sheet...', icon: 'lock_open', onClick: () => ctx.openDialog({ kind: 'unprotect-sheet' }) }
        : { title: 'Protect Sheet...', icon: 'lock', onClick: () => ctx.openDialog({ kind: 'protect-sheet' }) },
      { title: 'Lock Cell', checked: ctx.doc.styleAt(ctx.doc.selection.active.r, ctx.doc.selection.active.c).locked !== false, onClick: () => cmd.format(ctx, { locked: ctx.doc.styleAt(ctx.doc.selection.active.r, ctx.doc.selection.active.c).locked === false }, 'Lock cell') },
      { title: 'Format Cells...', icon: 'format_cells', shortcut: 'Ctrl+1', onClick: () => ctx.openDialog({ kind: 'format-cells' }) },
    ],
  } as Command,
  autosum: {
    id: 'autosum',
    label: 'AutoSum',
    icon: 'functions',
    shortcut: 'Alt+=',
    run: (ctx) => cmd.autoSum(ctx, 'SUM'),
    menu: (ctx) => [
      { title: 'Sum', onClick: () => cmd.autoSum(ctx, 'SUM') },
      { title: 'Average', onClick: () => cmd.autoSum(ctx, 'AVERAGE') },
      { title: 'Count Numbers', onClick: () => cmd.autoSum(ctx, 'COUNT') },
      { title: 'Max', onClick: () => cmd.autoSum(ctx, 'MAX') },
      { title: 'Min', onClick: () => cmd.autoSum(ctx, 'MIN') },
      { type: 'line' },
      { title: 'More Functions...', onClick: () => ctx.openDialog({ kind: 'insert-function' }) },
    ],
  } as Command,
  fill: {
    id: 'fill',
    label: 'Fill',
    icon: 'fill_down',
    menu: (ctx) => [
      { title: 'Down', shortcut: 'Ctrl+D', onClick: () => cmd.fillDirection(ctx, 'down') },
      { title: 'Right', shortcut: 'Ctrl+R', onClick: () => cmd.fillDirection(ctx, 'right') },
      { title: 'Up', onClick: () => cmd.fillDirection(ctx, 'up') },
      { title: 'Left', onClick: () => cmd.fillDirection(ctx, 'left') },
      { type: 'line' },
      { title: 'Series...', onClick: () => ctx.openDialog({ kind: 'series' }) },
    ],
  } as Command,
  clear: {
    id: 'clear',
    label: 'Clear',
    icon: 'format_clear',
    menu: (ctx) => [
      { title: 'Clear All', testId: 'clear-all', onClick: () => cmd.clear(ctx, 'all') },
      { title: 'Clear Formats', testId: 'clear-formats', onClick: () => cmd.clear(ctx, 'formats') },
      { title: 'Clear Contents', testId: 'clear-contents', shortcut: 'Delete', onClick: () => cmd.clear(ctx, 'contents') },
      { title: 'Clear Comments and Notes', onClick: () => cmd.clear(ctx, 'comments') },
      { title: 'Clear Hyperlinks', onClick: () => cmd.clear(ctx, 'links') },
    ],
  } as Command,
  sortFilter: {
    id: 'sort-filter',
    label: 'Sort & Filter',
    icon: 'filter_alt',
    menu: (ctx) => [
      { title: 'Sort A to Z', icon: 'sort_by_alpha', onClick: () => cmd.sortSelection(ctx, true) },
      { title: 'Sort Z to A', icon: 'sort', onClick: () => cmd.sortSelection(ctx, false) },
      { title: 'Custom Sort...', onClick: () => ctx.openDialog({ kind: 'sort' }) },
      { type: 'line' },
      { title: 'Filter', icon: 'filter_alt', checked: ctx.state.hasFilter, shortcut: 'Ctrl+Shift+L', onClick: () => cmd.toggleFilter(ctx) },
      { title: 'Clear', icon: 'filter_alt_off', onClick: () => cmd.clearFilter(ctx) },
      { title: 'Reapply', icon: 'replay', onClick: () => cmd.reapplyFilter(ctx) },
    ],
  } as Command,
  find: {
    id: 'find-select',
    label: 'Find & Select',
    icon: 'search',
    menu: (ctx) => [
      { title: 'Find...', icon: 'search', shortcut: 'Ctrl+F', testId: 'find-dialog', onClick: () => ctx.openDialog({ kind: 'find', tab: 'find' }) },
      { title: 'Replace...', icon: 'find_replace', shortcut: 'Ctrl+H', onClick: () => ctx.openDialog({ kind: 'find', tab: 'replace' }) },
      { title: 'Go To...', shortcut: 'Ctrl+G', onClick: () => ctx.openDialog({ kind: 'find', tab: 'goto' }) },
      { title: 'Go To Special...', onClick: () => ctx.openDialog({ kind: 'goto-special' }) },
      { type: 'line' },
      { title: 'Formulas', onClick: () => ctx.doc.selectSpecial('formulas') },
      { title: 'Comments & Notes', onClick: () => ctx.doc.selectSpecial('comments') },
      { title: 'Conditional Formatting', onClick: () => ctx.doc.selectSpecial('condFormats') },
      { title: 'Constants', onClick: () => ctx.doc.selectSpecial('constants') },
      { title: 'Data Validation', onClick: () => ctx.doc.selectSpecial('validation') },
      { type: 'line' },
      { title: 'Selection Pane...', icon: 'layers', onClick: () => ctx.setPanel('selection') },
    ],
  } as Command,
  fontSettings: { id: 'font-settings', label: 'Font Settings', tip: 'Font Settings (Ctrl+Shift+F)', run: (ctx) => ctx.openDialog({ kind: 'format-cells', tab: 'font' }) } as Command,
  alignSettings: { id: 'alignment-settings', label: 'Alignment Settings', run: (ctx) => ctx.openDialog({ kind: 'format-cells', tab: 'alignment' }) } as Command,
  numberSettings: { id: 'number-settings', label: 'Number Format', tip: 'Number Format (Ctrl+1)', run: (ctx) => ctx.openDialog({ kind: 'format-cells', tab: 'number' }) } as Command,
};

const fontCombo: Control = {
  type: 'combo',
  id: 'font-name',
  label: 'Font',
  width: 132,
  editable: true,
  value: (ctx) => fmt(ctx).fontName,
  options: (ctx) => {
    const theme = ctx.doc.wb.theme;
    const list = [...new Set([theme.majorFont, theme.minorFont, ...FONT_NAMES])];
    return list.map((f) => ({ value: f, label: f === theme.majorFont ? `${f} (Headings)` : f === theme.minorFont ? `${f} (Body)` : f, font: `"${f}"` }));
  },
  onChange: (ctx, v) => cmd.setFontName(ctx, v),
};

const sizeCombo: Control = {
  type: 'combo',
  id: 'font-size',
  label: 'Font Size',
  width: 54,
  editable: true,
  value: (ctx) => String(fmt(ctx).fontSize),
  options: () => FONT_SIZES.map((s) => ({ value: String(s) })),
  onChange: (ctx, v) => cmd.setFontSize(ctx, Number(v)),
};

const numberCombo: Control = {
  type: 'combo',
  id: 'number-format',
  label: 'Number Format',
  width: 116,
  value: (ctx) => numberFormatLabel(fmt(ctx).numFmt),
  options: () => [...NUMBER_FORMATS.map((f) => ({ value: f.label })), { value: 'More Number Formats...' }],
  onChange: (ctx, v) => {
    if (v === 'More Number Formats...') {
      ctx.openDialog({ kind: 'format-cells', tab: 'number' });
      return;
    }
    const f = NUMBER_FORMATS.find((x) => x.label === v);
    if (f) cmd.setNumberFormat(ctx, f.code);
  },
};

export const fillColor: Control = {
  type: 'color',
  cmd: { id: 'fill-color', label: 'Fill Color', icon: 'format_color_fill' },
  value: (ctx) => fmt(ctx).fill,
  apply: (ctx, c) => cmd.setFillColor(ctx, c),
  noneLabel: 'No Fill',
  initial: '#FFFF00',
};

export const fontColor: Control = {
  type: 'color',
  cmd: { id: 'font-color', label: 'Font Color', icon: 'format_color_text' },
  value: (ctx) => fmt(ctx).color,
  apply: (ctx, c) => cmd.setFontColor(ctx, c),
  noneLabel: 'Automatic',
  initial: '#FF0000',
};

export const homeTab: Tab = {
  id: 'home',
  label: 'Home',
  groups: () => [
    { id: 'undo', label: 'Undo', icon: 'undo', controls: [{ type: 'stack', items: [{ type: 'small', cmd: C.undo, split: true, iconOnly: true }, { type: 'small', cmd: C.redo, split: true, iconOnly: true }] }] },
    {
      id: 'clipboard',
      label: 'Clipboard',
      icon: 'content_paste',
      controls: [{ type: 'large', cmd: C.paste, split: true }, { type: 'stack', items: [{ type: 'small', cmd: C.cut, iconOnly: true }, { type: 'small', cmd: C.copy, iconOnly: true }, { type: 'small', cmd: C.painter, iconOnly: true }] }],
    },
    {
      id: 'font',
      label: 'Font',
      icon: 'format_bold',
      launcher: C.fontSettings,
      controls: [
        {
          type: 'stack',
          items: [
            { type: 'row', items: [fontCombo, sizeCombo, { type: 'small', cmd: C.growFont, iconOnly: true }, { type: 'small', cmd: C.shrinkFont, iconOnly: true }] },
            { type: 'row', items: [{ type: 'small', cmd: C.bold, iconOnly: true }, { type: 'small', cmd: C.italic, iconOnly: true }, { type: 'small', cmd: C.underline, iconOnly: true }, { type: 'small', cmd: C.doubleUnderline, iconOnly: true }, { type: 'small', cmd: C.borders, split: true, iconOnly: true }, fillColor, fontColor] },
          ],
        },
      ],
    },
    {
      id: 'alignment',
      label: 'Alignment',
      icon: 'format_align_left',
      launcher: C.alignSettings,
      controls: [
        {
          type: 'stack',
          items: [
            { type: 'row', items: [{ type: 'small', cmd: C.alignTop, iconOnly: true }, { type: 'small', cmd: C.alignMiddle, iconOnly: true }, { type: 'small', cmd: C.alignBottom, iconOnly: true }, { type: 'small', cmd: C.orientation, iconOnly: true }] },
            { type: 'row', items: [{ type: 'small', cmd: C.alignLeft, iconOnly: true }, { type: 'small', cmd: C.alignCenter, iconOnly: true }, { type: 'small', cmd: C.alignRight, iconOnly: true }, { type: 'small', cmd: C.indentDec, iconOnly: true }, { type: 'small', cmd: C.indentInc, iconOnly: true }] },
          ],
        },
        { type: 'stack', items: [{ type: 'small', cmd: C.wrap }, { type: 'small', cmd: C.merge, split: true }] },
      ],
    },
    {
      id: 'number',
      label: 'Number',
      icon: 'percent',
      launcher: C.numberSettings,
      controls: [
        {
          type: 'stack',
          items: [
            numberCombo,
            { type: 'row', items: [{ type: 'small', cmd: C.accounting, split: true, iconOnly: true }, { type: 'small', cmd: C.percent, iconOnly: true }, { type: 'small', cmd: C.comma, iconOnly: true }, { type: 'small', cmd: C.decInc, iconOnly: true }, { type: 'small', cmd: C.decDec, iconOnly: true }] },
          ],
        },
      ],
    },
    {
      id: 'styles',
      label: 'Styles',
      icon: 'cell_styles',
      controls: [{ type: 'large', cmd: C.condFormat }, { type: 'large', cmd: C.formatTable }, { type: 'large', cmd: C.cellStyles }],
    },
    {
      id: 'cells',
      label: 'Cells',
      icon: 'insert_cells',
      controls: [{ type: 'large', cmd: C.insert, split: true }, { type: 'large', cmd: C.delete, split: true }, { type: 'large', cmd: C.format }],
    },
    {
      id: 'editing',
      label: 'Editing',
      icon: 'functions',
      controls: [
        { type: 'stack', items: [{ type: 'small', cmd: C.autosum, split: true }, { type: 'small', cmd: C.fill }, { type: 'small', cmd: C.clear }] },
        { type: 'large', cmd: C.sortFilter },
        { type: 'large', cmd: C.find },
      ],
    },
  ],
  toolbar: () => [
    { type: 'small', cmd: C.undo, split: true, iconOnly: true },
    { type: 'small', cmd: C.paste, split: true, iconOnly: true },
    { type: 'small', cmd: C.painter, iconOnly: true },
    fontCombo,
    sizeCombo,
    { type: 'small', cmd: C.bold, iconOnly: true },
    { type: 'small', cmd: C.italic, iconOnly: true },
    { type: 'small', cmd: C.underline, iconOnly: true },
    { type: 'small', cmd: C.borders, split: true, iconOnly: true },
    fillColor,
    fontColor,
    { type: 'small', cmd: C.alignLeft, iconOnly: true },
    { type: 'small', cmd: C.alignCenter, iconOnly: true },
    { type: 'small', cmd: C.alignRight, iconOnly: true },
    { type: 'small', cmd: C.wrap, iconOnly: false },
    { type: 'small', cmd: C.merge, split: true, iconOnly: false },
    numberCombo,
    { type: 'small', cmd: C.decInc, iconOnly: true },
    { type: 'small', cmd: C.decDec, iconOnly: true },
    { type: 'small', cmd: C.condFormat },
    { type: 'small', cmd: C.formatTable },
    { type: 'small', cmd: C.cellStyles },
    { type: 'small', cmd: C.insert, split: true },
    { type: 'small', cmd: C.delete, split: true },
    { type: 'small', cmd: C.format },
    { type: 'small', cmd: C.autosum, split: true },
    { type: 'small', cmd: C.sortFilter },
    { type: 'small', cmd: C.find },
  ],
};
