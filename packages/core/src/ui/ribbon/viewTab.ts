/** View tab: Show (headers, gridlines, tab strip, focus cell...), Zoom, Window (freeze panes), Panes. */
import * as cmd from '../commands';
import { TINT, type Command, type Ctx, type Tab } from './types';

const sheet = (ctx: Ctx) => ctx.doc.sheet;

export const V = {
  rowHeader: { id: 'view-row-header', label: 'Row Header', checked: (ctx) => sheet(ctx).showRowHeaders, run: (ctx) => ctx.doc.setSheetView({ showRowHeaders: !sheet(ctx).showRowHeaders }) } as Command,
  colHeader: { id: 'view-column-header', label: 'Column Header', checked: (ctx) => sheet(ctx).showColHeaders, run: (ctx) => ctx.doc.setSheetView({ showColHeaders: !sheet(ctx).showColHeaders }) } as Command,
  vGrid: { id: 'view-vertical-gridlines', label: 'Vertical Gridlines', checked: (ctx) => sheet(ctx).showGridLines && sheet(ctx).gridV, run: (ctx) => ctx.doc.setSheetView({ gridV: !(sheet(ctx).showGridLines && sheet(ctx).gridV), gridH: sheet(ctx).showGridLines && sheet(ctx).gridH }) } as Command,
  hGrid: { id: 'view-horizontal-gridlines', label: 'Horizontal Gridlines', checked: (ctx) => sheet(ctx).showGridLines && sheet(ctx).gridH, run: (ctx) => ctx.doc.setSheetView({ gridH: !(sheet(ctx).showGridLines && sheet(ctx).gridH), gridV: sheet(ctx).showGridLines && sheet(ctx).gridV }) } as Command,
  tabStrip: { id: 'view-tab-strip', label: 'Tab Strip', checked: (ctx) => ctx.view.tabStrip, run: (ctx) => ctx.setView({ tabStrip: !ctx.view.tabStrip }) } as Command,
  newTab: { id: 'view-new-tab', label: 'New Tab', checked: (ctx) => ctx.view.newTab, run: (ctx) => ctx.setView({ newTab: !ctx.view.newTab }) } as Command,
  formulaBar: { id: 'view-formula-bar', label: 'Formula Bar', checked: (ctx) => ctx.view.formulaBar, run: (ctx) => ctx.setView({ formulaBar: !ctx.view.formulaBar }) } as Command,
  statusBar: { id: 'view-status-bar', label: 'Status Bar', checked: (ctx) => ctx.view.statusBar, run: (ctx) => ctx.setView({ statusBar: !ctx.view.statusBar }) } as Command,
  zeros: { id: 'view-zero-values', label: 'Zero Values', checked: (ctx) => sheet(ctx).showZeros, run: (ctx) => ctx.doc.setSheetView({ showZeros: !sheet(ctx).showZeros }) } as Command,
  focusCell: { id: 'view-focus-cell', label: 'Focus Cell', icon: 'center_focus_strong', tint: TINT.green, checked: (ctx) => ctx.view.focusCell, run: (ctx) => ctx.setView({ focusCell: !ctx.view.focusCell }) } as Command,
  zoom: { id: 'zoom', label: 'Zoom', icon: 'zoom_in', tint: TINT.blue, run: (ctx) => ctx.openDialog({ kind: 'zoom' }) } as Command,
  zoom100: { id: 'zoom-100', label: '100%', tip: 'Zoom to 100%', icon: 'search', run: (ctx) => cmd.zoomTo(ctx, 1) } as Command,
  zoomSelection: { id: 'zoom-selection', label: 'Zoom to Selection', icon: 'fit_screen', run: (ctx) => cmd.zoomToSelection(ctx) } as Command,
  freeze: {
    id: 'freeze-panes',
    label: 'Freeze Panes',
    icon: 'freeze_panes',
    tint: TINT.blue,
    menu: (ctx) => {
      const f = sheet(ctx).freeze;
      const frozen = !!(f.rows || f.cols);
      return [
        frozen
          ? { title: 'Unfreeze Panes', tip: 'Unlock all rows and columns to scroll through the entire worksheet.', icon: 'freeze_panes', testId: 'unfreeze', onClick: () => cmd.freeze(ctx, 'none') }
          : { title: 'Freeze Panes', tip: 'Keep rows and columns visible while the rest of the worksheet scrolls (based on current selection).', icon: 'freeze_panes', testId: 'freeze-panes-item', onClick: () => cmd.freeze(ctx, 'panes') },
        { title: 'Freeze Top Row', icon: 'freeze_rows', testId: 'freeze-top-row', onClick: () => cmd.freeze(ctx, 'top-row') },
        { title: 'Freeze First Column', icon: 'freeze_cols', testId: 'freeze-first-col', onClick: () => cmd.freeze(ctx, 'first-col') },
      ];
    },
  } as Command,
  selectionPane: { id: 'view-selection-pane', label: 'Selection Pane', icon: 'layers', checked: (ctx) => ctx.panel === 'selection', run: (ctx) => ctx.setPanel(ctx.panel === 'selection' ? null : 'selection') } as Command,
  formatPane: { id: 'view-format-pane', label: 'Format Pane', icon: 'view_sidebar', checked: (ctx) => ctx.panel === 'format', disabled: (ctx) => !ctx.selectedObject, run: (ctx) => ctx.setPanel(ctx.panel === 'format' ? null : 'format') } as Command,
  palette: { id: 'command-palette', label: 'Command Palette', icon: 'terminal', tint: TINT.gray, shortcut: 'Alt+Q', keywords: 'search actions tell me', run: (ctx) => ctx.openDialog({ kind: 'command-palette' }) } as Command,
  shortcuts: { id: 'keyboard-shortcuts', label: 'Keyboard Shortcuts', icon: 'keyboard', run: (ctx) => ctx.openDialog({ kind: 'shortcuts' }) } as Command,
};

export const viewTab: Tab = {
  id: 'view',
  label: 'View',
  groups: () => [
    {
      id: 'show',
      label: 'Show',
      icon: 'visibility',
      controls: [
        { type: 'stack', items: [{ type: 'check', cmd: V.rowHeader }, { type: 'check', cmd: V.colHeader }, { type: 'check', cmd: V.formulaBar }] },
        { type: 'stack', items: [{ type: 'check', cmd: V.vGrid }, { type: 'check', cmd: V.hGrid }, { type: 'check', cmd: V.zeros }] },
        { type: 'stack', items: [{ type: 'check', cmd: V.tabStrip }, { type: 'check', cmd: V.newTab }, { type: 'check', cmd: V.statusBar }] },
        { type: 'large', cmd: V.focusCell },
      ],
    },
    {
      id: 'zoom',
      label: 'Zoom',
      icon: 'zoom_in',
      controls: [{ type: 'large', cmd: V.zoom }, { type: 'large', cmd: V.zoom100 }, { type: 'large', cmd: V.zoomSelection }],
    },
    { id: 'window', label: 'Window', icon: 'freeze_panes', controls: [{ type: 'large', cmd: V.freeze }] },
    {
      id: 'panes',
      label: 'Panes',
      icon: 'layers',
      controls: [{ type: 'large', cmd: V.selectionPane }, { type: 'large', cmd: V.formatPane }],
    },
    { id: 'tools', label: 'Tools', icon: 'terminal', controls: [{ type: 'large', cmd: V.palette }, { type: 'large', cmd: V.shortcuts }] },
  ],
  toolbar: () => [
    { type: 'check', cmd: V.rowHeader },
    { type: 'check', cmd: V.colHeader },
    { type: 'check', cmd: V.vGrid },
    { type: 'check', cmd: V.hGrid },
    { type: 'check', cmd: V.tabStrip },
    { type: 'small', cmd: V.focusCell },
    { type: 'small', cmd: V.zoom },
    { type: 'small', cmd: V.zoom100 },
    { type: 'small', cmd: V.freeze },
    { type: 'small', cmd: V.selectionPane },
    { type: 'small', cmd: V.palette },
  ],
};
