/** Data tab: Sort & Filter, Data Tools, Forecast (What-If Analysis), Outline. */
import * as cmd from '../commands';
import { TINT, type Command, type Ctx, type Tab } from './types';
import { isFullColumns, isFullRows } from '../../core/refs';

/** Rows (or columns) of the selection to group; whole rows/columns pick the axis directly. */
function groupTarget(ctx: Ctx): { axis: 'row' | 'col'; from: number; to: number } | null {
  const r = ctx.doc.range;
  if (isFullRows(r)) return { axis: 'row', from: r.r0, to: r.r1 };
  if (isFullColumns(r)) return { axis: 'col', from: r.c0, to: r.c1 };
  return null;
}

function group(ctx: Ctx, ungroup: boolean): void {
  cmd.commit(ctx);
  const t = groupTarget(ctx);
  if (!t) {
    ctx.openDialog({ kind: 'group', ungroup });
    return;
  }
  if (ungroup) {
    if (!ctx.doc.ungroupLines(t.axis, t.from, t.to)) ctx.notify('The selection is not grouped.');
    return;
  }
  const error = ctx.doc.groupLines(t.axis, t.from, t.to);
  if (error) ctx.notify(error);
}

function circleInvalid(ctx: Ctx): void {
  const g = ctx.grid;
  if (!g) return;
  const cells = ctx.doc.invalidCells();
  g.circles = cells;
  g.requestRender();
  if (!cells.length) ctx.notify('No invalid data was found.');
}

const ro = (ctx: Ctx) => ctx.state.readOnly;

export const D = {
  sortAsc: { id: 'sort-asc', label: 'Sort A to Z', tip: 'Sort Smallest to Largest', icon: 'sort_by_alpha', run: (ctx) => cmd.sortSelection(ctx, true), disabled: ro } as Command,
  sortDesc: { id: 'sort-desc', label: 'Sort Z to A', tip: 'Sort Largest to Smallest', icon: 'sort', run: (ctx) => cmd.sortSelection(ctx, false), disabled: ro } as Command,
  sort: { id: 'custom-sort', label: 'Sort', tip: 'Custom Sort', icon: 'swap_vert', tint: TINT.blue, run: (ctx) => ctx.openDialog({ kind: 'sort' }), disabled: ro } as Command,
  filter: { id: 'filter', label: 'Filter', icon: 'filter_alt', tint: TINT.blue, shortcut: 'Ctrl+Shift+L', keywords: 'auto filter', checked: (ctx) => ctx.state.hasFilter, run: (ctx) => cmd.toggleFilter(ctx), disabled: ro } as Command,
  clearFilter: { id: 'filter-clear', label: 'Clear', tip: 'Clear filters', icon: 'filter_alt_off', run: (ctx) => cmd.clearFilter(ctx), disabled: (ctx) => !ctx.state.hasFilter } as Command,
  reapply: { id: 'filter-reapply', label: 'Reapply', icon: 'replay', run: (ctx) => cmd.reapplyFilter(ctx), disabled: (ctx) => !ctx.state.hasFilter } as Command,
  textToColumns: { id: 'text-to-columns', label: 'Text to Columns', icon: 'text_to_columns', tint: TINT.blue, keywords: 'split delimiter', run: (ctx) => ctx.openDialog({ kind: 'text-to-columns' }), disabled: ro } as Command,
  removeDuplicates: { id: 'remove-duplicates', label: 'Remove Duplicates', icon: 'remove_duplicates', tint: TINT.blue, keywords: 'unique', run: (ctx) => ctx.openDialog({ kind: 'remove-duplicates' }), disabled: ro } as Command,
  validation: {
    id: 'data-validation',
    label: 'Data Validation',
    icon: 'rule',
    tint: TINT.green,
    run: (ctx) => ctx.openDialog({ kind: 'validation' }),
    menu: (ctx) => [
      { title: 'Data Validation...', icon: 'rule', testId: 'validation-dialog', onClick: () => ctx.openDialog({ kind: 'validation' }) },
      { title: 'Circle Invalid Data', icon: 'radio_button_unchecked', testId: 'validation-circle', onClick: () => circleInvalid(ctx) },
      {
        title: 'Clear Validation Circles',
        icon: 'block',
        disabled: !ctx.grid?.circles?.length,
        onClick: () => {
          if (!ctx.grid) return;
          ctx.grid.circles = null;
          ctx.grid.requestRender();
        },
      },
    ],
    disabled: ro,
  } as Command,
  whatIf: {
    id: 'what-if',
    label: 'What-If Analysis',
    icon: 'what_if',
    tint: TINT.purple,
    menu: (ctx) => [
      { title: 'Scenario Manager...', icon: 'view_agenda', testId: 'scenario-manager', onClick: () => ctx.openDialog({ kind: 'scenarios' }) },
      { title: 'Goal Seek...', icon: 'goal_seek', testId: 'goal-seek', onClick: () => ctx.openDialog({ kind: 'goal-seek' }) },
      { title: 'Data Table...', icon: 'table_view', testId: 'data-table', onClick: () => ctx.openDialog({ kind: 'data-table' }) },
    ],
    disabled: ro,
  } as Command,
  group: {
    id: 'group',
    label: 'Group',
    icon: 'group_rows',
    shortcut: 'Shift+Alt+Right',
    run: (ctx) => group(ctx, false),
    menu: (ctx) => [
      { title: 'Group...', icon: 'group_rows', onClick: () => group(ctx, false) },
      {
        title: 'Auto Outline',
        onClick: () => {
          if (!ctx.doc.autoOutline()) ctx.notify('Cannot create an outline: no summary formulas (SUM of the cells above or to the left) were found.');
        },
      },
    ],
    disabled: ro,
  } as Command,
  ungroup: {
    id: 'ungroup',
    label: 'Ungroup',
    icon: 'ungroup_rows',
    shortcut: 'Shift+Alt+Left',
    run: (ctx) => group(ctx, true),
    menu: (ctx) => [
      { title: 'Ungroup...', icon: 'ungroup_rows', onClick: () => group(ctx, true) },
      { title: 'Clear Outline', onClick: () => ctx.doc.clearOutline() },
    ],
    disabled: ro,
  } as Command,
  subtotal: { id: 'subtotal', label: 'Subtotal', icon: 'subtotal', run: (ctx) => ctx.openDialog({ kind: 'subtotal' }), disabled: ro } as Command,
  showDetail: { id: 'show-detail', label: 'Show Detail', icon: 'unfold_more', run: (ctx) => { if (!ctx.doc.showDetail(true)) ctx.notify('There is no hidden detail at the selected cell.'); } } as Command,
  hideDetail: { id: 'hide-detail', label: 'Hide Detail', icon: 'unfold_less', run: (ctx) => { if (!ctx.doc.showDetail(false)) ctx.notify('The selected cell is not in a group.'); } } as Command,
  outlineSettings: { id: 'outline-settings', label: 'Outline Settings', run: (ctx) => ctx.openDialog({ kind: 'outline-settings' }) } as Command,
};

export const dataTab: Tab = {
  id: 'data',
  label: 'Data',
  groups: () => [
    {
      id: 'sort-filter',
      label: 'Sort & Filter',
      icon: 'filter_alt',
      controls: [
        { type: 'stack', items: [{ type: 'small', cmd: D.sortAsc, iconOnly: true }, { type: 'small', cmd: D.sortDesc, iconOnly: true }] },
        { type: 'large', cmd: D.sort },
        { type: 'large', cmd: D.filter },
        { type: 'stack', items: [{ type: 'small', cmd: D.clearFilter }, { type: 'small', cmd: D.reapply }] },
      ],
    },
    {
      id: 'data-tools',
      label: 'Data Tools',
      icon: 'rule',
      controls: [{ type: 'large', cmd: D.textToColumns }, { type: 'large', cmd: D.removeDuplicates }, { type: 'large', cmd: D.validation, split: true }],
    },
    { id: 'forecast', label: 'Forecast', icon: 'what_if', controls: [{ type: 'large', cmd: D.whatIf }] },
    {
      id: 'outline',
      label: 'Outline',
      icon: 'group_rows',
      launcher: D.outlineSettings,
      controls: [
        { type: 'large', cmd: D.group, split: true },
        { type: 'large', cmd: D.ungroup, split: true },
        { type: 'large', cmd: D.subtotal },
        { type: 'stack', items: [{ type: 'small', cmd: D.showDetail }, { type: 'small', cmd: D.hideDetail }] },
      ],
    },
  ],
  toolbar: () => [
    { type: 'small', cmd: D.sortAsc, iconOnly: true },
    { type: 'small', cmd: D.sortDesc, iconOnly: true },
    { type: 'small', cmd: D.sort },
    { type: 'small', cmd: D.filter },
    { type: 'small', cmd: D.clearFilter },
    { type: 'small', cmd: D.textToColumns },
    { type: 'small', cmd: D.removeDuplicates },
    { type: 'small', cmd: D.validation, split: true },
    { type: 'small', cmd: D.whatIf },
    { type: 'small', cmd: D.group, split: true },
    { type: 'small', cmd: D.ungroup, split: true },
    { type: 'small', cmd: D.subtotal },
  ],
};
