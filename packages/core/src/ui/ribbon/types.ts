/**
 * The ribbon is described as data: tabs hold groups, groups hold controls, controls run commands.
 * The same description drives the full ribbon, the single-row toolbar layout, the phone layout
 * (group buttons opening bottom sheets) and the command palette.
 */
import type { IconName } from '../icons';
import type { MenuItem } from '../menuItems';
import type { UIContextValue } from '../context';
import type { ColorSpec } from '../../core/model/color';

export type Ctx = UIContextValue;

export interface Command {
  /** Unique id; buttons get data-testid `rb-<id>`. */
  id: string;
  label: string;
  /** Tooltip (defaults to the label). */
  tip?: string;
  icon?: IconName;
  /** Icon colour. */
  tint?: string;
  run?: (ctx: Ctx) => void;
  /** Drop-down menu items. */
  menu?: (ctx: Ctx) => MenuItem[];
  /** Drop-down with a gallery or palette, drawn by the interface. */
  panel?: PanelSpec;
  checked?: (ctx: Ctx) => boolean;
  disabled?: (ctx: Ctx) => boolean;
  /** Shortcut shown in tooltips and the command palette. */
  shortcut?: string;
  /** Keywords for the command palette search. */
  keywords?: string;
}

/** The galleries of ribbon drop-downs. */
export type PanelSpec =
  | { kind: 'paste' }
  /** `pick` applies the chosen style; `showNone` adds the "None" style first. */
  | { kind: 'table-styles'; showNone?: boolean; pick: (ctx: Ctx, name: string) => void }
  | { kind: 'cell-styles' }
  | { kind: 'shapes' }
  | { kind: 'change-shape' }
  | { kind: 'shape-styles' }
  | { kind: 'themes' }
  | { kind: 'sparkline-styles' };

/** Controls with fields of their own, drawn by the interface. */
export type ViewSpec =
  /** The name of the table at the selection (Table Design). */
  | { kind: 'table-name' }
  /** Height and width of the selected object (Shape and Picture Format). */
  | { kind: 'object-size' }
  /** Columns of check boxes under a heading (Page Layout > Sheet Options). */
  | { kind: 'sheet-options'; columns: { head: string; items: Command[] }[] }
  /** A separator between the groups of the toolbar layout. */
  | { kind: 'separator' };

export interface ComboOption {
  value: string;
  label?: string;
  /** CSS font family to preview the option (font names). */
  font?: string;
}

export type Control =
  /** Large button: 32px icon over the label. With a menu/panel it opens it (or `split`: icon runs, label opens). */
  | { type: 'large'; cmd: Command; split?: boolean }
  /** Small button: 16px icon with the label at its right (`iconOnly` hides the label). */
  | { type: 'small'; cmd: Command; split?: boolean; iconOnly?: boolean }
  /** Up to three small controls stacked vertically. */
  | { type: 'stack'; items: Control[] }
  /** Icon buttons side by side (B I U...). */
  | { type: 'row'; items: Control[] }
  | {
      type: 'combo';
      id: string;
      label: string;
      width: number;
      value: (ctx: Ctx) => string;
      options: (ctx: Ctx) => ComboOption[];
      onChange: (ctx: Ctx, value: string) => void;
      /** Free text entry (font size, scale). */
      editable?: boolean;
      disabled?: (ctx: Ctx) => boolean;
    }
  /** Split colour button: the icon applies the last colour, the arrow opens the palette. */
  | {
      type: 'color';
      cmd: Command;
      value: (ctx: Ctx) => string | null;
      apply: (ctx: Ctx, color: ColorSpec | null) => void;
      /** Label of the "no colour" entry ("No Fill", "Automatic"). */
      noneLabel: string;
      /** Colour shown under the icon before any pick. */
      initial: string;
    }
  /** Check box with a label (View tab, sheet options). */
  | { type: 'check'; cmd: Command }
  | { type: 'view'; id: string; view: ViewSpec };

export interface Group {
  id: string;
  label: string;
  /** Icon of the collapsed group button (phones). */
  icon: IconName;
  controls: Control[];
  /** Dialog launcher at the bottom-right corner of the group. */
  launcher?: Command;
}

export interface Tab {
  id: string;
  label: string;
  groups: (ctx: Ctx) => Group[];
  /** Contextual tabs (Table Design, Chart Design...) show only while their object is selected. */
  contextual?: boolean;
  visible?: (ctx: Ctx) => boolean;
  /** Single-row toolbar layout of this tab; defaults to every control as a small button. */
  toolbar?: (ctx: Ctx) => Control[];
}

/** Every command reachable from a control (for the command palette and phone sheets). */
export function commandsOf(controls: Control[]): Command[] {
  const out: Command[] = [];
  const visit = (c: Control) => {
    switch (c.type) {
      case 'large':
      case 'small':
      case 'check':
      case 'color':
        out.push(c.cmd);
        break;
      case 'stack':
      case 'row':
        c.items.forEach(visit);
        break;
      default:
        break;
    }
  };
  controls.forEach(visit);
  return out;
}

/** Whether a command shows as checked (not when its test fails). */
export function commandChecked(cmd: Command, ctx: Ctx): boolean {
  try {
    return !!cmd.checked?.(ctx);
  } catch {
    return false;
  }
}

/** Whether a command is disabled (not when its test fails). */
export function commandDisabled(cmd: Command, ctx: Ctx): boolean {
  try {
    return !!cmd.disabled?.(ctx);
  } catch {
    return false;
  }
}

/** Tooltip of a command, with its shortcut. */
export function commandTip(cmd: Command): string {
  const tip = cmd.tip ?? cmd.label;
  return cmd.shortcut ? `${tip} (${cmd.shortcut})` : tip;
}

/** Runs a command from the ribbon; a cell being edited is committed first (except by font commands and Undo). */
export function runRibbonCommand(cmd: Command, ctx: Ctx): void {
  if (ctx.grid?.isEditing() && !cmd.id.startsWith('font') && cmd.id !== 'undo') ctx.grid.commitEdit();
  cmd.run?.(ctx);
}

/** Controls as small buttons in one row: large buttons become small, stacks and rows are flattened. */
export function smallControls(controls: Control[]): Control[] {
  const out: Control[] = [];
  for (const c of controls) {
    if (c.type === 'large') out.push({ type: 'small', cmd: c.cmd, split: c.split });
    else if (c.type === 'stack' || c.type === 'row') out.push(...smallControls(c.items));
    else out.push(c);
  }
  return out;
}

/** The controls of a tab in the toolbar layout: its own toolbar, or every group's as small buttons between separators. */
export function toolbarControls(tab: Tab, ctx: Ctx): Control[] {
  if (tab.toolbar) return tab.toolbar(ctx);
  return tab.groups(ctx).flatMap((g, i) => [...(i ? [{ type: 'view', id: `sep${i}`, view: { kind: 'separator' } } as Control] : []), ...smallControls(g.controls)]);
}

export const TINT = {
  blue: '#2b6cb0',
  green: '#217346',
  red: '#c43e1c',
  orange: '#d9822b',
  purple: '#7a4fa3',
  gray: '#555b61',
  teal: '#11847a',
};
