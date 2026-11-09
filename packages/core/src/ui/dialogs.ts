/** What the dialogs and the File backstage of both interfaces share: lists, labels and small helpers. */
import type { PasteWhat } from '../core/document';
import type { SpecialKind } from '../core/doc/select';
import type { UIContextValue } from './context';
import type { IconName } from './icons';
import type { MenuItem } from './menuItems';
import { allCommands } from './ribbon/tabs';

// -----------------------------------------------------------------------------------------------
// File backstage

/** The pages of the File backstage; a separator goes between groups. */
export const BACKSTAGE_PAGES: { id: string; label: string; group: number }[] = [
  { id: 'new', label: 'New', group: 0 },
  { id: 'open', label: 'Open', group: 1 },
  { id: 'save', label: 'Save', group: 1 },
  { id: 'import', label: 'Import', group: 2 },
  { id: 'export', label: 'Export', group: 2 },
  { id: 'print', label: 'Print', group: 3 },
  { id: 'info', label: 'Info', group: 4 },
];

/** The name of a workbook file without its extension. */
export function fileBaseName(name: string): string {
  return name.replace(/\.(xlsx|xlsm|xltx|xltm|xls|xlsb|csv|tsv|txt)$/i, '') || 'Book1';
}

/** The Scaling choice of File > Print matching the page setup's fit-to settings. */
export function printScalingKind(fit: boolean, w: number, h: number): string {
  if (!fit) return 'none';
  if (w === 1 && h === 1) return 'sheet';
  if (w === 1 && h === 0) return 'columns';
  if (w === 0 && h === 1) return 'rows';
  return 'custom';
}

/** The document properties edited on File > Info. */
export const DOC_PROPERTIES: { key: string; label: string; placeholder: string }[] = [
  { key: 'title', label: 'Title', placeholder: 'Add a title' },
  { key: 'keywords', label: 'Tags', placeholder: 'Add a tag' },
  { key: 'description', label: 'Comments', placeholder: 'Add comments' },
  { key: 'category', label: 'Categories', placeholder: 'Add a category' },
  { key: 'subject', label: 'Subject', placeholder: 'Specify the subject' },
  { key: 'hyperlinkBase', label: 'Hyperlink Base', placeholder: 'Add text' },
  { key: 'company', label: 'Company', placeholder: 'Specify the company' },
  { key: 'manager', label: 'Manager', placeholder: 'Specify the manager' },
  { key: 'creator', label: 'Author', placeholder: 'Add an author' },
];

// -----------------------------------------------------------------------------------------------
// Editing dialogs

/** The choices of Go To Special. */
export const GOTO_SPECIALS: { kind: SpecialKind; label: string }[] = [
  { kind: 'comments', label: 'Comments' },
  { kind: 'constants', label: 'Constants' },
  { kind: 'formulas', label: 'Formulas' },
  { kind: 'blanks', label: 'Blanks' },
  { kind: 'currentRegion', label: 'Current region' },
  { kind: 'currentArray', label: 'Current array' },
  { kind: 'objects', label: 'Objects' },
  { kind: 'rowDifferences', label: 'Row differences' },
  { kind: 'columnDifferences', label: 'Column differences' },
  { kind: 'lastCell', label: 'Last cell' },
  { kind: 'visible', label: 'Visible cells only' },
  { kind: 'condFormats', label: 'Conditional formats' },
  { kind: 'validation', label: 'Data validation' },
];

/** What Paste Special can paste. */
export const PASTE_SPECIAL_WHAT: { what: PasteWhat; label: string }[] = [
  { what: 'all', label: 'All' },
  { what: 'formulas', label: 'Formulas' },
  { what: 'values', label: 'Values' },
  { what: 'formats', label: 'Formats' },
  { what: 'comments', label: 'Comments and Notes' },
  { what: 'validation', label: 'Validation' },
  { what: 'allExceptBorders', label: 'All except borders' },
  { what: 'colWidths', label: 'Column widths' },
  { what: 'formulasNumFmt', label: 'Formulas and number formats' },
  { what: 'valuesNumFmt', label: 'Values and number formats' },
];

/** The characters of the Symbol dialog. */
export const SYMBOL_CHARS = '€£¥¢©®™°±×÷µ¶§•…†‡‰′″‹›«»–—‘’“”¹²³¼½¾αβγδεθλμπσφωΔΣΩ∞≈≠≤≥√∑∏∫∂∆∇←↑→↓↔⇒⇔✓✔✗✘★☆♠♣♥♦☐☑☒⌘⌥⏎'.split('');

// -----------------------------------------------------------------------------------------------
// Command Palette and the shortcut list

/** A palette entry: a ribbon command, or an item of a command's menu ("Freeze Panes › Freeze Top Row"). */
export interface PaletteEntry {
  id: string;
  label: string;
  icon?: IconName;
  shortcut?: string;
  hay: string;
  run: () => void;
}

/** Everything the Command Palette can run, enabled now. */
export function paletteEntries(ui: UIContextValue): PaletteEntry[] {
  const out: PaletteEntry[] = [];
  for (const c of allCommands(ui)) {
    if (c.disabled?.(ui)) continue;
    if (c.run) out.push({ id: c.id, label: c.label, icon: c.icon, shortcut: c.shortcut, hay: `${c.label} ${c.tip ?? ''} ${c.keywords ?? ''}`.toLowerCase(), run: () => c.run!(ui) });
    if (!c.menu) continue;
    const walk = (items: MenuItem[], path: string) => {
      for (const it of items) {
        if (it.type || !it.title || it.disabled) continue;
        const label = `${path} › ${it.title.replace(/\.\.\.$/, '')}`;
        if (it.submenu) walk(it.submenu, label);
        else if (it.onClick) out.push({ id: `${c.id}/${it.title}`, label, icon: it.icon ?? c.icon, shortcut: it.shortcut, hay: label.toLowerCase(), run: it.onClick });
      }
    };
    walk(c.menu(ui), c.label);
  }
  // Commands shared by several tabs (and their menus) appear once.
  return [...new Map(out.map((e) => [e.id, e])).values()];
}

/** The Keyboard Shortcuts list. */
export const KEYBOARD_SHORTCUTS: [string, string][] = [
  ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'],
  ['Ctrl+C / Ctrl+X / Ctrl+V', 'Copy / Cut / Paste'],
  ['Ctrl+Alt+V', 'Paste Special'],
  ['Ctrl+B / Ctrl+I / Ctrl+U', 'Bold / Italic / Underline'],
  ['Ctrl+5', 'Strikethrough'],
  ['Ctrl+1', 'Format Cells'],
  ['Ctrl+Shift+F', 'Format Cells: Font'],
  ['Ctrl+Shift+$ / % / !', 'Currency / Percent / Number format'],
  ['Ctrl+F / Ctrl+H', 'Find / Replace'],
  ['Ctrl+G or F5', 'Go To'],
  ['Ctrl+K', 'Insert Link'],
  ['Ctrl+T', 'Create Table'],
  ['Ctrl+Shift+L', 'Filter'],
  ['Alt+=', 'AutoSum'],
  ['Alt+F1', 'Insert default chart'],
  ['Shift+F3', 'Insert Function'],
  ['Ctrl+F3', 'Name Manager'],
  ['F9 / Shift+F9', 'Calculate Now / Calculate Sheet'],
  ['Ctrl+`', 'Show Formulas'],
  ['Ctrl+9 / Ctrl+0', 'Hide rows / Hide columns'],
  ['Ctrl++ / Ctrl+-', 'Insert / Delete cells'],
  ['Ctrl+D / Ctrl+R', 'Fill Down / Fill Right'],
  ['Ctrl+; / Ctrl+Shift+;', 'Insert date / time'],
  ['Ctrl+Space / Shift+Space', 'Select column / row'],
  ['Ctrl+A', 'Select all'],
  ['F2', 'Edit cell'],
  ['Shift+F2', 'New or edit note'],
  ['Alt+Q', 'Command Palette'],
  ['Ctrl+PageUp / PageDown', 'Previous / next sheet'],
  ['Shift+F11', 'Insert sheet'],
  ['Ctrl+S / Ctrl+O / Ctrl+P', 'Save / Open / Print'],
];
