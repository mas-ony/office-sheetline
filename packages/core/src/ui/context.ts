/** What the interface's controls work with: the document, the grid, dialogs, side panels, notices and file actions. */
import type { SpreadsheetDocument } from '../core/document';
import type { EditState, GridView } from '../grid/GridView';
import type { DocState } from './docState';
import type { SheetlineUI } from './controller';

/** The presets of Home > Conditional Formatting that open a dialog. */
export type CfPresetKind = 'greater' | 'less' | 'between' | 'equal' | 'text' | 'date' | 'duplicate' | 'top10' | 'top10pct' | 'bottom10' | 'bottom10pct' | 'above' | 'below';

export type ExportFormat = 'xlsx' | 'xltx' | 'xlsm' | 'xltm' | 'csv' | 'pdf';

export interface FileActions {
  newWorkbook(): void;
  /** Opens a spreadsheet file from the device. */
  open(): void;
  /** Opens a workbook from bytes (XLSX, XLSM, XLTX, XLS, XLSB, CSV) under a file name; resolves to whether it could be read. */
  openData(bytes: Uint8Array, name: string): Promise<boolean>;
  /** Ctrl+S: writes the workbook (to the opened file where the browser allows, else downloads). */
  save(): void;
  /** Writes a copy in the given format with the given base name. */
  exportAs(format: ExportFormat, name?: string): void;
  /** Imports a file into a new workbook (same as open; CSV options apply). */
  importFile(kind: 'excel' | 'csv'): void;
  print(): void;
  rename(name: string): void;
  share?: () => void;
}

export type DialogSpec =
  | { kind: 'format-cells'; tab?: 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection' }
  | { kind: 'cf-preset'; preset: CfPresetKind }
  | { kind: 'cf-rule'; ruleId?: string; ruleType?: string }
  | { kind: 'cf-manager' }
  | { kind: 'create-table'; style?: string }
  | { kind: 'resize-table'; tableId: number }
  | { kind: 'insert-chart'; category?: string; changeId?: string }
  | { kind: 'select-data'; chartId: string }
  | { kind: 'move-chart'; chartId: string }
  | { kind: 'sparkline'; type: 'line' | 'column' | 'stacked'; groupId?: string; single?: boolean }
  | { kind: 'hyperlink' }
  | { kind: 'insert-function'; category?: string }
  | { kind: 'name-manager' }
  | { kind: 'new-name'; name?: string; scope?: number }
  | { kind: 'sort' }
  | { kind: 'text-to-columns' }
  | { kind: 'validation' }
  | { kind: 'remove-duplicates' }
  | { kind: 'group'; ungroup?: boolean }
  | { kind: 'subtotal' }
  | { kind: 'goal-seek' }
  | { kind: 'data-table' }
  | { kind: 'scenarios' }
  | { kind: 'find'; tab?: 'find' | 'replace' | 'goto' }
  | { kind: 'goto-special' }
  | { kind: 'zoom' }
  | { kind: 'page-setup'; tab?: 'page' | 'margins' | 'header' | 'sheet' }
  | { kind: 'insert-cells' }
  | { kind: 'delete-cells' }
  | { kind: 'size'; what: 'row-height' | 'col-width' | 'default-height' | 'default-width' }
  | { kind: 'protect-sheet' }
  | { kind: 'unprotect-sheet' }
  | { kind: 'protect-workbook' }
  | { kind: 'move-copy-sheet'; index: number }
  | { kind: 'unhide-sheet' }
  | { kind: 'paste-special' }
  | { kind: 'series' }
  | { kind: 'spread-settings'; tab?: string }
  | { kind: 'sheet-settings'; tab?: string }
  | { kind: 'customize-colors' }
  | { kind: 'alt-text'; id: string }
  | { kind: 'command-palette' }
  | { kind: 'new-cell-style' }
  | { kind: 'shortcuts' }
  | { kind: 'symbol' }
  | { kind: 'shape-text'; id: string }
  | { kind: 'names-from-selection' }
  | { kind: 'outline-settings' };

export type SidePanel = 'comments' | 'selection' | 'format' | null;

export interface ConfirmOptions {
  title?: string;
  ok?: string;
  cancel?: string;
  /** A third answer (Yes / No / Cancel): resolves 'alt'. */
  alt?: string;
  danger?: boolean;
}

export interface ViewSettings {
  showFormulas: boolean;
  tabStrip: boolean;
  newTab: boolean;
  focusCell: boolean;
  formulaBar: boolean;
  statusBar: boolean;
}

/** Spread Settings that are browser preferences (not stored in the file). */
export interface AppSettings {
  allowFill: boolean;
  allowZoom: boolean;
  showHScroll: boolean;
  showVScroll: boolean;
  formulaHints: boolean;
}

export const DEFAULT_SETTINGS: AppSettings = { allowFill: true, allowZoom: true, showHScroll: true, showVScroll: true, formulaHints: true };

export interface UIContextValue {
  doc: SpreadsheetDocument;
  grid: GridView | null;
  state: DocState;
  /** The cell edit in progress (mirrored by the formula bar). */
  edit: EditState;
  phone: boolean;
  files: FileActions;
  dialog: DialogSpec | null;
  openDialog(d: DialogSpec): void;
  closeDialog(): void;
  panel: SidePanel;
  setPanel(p: SidePanel): void;
  notify(message: string): void;
  confirm(message: string, opts?: ConfirmOptions): Promise<boolean>;
  /** Yes / No / Cancel. */
  ask(message: string, opts?: ConfirmOptions): Promise<'ok' | 'alt' | 'cancel'>;
  prompt(message: string, value?: string, opts?: ConfirmOptions): Promise<string | null>;
  /** Threaded comment card at a cell. */
  openComments(r: number, c: number, startNew?: boolean): void;
  /** Note editor at a cell (legacy notes). */
  openNote(r: number, c: number): void;
  /** Filter drop-down of a column of an auto filter or table. */
  openFilter(col: number, at: { x: number; y: number }, tableId?: number): void;
  /** Chart, shape or picture currently selected. */
  selectedObject: string | null;
  selectObject(id: string | null): void;
  /** File backstage. */
  backstage: string | null;
  openBackstage(page?: string): void;
  closeBackstage(): void;
  view: ViewSettings;
  setView(patch: Partial<ViewSettings>): void;
  settings: AppSettings;
  setSettings(patch: Partial<AppSettings>): void;
  /** Format painter: armed (once or sticky). */
  painter: 'once' | 'sticky' | null;
  setPainter(mode: 'once' | 'sticky' | null): void;
  /** Runs a ribbon command by id (command palette, shortcuts). */
  runCommand(id: string): void;
  /** Activates a ribbon tab. */
  showTab(id: string): void;
  /** The interface this context belongs to. */
  ui: SheetlineUI;
  /** Where menus, dialogs and messages are rendered (a container on the page's body). */
  layer: HTMLElement;
  /** Settings of the interface that are kept between visits (with the storageKey option). */
  prefs: Prefs;
}

/** Remembered values (recent colours, cell styles); kept in localStorage with the storageKey option. */
export interface Prefs {
  get<T>(name: string, initial: T): T;
  set(name: string, value: unknown): void;
}
