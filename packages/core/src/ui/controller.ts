/**
 * The editor without its views: the state the React and Vue interfaces render, the actions their
 * controls call, and what needs no view (the grid and its callbacks, file actions, autosave and
 * recovery, shortcuts, drag and drop, the phone layout). A view creates one, renders `getState()`
 * and `ctx`, calls `attach()` once its root element exists, `attachGrid()` with the grid's element
 * (again for each workbook opened: `docKey` changes) and `detach()` when it goes away.
 */
import type { SpreadsheetDocument } from '../core/document';
import { PROTECTED_MESSAGE } from '../core/doc/protection';
import { sheetToCsv } from '../core/csv';
import { GridView, type EditState } from '../grid/GridView';
import { clearMeasureCache } from '../grid/fonts';
import { clearTextCache } from '../grid/render';
import { ACCEPT, newDocument, openDocument } from '../api/open';
import { clearAutosave, loadAutosave, saveAutosave, type AutosaveRecord } from '../api/autosave';
import { createApi, type ApiEvents, type SpreadsheetApi } from '../api/api';
import { documentOf, downloadBlob, nameOfUrl, type DocumentInput, type DocumentSource } from '../api/source';
import { DEFAULT_SETTINGS, type AppSettings, type ConfirmOptions, type DialogSpec, type ExportFormat, type FileActions, type Prefs, type SidePanel, type UIContextValue, type ViewSettings } from './context';
import { computeDocState, DOC_STATE_EVENTS, type DocState } from './docState';
import { isTouchDevice, matchesMedia, PHONE_QUERY, watchMedia } from './dom';
import { FormulaAssistState } from './formulaAssist';
import type { MenuAnchor, MenuItem } from './menuItems';
import { cellMenu, columnMenu, cornerMenu, objectMenu, rowMenu, tabMenu } from './contextMenus';
import { allCommands, RIBBON_TABS } from './ribbon/tabs';
import { handleShortcut } from './shortcuts';
import * as cmd from './commands';

export type Layout = 'ribbon' | 'toolbar';
export type ThemeName = 'light' | 'dark';

/** Options of the SheetlineEditor components (React and Vue) and of mount(). */
export interface SheetlineOptions {
  /** Opened at start, and again whenever it changes. */
  document?: DocumentInput;
  /** Workbooks open view-only. */
  readOnly?: boolean;
  /** Ribbon or the one-row toolbar (desktop); by default the one last chosen, else the ribbon. */
  layout?: Layout;
  theme?: ThemeName;
  /** Shown on comments and notes, and written as the author of new workbooks. */
  author?: string;
  /** Keeps the layout, theme, settings and recent colours in localStorage under `<storageKey>-…`. */
  storageKey?: string;
  /** Keeps unsaved changes in an IndexedDB database of this name and offers to recover them. */
  autosave?: string;
  /** With `autosave`: offer recovery when the editor starts without a document (default true). */
  offerRecovery?: boolean;
  /** Replace some of the built-in file actions (New, Open, Save, Export, Print…). */
  files?: Partial<FileActions>;
  /** Handle the editor's shortcuts (Ctrl+1, Alt+Q, F9…) anywhere on the page, not only inside the editor. */
  globalShortcuts?: boolean;
  /** Called once the editor is ready and the initial document has opened. */
  onReady?: (api: SpreadsheetApi) => void;
  /** Called when the reader switches between the ribbon and the toolbar. */
  onLayoutChange?: (layout: Layout) => void;
  /** Called with the theme in use when the editor starts, and whenever the reader switches it. */
  onThemeChange?: (theme: ThemeName) => void;
}

/** A question shown as a dialog (confirm, Yes / No / Cancel, or a text prompt). */
export type PendingQuestion =
  | { kind: 'confirm'; message: string; opts?: ConfirmOptions; resolve: (v: boolean) => void }
  | { kind: 'ask'; message: string; opts?: ConfirmOptions; resolve: (v: 'ok' | 'alt' | 'cancel') => void }
  | { kind: 'prompt'; message: string; value: string; opts?: ConfirmOptions; resolve: (v: string | null) => void };

/** A context menu or list opened by the grid or the sheet tabs. */
export interface OpenMenu {
  items: MenuItem[];
  at: MenuAnchor;
  title?: string;
  testId?: string;
}

/** State of the interface around the workbook. */
export interface UIState {
  doc: SpreadsheetDocument;
  /** Changes with each workbook opened: views draw the editor for it afresh. */
  docKey: number;
  /** Counts the document's events; the document state is computed for it. */
  docVersion: number;
  grid: GridView | null;
  /** The cell edit in progress (mirrored by the formula bar). */
  edit: EditState;
  phone: boolean;
  dialog: DialogSpec | null;
  panel: SidePanel;
  /** The File backstage page, when it is open. */
  backstage: string | null;
  menu: OpenMenu | null;
  toast: string | null;
  pending: PendingQuestion | null;
  comments: { r: number; c: number; rect: DOMRect; startNew?: boolean } | null;
  note: { r: number; c: number; rect: DOMRect } | null;
  filter: { col: number; at: { x: number; y: number }; tableId?: number } | null;
  selectedObject: string | null;
  painter: 'once' | 'sticky' | null;
  view: ViewSettings;
  layout: Layout;
  theme: ThemeName;
  settings: AppSettings;
  /** The ribbon shows only its tabs. */
  collapsed: boolean;
  activeTab: string;
  /** Unsaved work found by autosave, offered for recovery. */
  restore: { fileName: string; savedAt: number } | null;
  /** A message that stays until dismissed (a file that could not be opened). */
  notice: string | null;
  /** A file is being read. */
  busy: string | null;
  /** Files are dragged over the editor. */
  dragging: boolean;
  /** File types of the hidden file input (Open, Import). */
  accept: string;
  /** The initial document has opened and the page API is available. */
  ready: boolean;
}

export const DISCARD_PROMPT = 'Discard unsaved changes to the current workbook?';

const DEFAULT_VIEW: ViewSettings = { showFormulas: false, tabStrip: true, newTab: true, focusCell: false, formulaBar: true, statusBar: true };
const NO_EDIT: EditState = { active: false, text: '', mode: 'enter', origin: 'cell', caret: 0, r: 0, c: 0 };
const TOAST_MS = 6000;
const AUTOSAVE_DELAY = 2500;
const AUTOSAVE_MAX_CELLS = 250_000;
const RECOVERY_MS = 1000 * 60 * 60 * 24 * 14;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const MIME: Record<Exclude<ExportFormat, 'csv' | 'pdf'>, string> = {
  xlsx: XLSX_MIME,
  xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
  xltm: 'application/vnd.ms-excel.template.macroEnabled.12',
};
const EXCEL_ACCEPT = '.xlsx,.xlsm,.xltx,.xltm,.xlsb,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel';
const CSV_ACCEPT = '.csv,.tsv,.txt,text/csv,text/plain';

interface FilePickerWindow {
  showOpenFilePicker?: (opts: unknown) => Promise<FileSystemFileHandle[]>;
}

/** The state of the interface that belongs to one workbook (reset when another opens). */
const PER_DOCUMENT = {
  grid: null,
  edit: NO_EDIT,
  dialog: null,
  panel: null,
  backstage: null,
  menu: null,
  toast: null,
  pending: null,
  comments: null,
  note: null,
  filter: null,
  selectedObject: null,
  painter: null,
  view: DEFAULT_VIEW,
  activeTab: 'home',
} satisfies Partial<UIState>;

function cellCount(doc: SpreadsheetDocument): number {
  let n = 0;
  for (const sheet of doc.wb.sheets) for (const row of sheet.cells.values()) n += row.size;
  return n;
}

const baseName = (d: SpreadsheetDocument) => d.fileName.replace(/\.(xlsx|xlsm|xltx|xltm|xls|xlsb|csv|tsv|txt)$/i, '') || 'Book1';

const isSheetFile = (f: File) => /\.(xlsx|xlsm|xltx|xltm|xlsb|xls|csv|tsv)$/i.test(f.name) || f.type === XLSX_MIME || f.type === 'text/csv';

function sameDocument(a: DocumentInput | undefined, b: DocumentInput | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  const x = documentOf(a);
  const y = documentOf(b);
  return x.source === y.source && x.name === y.name;
}

let fontsWatched = false;

/** Text is measured on a canvas: once the bundled fonts finish loading, measurements change. */
function watchFonts(): void {
  if (fontsWatched || typeof document === 'undefined') return;
  fontsWatched = true;
  void document.fonts?.ready.then(() => {
    clearMeasureCache();
    clearTextCache();
    window.dispatchEvent(new Event('resize'));
  });
}

/** Classes of the interface's root element. */
export function rootClassOf(state: Pick<UIState, 'phone'>): string {
  return `sh-app${state.phone ? ' phone' : ''}`;
}

export class SheetlineUI {
  options: SheetlineOptions;
  /** The page API of this editor (it follows the open workbook). */
  readonly api: SpreadsheetApi;
  readonly files: FileActions;
  readonly prefs: Prefs;
  /**
   * Container on the page's body for menus, dialogs and messages, so that they overlay everything;
   * it carries the theme of the interface. Created by attach().
   */
  readonly layer: HTMLDivElement = document.createElement('div');
  /** The suggestion list and argument tip of formula editing. */
  assist: FormulaAssistState;
  private state: UIState;
  private ctxCache: UIContextValue | null = null;
  private docStateCache: { doc: SpreadsheetDocument; version: number; value: DocState } | null = null;
  private listeners = new Set<() => void>();
  private apiListeners = new Map<keyof ApiEvents, Set<(e: unknown) => void>>();
  private offs: (() => void)[] = [];
  private docOffs: (() => void)[] = [];
  private root: HTMLElement | null = null;
  private gridHost: HTMLElement | null = null;
  private input: HTMLInputElement | null = null;
  private handle: FileSystemFileHandle | null = null;
  private recovery: AutosaveRecord | null = null;
  private memory = new Map<string, unknown>();
  private toastTimer: ReturnType<typeof setTimeout> | null = null;
  private autosaveTimer: ReturnType<typeof setTimeout> | null = null;
  private autosaveOff = false;
  private lastTable: number | null = null;
  private circlesSheet = -1;
  private settling = false;
  private gridWaiters: (() => void)[] = [];
  private started = false;

  constructor(options: SheetlineOptions = {}) {
    this.options = { ...options };
    this.prefs = { get: (name, initial) => this.getPref(name, initial), set: (name, value) => this.setPref(name, value) };
    const phone = matchesMedia(PHONE_QUERY);
    const doc = this.prepare(newDocument());
    this.state = {
      doc,
      docKey: 1,
      docVersion: 0,
      ...PER_DOCUMENT,
      phone,
      layout: options.layout ?? this.getPref<Layout>('layout', 'ribbon'),
      theme: options.theme ?? this.getPref<ThemeName>('theme', 'light'),
      settings: this.getPref<AppSettings>('settings', DEFAULT_SETTINGS),
      collapsed: this.getPref('ribbon-collapsed', false),
      restore: null,
      notice: null,
      busy: null,
      dragging: false,
      accept: ACCEPT,
      ready: false,
    };
    this.assist = this.makeAssist(doc);
    this.files = this.makeFiles();
    this.api = createApi(this);
    this.wireDocument(doc);
  }

  // ---------------------------------------------------------------------------------------------
  // State

  subscribe = (fn: () => void): (() => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };

  getState = (): UIState => this.state;

  /** The context the controls work with; a new object after every change. */
  getCtx = (): UIContextValue => (this.ctxCache ??= this.makeCtx());

  get ctx(): UIContextValue {
    return this.getCtx();
  }

  get doc(): SpreadsheetDocument {
    return this.state.doc;
  }

  /** The interface's root element while attached. */
  get rootElement(): HTMLElement | null {
    return this.root;
  }

  get grid(): GridView | null {
    return this.state.grid;
  }

  set(patch: Partial<UIState>): void {
    const keys = Object.keys(patch) as (keyof UIState)[];
    if (keys.every((k) => Object.is(this.state[k], patch[k]))) return;
    const prev = this.state;
    this.state = { ...prev, ...patch };
    this.ctxCache = null;
    if (prev.theme !== this.state.theme) this.applyTheme();
    if (prev.view !== this.state.view || prev.settings !== this.state.settings) this.applyGridOptions();
    const s = this.state;
    if (s.docVersion !== prev.docVersion || s.doc !== prev.doc || s.activeTab !== prev.activeTab || s.selectedObject !== prev.selectedObject || s.grid !== prev.grid) this.scheduleSettle();
    this.notify();
  }

  private notify(): void {
    for (const fn of [...this.listeners]) fn();
  }

  /** Whether a node belongs to this interface (its root or its layer). */
  contains(node: Node | null): boolean {
    return !!node && (!!this.root?.contains(node) || this.layer.contains(node));
  }

  /** The document state for the current version. */
  docState(): DocState {
    const { doc, docVersion } = this.state;
    const c = this.docStateCache;
    if (c && c.doc === doc && c.version === docVersion) return c.value;
    const value = computeDocState(doc, docVersion);
    this.docStateCache = { doc, version: docVersion, value };
    return value;
  }

  private makeCtx(): UIContextValue {
    const s = this.state;
    const value: UIContextValue = {
      ui: this,
      layer: this.layer,
      prefs: this.prefs,
      doc: s.doc,
      grid: s.grid,
      state: this.docState(),
      edit: s.edit,
      phone: s.phone,
      files: this.files,
      dialog: s.dialog,
      openDialog: this.openDialog,
      closeDialog: this.closeDialog,
      panel: s.panel,
      setPanel: this.setPanel,
      notify: this.notify_,
      confirm: this.confirm,
      ask: this.ask,
      prompt: this.prompt,
      openComments: this.openComments,
      openNote: this.openNote,
      openFilter: this.openFilter,
      selectedObject: s.selectedObject,
      selectObject: this.selectObject,
      backstage: s.backstage,
      openBackstage: this.openBackstage,
      closeBackstage: this.closeBackstage,
      view: s.view,
      setView: this.setView,
      settings: { ...DEFAULT_SETTINGS, ...s.settings },
      setSettings: this.setSettings,
      painter: s.painter,
      setPainter: this.setPainter,
      runCommand: this.runCommand,
      showTab: this.showTab,
    };
    return value;
  }

  // ---------------------------------------------------------------------------------------------
  // Remembered settings

  private key(name: string): string | undefined {
    return this.options.storageKey && `${this.options.storageKey}-${name}`;
  }

  private getPref<T>(name: string, initial: T): T {
    if (this.memory.has(name)) return this.memory.get(name) as T;
    const key = this.key(name);
    if (!key) return initial;
    try {
      const raw = localStorage.getItem(key);
      return raw === null ? initial : (JSON.parse(raw) as T);
    } catch {
      return initial;
    }
  }

  private setPref(name: string, value: unknown): void {
    this.memory.set(name, value);
    const key = this.key(name);
    if (!key) return;
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch {
      // Optional.
    }
  }

  private author(): string {
    if (this.options.author) return this.options.author;
    const key = this.key('author');
    try {
      return (key && localStorage.getItem(key)) || 'Guest';
    } catch {
      return 'Guest';
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Actions

  /** A message at the bottom that goes away by itself. */
  notify_ = (message: string): void => {
    this.set({ toast: message });
    if (this.toastTimer) clearTimeout(this.toastTimer);
    this.toastTimer = setTimeout(() => this.set({ toast: null }), TOAST_MS);
  };

  dismissToast = (): void => this.set({ toast: null });

  /** Show a message that stays until dismissed (null removes it). */
  notice = (message: string | null): void => this.set({ notice: message });

  confirm = (message: string, opts?: ConfirmOptions): Promise<boolean> => new Promise((resolve) => this.set({ pending: { kind: 'confirm', message, opts, resolve } }));

  ask = (message: string, opts?: ConfirmOptions): Promise<'ok' | 'alt' | 'cancel'> => new Promise((resolve) => this.set({ pending: { kind: 'ask', message, opts, resolve } }));

  prompt = (message: string, value = '', opts?: ConfirmOptions): Promise<string | null> => new Promise((resolve) => this.set({ pending: { kind: 'prompt', message, value, opts, resolve } }));

  /** Answer the question shown (the type of `value` follows the question's kind). */
  answer = (value: boolean | 'ok' | 'alt' | 'cancel' | string | null): void => {
    const p = this.state.pending;
    if (!p) return;
    this.set({ pending: null });
    (p.resolve as (v: typeof value) => void)(value);
  };

  /** Client rectangle of a cell (merged cells: the whole merge). */
  cellRect(r: number, c: number): DOMRect {
    const g = this.state.grid;
    if (!g) return new DOMRect(100, 100, 80, 24);
    const vt = g.transform();
    const m = this.state.doc.sheet.mergeAt(r, c);
    const rect = m ? vt.rangeRect(m) : vt.cellRect(r, c);
    const root = g.root.getBoundingClientRect();
    return new DOMRect(root.left + rect.x, root.top + rect.y, rect.w, rect.h);
  }

  private focusGrid(): void {
    this.state.grid?.focusIdle();
  }

  openDialog = (d: DialogSpec): void => {
    const grid = this.state.grid;
    if (grid?.isEditing()) {
      // The function list adds to the formula being typed, at its caret, so that edit stays open
      // behind it (without the focus, so typing cannot change it unseen). Other dialogs end the edit.
      const formula = d.kind === 'insert-function' && grid.editState.text.startsWith('=');
      if (formula) (document.activeElement as HTMLElement | null)?.blur?.();
      else if (d.kind !== 'find') grid.commitEdit();
    }
    this.set({ backstage: null, dialog: d });
  };

  closeDialog = (): void => {
    this.set({ dialog: null });
    const grid = this.state.grid;
    // An edit the dialog left open gets the focus back (and, on a phone, the keyboard).
    if (grid?.isEditing()) {
      if (grid.editState.origin === 'bar') this.root?.querySelector<HTMLElement>('[data-testid="formula-input"]')?.focus({ preventScroll: true });
      else grid.focusEdit();
    } else this.focusGrid();
  };

  setPanel = (panel: SidePanel): void => this.set({ panel });

  openComments = (r: number, c: number, startNew?: boolean): void => this.set({ comments: { r, c, rect: this.cellRect(r, c), startNew } });

  closeComments = (): void => {
    this.set({ comments: null });
    this.focusGrid();
  };

  openNote = (r: number, c: number): void => this.set({ note: { r, c, rect: this.cellRect(r, c) } });

  closeNote = (): void => {
    this.set({ note: null });
    this.focusGrid();
  };

  openFilter = (col: number, at: { x: number; y: number }, tableId?: number): void => this.set({ filter: { col, at, tableId } });

  closeFilter = (): void => {
    this.set({ filter: null });
    this.focusGrid();
  };

  openMenu = (menu: OpenMenu): void => this.set({ menu });

  closeMenu = (): void => {
    this.set({ menu: null });
    this.focusGrid();
  };

  /** The menu of a sheet tab. */
  openTabMenu = (index: number, at: MenuAnchor): void => this.openMenu({ items: tabMenu(this.ctx, index), at, testId: 'tab-menu' });

  selectObject = (id: string | null): void => {
    this.state.grid?.selectImage(id);
    this.set({ selectedObject: id });
  };

  openBackstage = (page = 'new'): void => {
    const grid = this.state.grid;
    if (grid?.isEditing()) grid.commitEdit();
    this.set({ backstage: page });
  };

  /** Show another page of the open backstage. */
  setBackstagePage = (page: string | null): void => this.set({ backstage: page });

  closeBackstage = (): void => {
    this.set({ backstage: null });
    this.focusGrid();
  };

  setView = (patch: Partial<ViewSettings>): void => this.set({ view: { ...this.state.view, ...patch } });

  setSettings = (patch: Partial<AppSettings>): void => {
    const settings = { ...DEFAULT_SETTINGS, ...this.state.settings, ...patch };
    this.set({ settings });
    this.setPref('settings', settings);
  };

  setPainter = (mode: 'once' | 'sticky' | null): void => {
    this.set({ painter: mode });
    if (!mode) {
      this.state.doc.clearClip();
      this.state.grid?.markClip();
    }
  };

  /** Runs a ribbon command by id (command palette, shortcuts). */
  runCommand = (id: string): void => {
    const ctx = this.ctx;
    const c = allCommands(ctx).find((x) => x.id === id);
    if (c?.run) c.run(ctx);
  };

  showTab = (id: string): void => this.set({ activeTab: id });

  /** The tab shown: the active one while it is visible, else Home. */
  currentTab(): string {
    const ctx = this.ctx;
    const visible = RIBBON_TABS.filter((t) => !t.visible || t.visible(ctx)).map((t) => t.id);
    return visible.includes(this.state.activeTab) ? this.state.activeTab : 'home';
  }

  setLayout = (layout: Layout): void => {
    if (layout === this.state.layout) return;
    this.set({ layout });
    this.setPref('layout', layout);
    this.options.onLayoutChange?.(layout);
  };

  setTheme = (theme: ThemeName): void => {
    if (theme === this.state.theme) return;
    this.set({ theme });
    this.setPref('theme', theme);
    this.options.onThemeChange?.(theme);
  };

  setCollapsed = (collapsed: boolean): void => {
    this.set({ collapsed });
    this.setPref('ribbon-collapsed', collapsed);
  };

  /** Light / dark theme of the interface (cells keep their own colours). */
  private applyTheme(): void {
    const dark = this.state.theme === 'dark';
    this.layer.dataset.theme = this.state.theme;
    const grid = this.state.grid;
    if (grid) {
      grid.dark = dark;
      grid.refreshObjects();
      grid.requestRender();
    }
  }

  /** View options the grid draws. */
  private applyGridOptions(): void {
    const grid = this.state.grid;
    if (!grid) return;
    const { view, settings } = this.state;
    grid.showFormulas = view.showFormulas;
    grid.focusCell = view.focusCell;
    grid.allowFill = settings.allowFill;
    grid.allowZoom = settings.allowZoom;
    grid.requestRender();
  }

  /** Follow-ups of changes, once per batch: the contextual tabs and the invalid-data circles. */
  private scheduleSettle(): void {
    if (this.settling) return;
    this.settling = true;
    queueMicrotask(() => {
      this.settling = false;
      this.settle();
    });
  }

  private settle(): void {
    const s = this.state;
    const docState = this.docState();
    // Invalid-data circles belong to the sheet they were drawn on.
    if (s.grid && docState.activeIndex !== this.circlesSheet) {
      if (this.circlesSheet >= 0 && s.grid.circles) {
        s.grid.circles = null;
        s.grid.requestRender();
      }
      this.circlesSheet = docState.activeIndex;
    }
    // A new table brings its tab forward; a deselected object sends its tab away.
    let tab = s.activeTab;
    if (docState.tableId !== null && this.lastTable === null && tab === 'insert') tab = 'table-design';
    this.lastTable = docState.tableId;
    if (!s.selectedObject && ['chart-design', 'shape-format', 'picture-format'].includes(tab)) tab = 'home';
    if (tab !== s.activeTab) this.set({ activeTab: tab });
  }

  // ---------------------------------------------------------------------------------------------
  // Documents

  /** Settings of a workbook about to be shown: view-only and the author. */
  private prepare(doc: SpreadsheetDocument): SpreadsheetDocument {
    if (this.options.readOnly) doc.mode = 'viewOnly';
    doc.author = this.author();
    return doc;
  }

  private makeAssist(doc: SpreadsheetDocument): FormulaAssistState {
    const names = doc.wb.names.filter((n) => !n.hidden).map((n) => n.name);
    return new FormulaAssistState(names, (text, caret) => this.state.grid?.replaceEditText(text, caret), () => {
      this.ctxCache = null;
      this.notify();
    });
  }

  /** Show a workbook in place of the current one. */
  private load(next: SpreadsheetDocument, handle: FileSystemFileHandle | null = null): void {
    this.prepare(next);
    this.handle = handle;
    this.recovery = null;
    const prev = this.state.grid;
    this.unwireDocument();
    this.assist = this.makeAssist(next);
    this.lastTable = null;
    this.circlesSheet = -1;
    this.set({ ...PER_DOCUMENT, doc: next, docKey: this.state.docKey + 1, docVersion: 0, restore: null });
    // The old grid goes with its workbook; the view attaches a new one for this workbook.
    if (prev) {
      prev.destroy();
      this.gridHost = null;
    }
    this.wireDocument(next);
    this.emit('open', { fileName: next.fileName });
  }

  private wireDocument(doc: SpreadsheetDocument): void {
    const bump = () => this.set({ docVersion: this.state.docVersion + 1 });
    this.docOffs.push(...DOC_STATE_EVENTS.map((e) => doc.on(e as 'change', bump)));
    for (const name of ['change', 'selection', 'activeSheet', 'sheets', 'file'] as const) {
      this.docOffs.push(doc.on(name as 'change', (e: unknown) => this.emit(name, e)));
    }
    this.autosaveOff = false;
    this.docOffs.push(doc.on('change', () => this.scheduleAutosave(doc)));
  }

  private unwireDocument(): void {
    for (const off of this.docOffs.splice(0)) off();
    if (this.autosaveTimer) clearTimeout(this.autosaveTimer);
    this.autosaveTimer = null;
  }

  private scheduleAutosave(doc: SpreadsheetDocument): void {
    const db = this.options.autosave;
    if (!doc.modified || !db || this.autosaveOff) return;
    if (this.autosaveTimer) clearTimeout(this.autosaveTimer);
    this.autosaveTimer = setTimeout(() => {
      if (cellCount(doc) > AUTOSAVE_MAX_CELLS) {
        this.autosaveOff = true;
        return;
      }
      try {
        const started = performance.now();
        const bytes = doc.save(false);
        if (performance.now() - started > 1500) this.autosaveOff = true;
        void saveAutosave(db, { fileName: doc.fileName, bytes, savedAt: Date.now() });
      } catch (e) {
        console.warn('Autosave failed', e);
      }
    }, AUTOSAVE_DELAY);
  }

  private clearAutosave(): void {
    if (this.options.autosave) void clearAutosave(this.options.autosave);
  }

  /** Opens a workbook from bytes (XLSX, XLSM, XLTX, XLS, XLSB, CSV); resolves to whether it could be read. */
  openBytes(bytes: Uint8Array, name: string, handle: FileSystemFileHandle | null = null): Promise<boolean> {
    return new Promise<boolean>((resolve) => {
      this.set({ busy: `Opening “${name}”…` });
      // Importing is synchronous: let the message paint first.
      requestAnimationFrame(() =>
        setTimeout(() => {
          try {
            this.load(openDocument(bytes, name), handle);
            this.clearAutosave();
            resolve(true);
          } catch (e) {
            console.error(e);
            this.notice(`Could not open “${name}”: ${(e as Error)?.message ?? e}`);
            resolve(false);
          } finally {
            this.set({ busy: null });
          }
        }, 0),
      );
    });
  }

  openFile(file: File, handle: FileSystemFileHandle | null = null): Promise<boolean> {
    return file.arrayBuffer().then((buf) => this.openBytes(new Uint8Array(buf), file.name, handle));
  }

  async openUrl(url: string, name: string): Promise<boolean> {
    try {
      const r = await fetch(url);
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      return await this.openBytes(new Uint8Array(await r.arrayBuffer()), name);
    } catch (e) {
      this.notice(`Could not load ${url}: ${(e as Error).message}`);
      return false;
    }
  }

  /** Open a workbook: a URL, a File or Blob, or the bytes; the file picker when no source is given. */
  open = async (source?: DocumentSource, name?: string): Promise<boolean> => {
    if (source === undefined) {
      this.files.open();
      return false;
    }
    if (typeof source === 'string') return this.openUrl(source, name ?? nameOfUrl(source));
    let bytes: Uint8Array;
    try {
      bytes = source instanceof Uint8Array ? source : source instanceof ArrayBuffer ? new Uint8Array(source) : new Uint8Array(await source.arrayBuffer());
    } catch (e) {
      this.notice(`Could not open the file: ${(e as Error).message}`);
      return false;
    }
    return this.openBytes(bytes, name ?? (typeof File !== 'undefined' && source instanceof File ? source.name : 'Book1.xlsx'));
  };

  private async openDocument(input: DocumentInput): Promise<boolean> {
    const { source, name } = documentOf(input);
    return source === undefined ? false : this.open(source, name);
  }

  /** Recover the autosaved workbook offered. */
  recover = (): void => {
    const r = this.recovery;
    if (!r) return;
    try {
      const recovered = openDocument(r.bytes, r.fileName);
      recovered.modified = true;
      this.load(recovered);
    } catch (e) {
      this.notice(`Could not recover “${r.fileName}”: ${(e as Error).message}`);
    }
  };

  discardRecovery = (): void => {
    this.recovery = null;
    this.set({ restore: null });
    this.clearAutosave();
  };

  private async offerRecovery(): Promise<void> {
    const db = this.options.autosave;
    if (!db || this.options.offerRecovery === false) return;
    const doc = this.state.doc;
    const rec = await loadAutosave(db);
    if (this.state.doc !== doc || !rec || Date.now() - rec.savedAt >= RECOVERY_MS) return;
    this.recovery = rec;
    this.set({ restore: { fileName: rec.fileName, savedAt: rec.savedAt } });
  }

  /** Ask before replacing a workbook with unsaved changes. */
  confirmDiscard = (): boolean => !this.state.doc.modified || window.confirm(DISCARD_PROMPT);

  // ---------------------------------------------------------------------------------------------
  // Files

  private makeFiles(): FileActions {
    const pick = <K extends Exclude<keyof FileActions, 'share'>>(k: K): FileActions[K] => this.options.files?.[k] ?? this.builtinFiles[k];
    const files: FileActions = {
      newWorkbook: () => pick('newWorkbook')(),
      open: () => pick('open')(),
      openData: (bytes, name) => pick('openData')(bytes, name),
      save: () => pick('save')(),
      exportAs: (format, name) => pick('exportAs')(format, name),
      importFile: (kind) => pick('importFile')(kind),
      print: () => pick('print')(),
      rename: (name) => pick('rename')(name),
    };
    // Only where the browser can share files (or the share action was replaced).
    Object.defineProperty(files, 'share', { enumerable: true, get: () => this.options.files?.share ?? this.builtinFiles.share });
    return files;
  }

  private commitEdit(): void {
    const grid = this.state.grid;
    if (grid?.isEditing()) grid.commitEdit();
  }

  private excelBlob(d: SpreadsheetDocument, format: Exclude<ExportFormat, 'csv' | 'pdf'>, mark: boolean): Blob {
    const macro = /\.xl[st]m$/i.test(d.fileName);
    return new Blob([d.save(mark, { macros: format === 'xlsm' || format === 'xltm' ? true : macro ? false : undefined, template: format === 'xltx' || format === 'xltm' }) as BlobPart], { type: MIME[format] });
  }

  /** Opens the hidden file input with these file types. */
  private pickFile(types: string): void {
    this.set({ accept: types });
    if (this.input) this.input.accept = types;
    // The accept attribute must be updated before the click.
    requestAnimationFrame(() => this.input?.click());
  }

  private readonly builtinFiles: FileActions = (() => {
    const canShare = typeof navigator !== 'undefined' && !!navigator.canShare && (() => {
      try {
        return navigator.canShare({ files: [new File([new Uint8Array(1)], 'x.xlsx', { type: XLSX_MIME })] });
      } catch {
        return false;
      }
    })();
    const files: FileActions = {
      newWorkbook: () => {
        this.load(newDocument());
        this.clearAutosave();
      },
      open: async () => {
        const w = window as unknown as FilePickerWindow;
        if (w.showOpenFilePicker && window.isSecureContext) {
          try {
            const [handle] = await w.showOpenFilePicker({
              types: [{ description: 'Spreadsheets', accept: { [XLSX_MIME]: ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xlsb', '.xls'], 'text/csv': ['.csv', '.tsv', '.txt'] } }],
              multiple: false,
            });
            void this.openFile(await handle.getFile(), handle);
          } catch (e) {
            if ((e as DOMException)?.name !== 'AbortError') this.notice(`Could not open the file: ${(e as Error).message}`);
          }
          return;
        }
        this.pickFile(ACCEPT);
      },
      openData: (bytes, name) => this.openBytes(bytes, name),
      importFile: (kind) => this.pickFile(kind === 'csv' ? CSV_ACCEPT : EXCEL_ACCEPT),
      save: async () => {
        const d = this.state.doc;
        this.commitEdit();
        const handle = this.handle;
        if (handle && /\.(xlsx|xlsm)$/i.test(handle.name)) {
          try {
            const writable = await (handle as unknown as { createWritable(): Promise<FileSystemWritableFileStream> }).createWritable();
            await writable.write(this.excelBlob(d, /\.xlsm$/i.test(handle.name) ? 'xlsm' : 'xlsx', true));
            await writable.close();
            this.clearAutosave();
            return;
          } catch (e) {
            if ((e as DOMException)?.name === 'AbortError') return;
            console.warn('Saving to the file failed, downloading instead', e);
          }
        }
        const ext = (/\.(xlsx|xlsm|xltx|xltm)$/i.exec(d.fileName)?.[1].toLowerCase() ?? 'xlsx') as Exclude<ExportFormat, 'csv' | 'pdf'>;
        downloadBlob(this.excelBlob(d, ext, true), `${baseName(d)}.${ext}`);
        this.clearAutosave();
      },
      exportAs: (format, name) => {
        const d = this.state.doc;
        this.commitEdit();
        if (format === 'pdf') {
          this.set({ backstage: 'print' });
          return;
        }
        const base = (name ?? baseName(d)).replace(/\.(xlsx|xlsm|xltx|xltm|csv)$/i, '') || 'Book1';
        if (format === 'csv') {
          downloadBlob(new Blob([`﻿${sheetToCsv(d, d.sheet, ',')}`], { type: 'text/csv' }), `${base}.csv`);
          return;
        }
        // A copy: the workbook keeps its unsaved mark unless it is saved under its own name.
        const own = `${base}.${format}`.toLowerCase() === d.fileName.toLowerCase();
        downloadBlob(this.excelBlob(d, format, own), `${base}.${format}`);
        if (own) this.clearAutosave();
      },
      print: () => {
        this.commitEdit();
        this.set({ backstage: 'print' });
      },
      rename: (name) => {
        const d = this.state.doc;
        const ext = /\.(xlsx|xlsm|xltx|xltm|csv|tsv)$/i.exec(d.fileName)?.[0] ?? '.xlsx';
        d.setFileName(`${name.replace(/\.(xlsx|xlsm|xltx|xltm|xls|xlsb|csv|tsv)$/i, '')}${ext}`);
        this.handle = null;
      },
      share: canShare && typeof window !== 'undefined' && window.isSecureContext
        ? async () => {
            const d = this.state.doc;
            this.commitEdit();
            try {
              const format = /\.xl[st]m$/i.test(d.fileName) ? 'xlsm' : 'xlsx';
              const name = `${baseName(d)}.${format}`;
              await navigator.share({ files: [new File([this.excelBlob(d, format, false)], name, { type: MIME[format] })], title: name });
            } catch (e) {
              if ((e as DOMException)?.name !== 'AbortError') this.notice(`Sharing failed: ${(e as Error).message}`);
            }
          }
        : undefined,
    };
    return files;
  })();

  // ---------------------------------------------------------------------------------------------
  // Page API events

  /** Listen to events of the open workbook (they follow the workbooks opened later). */
  on<K extends keyof ApiEvents>(event: K, fn: (e: ApiEvents[K]) => void): () => void {
    let set = this.apiListeners.get(event);
    if (!set) this.apiListeners.set(event, (set = new Set()));
    const f = fn as (e: unknown) => void;
    set.add(f);
    return () => set.delete(f);
  }

  private emit(event: keyof ApiEvents, e: unknown): void {
    for (const fn of [...(this.apiListeners.get(event) ?? [])]) fn(e);
  }

  /** Resolves once the grid of the open workbook is attached. */
  whenGrid(): Promise<void> {
    if (this.state.grid) return Promise.resolve();
    return new Promise((resolve) => this.gridWaiters.push(resolve));
  }

  // ---------------------------------------------------------------------------------------------
  // Options

  /** New options from the component's props; the document, layout, theme and mode follow them. */
  setOptions(next: SheetlineOptions): void {
    const prev = this.options;
    this.options = { ...next };
    if (next.layout && next.layout !== prev.layout) this.set({ layout: next.layout });
    if (next.theme && next.theme !== prev.theme) this.set({ theme: next.theme });
    if (next.author !== prev.author) this.state.doc.author = this.author();
    if (!!next.readOnly !== !!prev.readOnly) this.state.doc.setMode(next.readOnly ? 'viewOnly' : 'editing');
    if (this.started && next.document && !sameDocument(next.document, prev.document)) void this.openDocument(next.document);
  }

  // ---------------------------------------------------------------------------------------------
  // Lifetime

  /** Start the interface in its root element: the layer, listeners, and the initial document. */
  attach(root: HTMLElement): void {
    if (this.root) this.detach();
    this.root = root;
    watchFonts();
    const layer = this.layer;
    layer.className = 'sh-layer';
    layer.dataset.theme = this.state.theme;
    document.body.appendChild(layer);
    this.options.onThemeChange?.(this.state.theme);
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = this.state.accept;
    input.hidden = true;
    input.dataset.testid = 'open-input';
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      input.value = '';
      if (file) void this.openFile(file);
    });
    layer.appendChild(input);
    this.input = input;
    this.ctxCache = null;

    // Shortcuts outside the grid (ribbon focused, dialogs closed).
    const keys = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement | null;
      if (!this.options.globalShortcuts && !this.contains(t)) return;
      if (t?.closest?.('input, textarea, [contenteditable], .sh-modal-layer')) return;
      if (this.state.grid?.root.contains(t)) return;
      if (handleShortcut(this.ctx, e)) e.preventDefault();
    };
    window.addEventListener('keydown', keys);
    this.offs.push(() => window.removeEventListener('keydown', keys));

    // Spreadsheets dropped on the editor open; images dropped on the grid are inserted.
    const hasFiles = (e: DragEvent) => Array.from(e.dataTransfer?.types ?? []).includes('Files');
    let depth = 0;
    const enter = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      depth++;
      this.set({ dragging: true });
    };
    const leave = () => {
      depth = Math.max(0, depth - 1);
      if (!depth) this.set({ dragging: false });
    };
    const over = (e: DragEvent) => {
      if (hasFiles(e)) e.preventDefault();
    };
    const drop = (e: DragEvent) => {
      depth = 0;
      this.set({ dragging: false });
      const all = Array.from(e.dataTransfer?.files ?? []);
      const file = all.find(isSheetFile);
      if (!file) {
        const img = all.find((f) => f.type.startsWith('image/'));
        if (img && !this.state.doc.isReadOnly()) {
          e.preventDefault();
          void this.state.grid?.insertImageFile(img, img.name);
        }
        return;
      }
      e.preventDefault();
      if (!this.confirmDiscard()) return;
      void this.openFile(file);
    };
    for (const el of [root, layer]) {
      el.addEventListener('dragenter', enter);
      el.addEventListener('dragleave', leave);
      el.addEventListener('dragover', over);
      el.addEventListener('drop', drop);
      this.offs.push(() => {
        el.removeEventListener('dragenter', enter);
        el.removeEventListener('dragleave', leave);
        el.removeEventListener('dragover', over);
        el.removeEventListener('drop', drop);
      });
    }

    // Volatile functions (NOW, TODAY, RAND) refresh every minute while the page is visible.
    const volatile = setInterval(() => {
      if (document.visibilityState === 'visible' && !this.state.grid?.isEditing()) this.state.doc.recalcVolatile();
    }, 60_000);
    this.offs.push(() => clearInterval(volatile), watchMedia(PHONE_QUERY, (phone) => this.set({ phone })));
    this.set({ phone: matchesMedia(PHONE_QUERY) });
    this.notify();
    if (!this.started) {
      this.started = true;
      void this.start();
    }
  }

  private async start(): Promise<void> {
    // Let a view that is mounted twice in a row (React's StrictMode) settle first.
    await Promise.resolve();
    const o = this.options;
    if (o.document) await this.openDocument(o.document);
    else void this.offerRecovery();
    await this.whenGrid();
    this.set({ ready: true });
    this.options.onReady?.(this.api);
  }

  /** The grid of the open workbook in its element (inside the root). */
  attachGrid(host: HTMLElement): void {
    if (this.gridHost === host && this.state.grid) return;
    this.state.grid?.destroy();
    const doc = this.state.doc;
    const g = new GridView(host, doc, {
      onEdit: (e) => {
        this.assist.update(e);
        this.set({ edit: e });
      },
      editorKey: (e) => this.assist.handleKey(e),
      onContextMenu: (x, y, target) => {
        const ctx = this.ctx;
        if (target === 'image') {
          const id = ctx.grid?.selectedImageId;
          if (id) this.openMenu({ items: objectMenu(ctx, id), at: { x, y }, testId: 'object-menu' });
          return;
        }
        const items = target === 'col' ? columnMenu(ctx) : target === 'row' ? rowMenu(ctx) : target === 'corner' ? cornerMenu(ctx) : cellMenu(ctx);
        this.openMenu({ items, at: { x, y }, title: ctx.state.rangeRef, testId: `${target === 'col' ? 'column' : target}-menu` });
      },
      onCommand: (c) => {
        const ctx = this.ctx;
        switch (c) {
          case 'find':
            this.set({ dialog: { kind: 'find', tab: 'find' } });
            break;
          case 'replace':
            this.set({ dialog: { kind: 'find', tab: 'replace' } });
            break;
          case 'link':
            if (!doc.isReadOnly()) this.set({ dialog: { kind: 'hyperlink' } });
            break;
          case 'comment':
            cmd.editNote(ctx);
            break;
          case 'save':
            ctx.files.save();
            break;
          case 'open':
            ctx.files.open();
            break;
          case 'print':
            this.set({ backstage: 'print' });
            break;
          case 'clear-format':
            if (!doc.isReadOnly()) cmd.clear(ctx, 'formats');
            break;
          case 'next-sheet':
          case 'prev-sheet': {
            const visible = doc.wb.sheets.map((s, i) => (s.state === 'visible' ? i : -1)).filter((i) => i >= 0);
            const pos = visible.indexOf(doc.activeIndex);
            doc.setActiveSheet(visible[(pos + (c === 'next-sheet' ? 1 : -1) + visible.length) % visible.length]);
            break;
          }
          default:
            break;
        }
      },
      onShortcut: (e) => handleShortcut(this.ctx, e),
      onSelectEnd: () => {
        const painter = this.state.painter;
        if (!painter) return;
        doc.pasteSpecial({ what: 'formats' });
        if (painter === 'once') {
          this.set({ painter: null });
          doc.clearClip();
          g.markClip();
        }
      },
      onImageSelect: (id) => {
        if (id) {
          const hit = doc.objectById(id);
          this.set({ selectedObject: id, activeTab: hit?.kind === 'chart' ? 'chart-design' : hit?.kind === 'shape' ? 'shape-format' : 'picture-format' });
        } else this.set({ selectedObject: null });
      },
      onObjectEdit: (id) => {
        if (doc.objectById(id)) this.set({ selectedObject: id, panel: 'format' });
      },
      onLinkTap: (link) => {
        if (link.location) {
          doc.goTo(link.location.replace(/^#/, ''));
          return;
        }
        if (link.url) {
          const w = window.open(link.url, '_blank', 'noopener,noreferrer');
          if (!w) this.notify_(`Open ${link.url}`);
        }
      },
      onValidationList: (options, rect, pick) => {
        this.openMenu({ items: options.map((o) => ({ title: o, onClick: () => pick(o) })), at: rect, testId: 'validation-list' });
      },
      validateEdit: (r, c, text, sheet, decide) => {
        const rule = doc.invalidInput(r, c, text, sheet);
        if (!rule || rule.formatStyle || rule.showErrorMessage === false) return false;
        const message = rule.error || 'This value doesn\'t match the data validation restrictions defined for this cell.';
        if (rule.errorStyle === 'stop') {
          void this.confirm(message, { title: rule.errorTitle || 'Microsoft Excel', ok: 'Retry', cancel: 'Cancel' }).then((retry) => {
            if (retry) decide(false);
            else g.cancelEdit();
          });
        } else {
          void this.confirm(`${message}\n\nContinue?`, { title: rule.errorTitle || (rule.errorStyle === 'warning' ? 'Warning' : 'Information'), ok: 'Yes', cancel: 'No' }).then((keep) => decide(keep));
        }
        return true;
      },
      canEdit: (r, c) => {
        if (!doc.sheet.protection || !doc.isCellLocked(r, c)) return true;
        this.notify_(PROTECTED_MESSAGE);
        return false;
      },
      onFilterButton: (col, x, y, tableId) => this.set({ filter: { col, at: { x, y }, tableId } }),
    }, { mobile: isTouchDevice() });
    g.dark = this.state.theme === 'dark';
    this.gridHost = host;
    this.set({ grid: g });
    this.applyGridOptions();
    if (!isTouchDevice()) g.focusIdle();
    for (const resolve of this.gridWaiters.splice(0)) resolve();
  }

  /** The grid's element goes away. */
  detachGrid(host: HTMLElement): void {
    if (this.gridHost !== host) return;
    this.state.grid?.destroy();
    this.gridHost = null;
    this.set({ grid: null });
  }

  /** Remove the interface and stop listening; attach() may be called again. */
  detach(): void {
    for (const off of this.offs.splice(0)) off();
    if (this.toastTimer) clearTimeout(this.toastTimer);
    this.toastTimer = null;
    this.input?.remove();
    this.input = null;
    this.layer.remove();
    this.root = null;
  }

  /** Remove the interface for good (the component unmounted): also the grid and the document's listeners. */
  destroy(): void {
    this.detach();
    if (this.gridHost) this.detachGrid(this.gridHost);
    this.unwireDocument();
  }
}
