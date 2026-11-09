/** Formulas tab: Function Library, Defined Names, Formula Auditing, Calculation. */
import type { MenuItem } from '../menuItems';
import { functionCatalog, signature } from '../functionCatalog';
import { TINT, type Command, type Ctx, type Tab } from './types';
import { C } from './homeTab';

/** Excel's names of the function categories. */
export const CATEGORY_LABELS: Record<string, string> = {
  'Math & trigonometry': 'Math & Trig',
  Statistical: 'Statistical',
  Text: 'Text',
  Logical: 'Logical',
  'Lookup & reference': 'Lookup & Reference',
  'Date & time': 'Date & Time',
  Financial: 'Financial',
  Engineering: 'Engineering',
  Information: 'Information',
  Database: 'Database',
  Web: 'Web',
  Compatibility: 'Compatibility',
};

/** Starts (or continues) a formula in the active cell with a function call. */
export function insertFunction(ctx: Ctx, name: string): void {
  const g = ctx.grid;
  if (!g || ctx.state.readOnly) return;
  if (g.isEditing() && g.editState.text.startsWith('=')) {
    const { text, caret } = g.editState;
    const insert = `${name}(`;
    g.replaceEditText(text.slice(0, caret) + insert + text.slice(caret), caret + insert.length);
    return;
  }
  g.startEdit(`=${name}(`, 'enter');
}

function categoryMenu(category: string): (ctx: Ctx) => MenuItem[] {
  return (ctx) => {
    const list = [...functionCatalog().values()].filter((f) => f.category === category);
    return [
      ...list.map((f) => ({ title: f.name, tip: signature(f), onClick: () => insertFunction(ctx, f.name) })),
      { type: 'line' as const },
      { title: 'Insert Function...', icon: 'functions' as const, onClick: () => ctx.openDialog({ kind: 'insert-function', category }) },
    ];
  };
}

function namesMenu(ctx: Ctx): MenuItem[] {
  const names = ctx.doc.wb.names.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.') && (n.localSheet === undefined || n.localSheet === ctx.doc.activeIndex));
  const tables = ctx.doc.wb.sheets.flatMap((s) => s.tables.map((t) => t.name));
  const items: MenuItem[] = [...names.map((n) => n.name), ...tables].sort().map((name) => ({
    title: name,
    onClick: () => {
      const g = ctx.grid;
      if (!g) return;
      if (g.isEditing() && g.editState.text.startsWith('=')) {
        const { text, caret } = g.editState;
        g.replaceEditText(text.slice(0, caret) + name + text.slice(caret), caret + name.length);
      } else g.startEdit(`=${name}`, 'enter');
    },
  }));
  if (!items.length) items.push({ title: '(No names defined)', disabled: true });
  return items;
}

const calcMode = (ctx: Ctx) => ctx.doc.wb.calcMode ?? 'auto';

export const F = {
  insertFunction: { id: 'insert-function', label: 'Insert Function', icon: 'functions', tint: TINT.blue, shortcut: 'Shift+F3', keywords: 'fx formula', run: (ctx) => ctx.openDialog({ kind: 'insert-function' }) } as Command,
  financial: { id: 'functions-financial', label: 'Financial', icon: 'attach_money', tint: TINT.green, menu: categoryMenu('Financial') } as Command,
  logical: { id: 'functions-logical', label: 'Logical', icon: 'help_outline', tint: TINT.blue, menu: categoryMenu('Logical') } as Command,
  text: { id: 'functions-text', label: 'Text', icon: 'text_fields', tint: TINT.blue, menu: categoryMenu('Text') } as Command,
  date: { id: 'functions-date', label: 'Date & Time', icon: 'schedule', tint: TINT.blue, menu: categoryMenu('Date & time') } as Command,
  lookup: { id: 'functions-lookup', label: 'Lookup & Reference', icon: 'manage_search', tint: TINT.blue, menu: categoryMenu('Lookup & reference') } as Command,
  math: { id: 'functions-math', label: 'Math & Trig', icon: 'calculate', tint: TINT.blue, menu: categoryMenu('Math & trigonometry') } as Command,
  more: {
    id: 'functions-more',
    label: 'More Functions',
    icon: 'more_horiz',
    menu: (ctx) =>
      ['Statistical', 'Engineering', 'Information', 'Database', 'Web', 'Compatibility']
        .filter((cat) => [...functionCatalog().values()].some((f) => f.category === cat))
        .map((cat) => ({ title: CATEGORY_LABELS[cat] ?? cat, submenu: categoryMenu(cat)(ctx) })),
  } as Command,
  nameManager: { id: 'name-manager', label: 'Name Manager', icon: 'label', tint: TINT.orange, shortcut: 'Ctrl+F3', keywords: 'defined names range names', run: (ctx) => ctx.openDialog({ kind: 'name-manager' }) } as Command,
  defineName: { id: 'define-name', label: 'Define Name', icon: 'drive_file_rename_outline', run: (ctx) => ctx.openDialog({ kind: 'new-name' }), disabled: (ctx) => ctx.state.readOnly } as Command,
  useInFormula: { id: 'use-in-formula', label: 'Use in Formula', icon: 'post_add', menu: namesMenu } as Command,
  fromSelection: { id: 'names-from-selection', label: 'Create from Selection', icon: 'select_all', shortcut: 'Ctrl+Shift+F3', run: (ctx) => ctx.openDialog({ kind: 'names-from-selection' }), disabled: (ctx) => ctx.state.readOnly } as Command,
  showFormulas: { id: 'show-formulas', label: 'Show Formulas', icon: 'functions', shortcut: 'Ctrl+`', checked: (ctx) => ctx.view.showFormulas, run: (ctx) => ctx.setView({ showFormulas: !ctx.view.showFormulas }) } as Command,
  formulaBar: { id: 'show-formula-bar', label: 'Formula Bar', icon: 'short_text', checked: (ctx) => ctx.view.formulaBar, run: (ctx) => ctx.setView({ formulaBar: !ctx.view.formulaBar }) } as Command,
  calcOptions: {
    id: 'calculation-options',
    label: 'Calculation Options',
    icon: 'calculate',
    menu: (ctx) => [
      { title: 'Automatic', checked: calcMode(ctx) === 'auto', testId: 'calc-auto', onClick: () => ctx.doc.setCalcMode('auto') },
      { title: 'Automatic Except for Data Tables', checked: calcMode(ctx) === 'autoNoTable', onClick: () => ctx.doc.setCalcMode('autoNoTable') },
      { title: 'Manual', checked: calcMode(ctx) === 'manual', testId: 'calc-manual', onClick: () => ctx.doc.setCalcMode('manual') },
    ],
  } as Command,
  calcNow: { id: 'calculate-now', label: 'Calculate Now', icon: 'refresh', shortcut: 'F9', run: (ctx) => ctx.doc.calculateNow() } as Command,
  calcSheet: { id: 'calculate-sheet', label: 'Calculate Sheet', icon: 'replay', shortcut: 'Shift+F9', run: (ctx) => ctx.doc.calculateSheet() } as Command,
};

export const formulasTab: Tab = {
  id: 'formulas',
  label: 'Formulas',
  groups: () => [
    {
      id: 'function-library',
      label: 'Function Library',
      icon: 'functions',
      controls: [
        { type: 'large', cmd: F.insertFunction },
        { type: 'large', cmd: C.autosum, split: true },
        { type: 'stack', items: [{ type: 'small', cmd: F.financial }, { type: 'small', cmd: F.logical }, { type: 'small', cmd: F.text }] },
        { type: 'stack', items: [{ type: 'small', cmd: F.date }, { type: 'small', cmd: F.lookup }, { type: 'small', cmd: F.math }] },
        { type: 'stack', items: [{ type: 'small', cmd: F.more }] },
      ],
    },
    {
      id: 'defined-names',
      label: 'Defined Names',
      icon: 'label',
      controls: [
        { type: 'large', cmd: F.nameManager },
        { type: 'stack', items: [{ type: 'small', cmd: F.defineName }, { type: 'small', cmd: F.useInFormula }, { type: 'small', cmd: F.fromSelection }] },
      ],
    },
    {
      id: 'formula-auditing',
      label: 'Formula Auditing',
      icon: 'fact_check',
      controls: [{ type: 'stack', items: [{ type: 'small', cmd: F.showFormulas }, { type: 'small', cmd: F.formulaBar }] }],
    },
    {
      id: 'calculation',
      label: 'Calculation',
      icon: 'calculate',
      controls: [{ type: 'large', cmd: F.calcOptions }, { type: 'stack', items: [{ type: 'small', cmd: F.calcNow }, { type: 'small', cmd: F.calcSheet }] }],
    },
  ],
  toolbar: () => [
    { type: 'small', cmd: F.insertFunction },
    { type: 'small', cmd: C.autosum, split: true },
    { type: 'small', cmd: F.financial },
    { type: 'small', cmd: F.logical },
    { type: 'small', cmd: F.text },
    { type: 'small', cmd: F.date },
    { type: 'small', cmd: F.lookup },
    { type: 'small', cmd: F.math },
    { type: 'small', cmd: F.more },
    { type: 'small', cmd: F.nameManager },
    { type: 'small', cmd: F.defineName },
    { type: 'small', cmd: F.showFormulas },
    { type: 'small', cmd: F.calcOptions },
    { type: 'small', cmd: F.calcNow },
  ],
};
