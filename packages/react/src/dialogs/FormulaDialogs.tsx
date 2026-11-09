/** Formula dialogs: Insert Function, Name Manager, New / Edit Name, Create Names from Selection. */
import { useMemo, useState } from 'react';
import { useUI } from '../context';
import { Check, Dialog, Field, ListBox, RangeField, rangeText } from './Dialog';
import { CATEGORY_LABELS, CellError, type DefinedName, functionCatalog, type FunctionInfo, insertFunction, makeRange, signature } from '@sheetline/core/views';

// -----------------------------------------------------------------------------------------------
// Insert Function

export function InsertFunctionDialog({ category, onClose }: { category?: string; onClose: () => void }) {
  const ui = useUI();
  const all = useMemo(() => [...functionCatalog().values()], []);
  const categories = useMemo(() => [...new Set(all.map((f) => f.category))].sort((a, b) => (CATEGORY_LABELS[a] ?? a).localeCompare(CATEGORY_LABELS[b] ?? b)), [all]);
  const [cat, setCat] = useState(category ?? 'All');
  const [query, setQuery] = useState('');
  const list = useMemo(() => {
    const q = query.trim().toUpperCase();
    return all.filter((f) => (cat === 'All' || f.category === cat) && (!q || f.name.includes(q) || f.description.toUpperCase().includes(q)));
  }, [all, cat, query]);
  const [sel, setSel] = useState<FunctionInfo | null>(list[0] ?? null);
  const current = sel && list.includes(sel) ? sel : list[0] ?? null;
  return (
    <Dialog
      title="Insert Function"
      onClose={onClose}
      width={480}
      testId="insert-function"
      okDisabled={!current}
      onOk={() => {
        if (!current) return false;
        // Opening the dialog ended any edit but a formula, which takes the function at its caret.
        // The dialog closes first so the cell editor keeps the focus.
        window.setTimeout(() => insertFunction(ui, current.name), 0);
        return true;
      }}
    >
      <Field label="Search for a function:" wide>
        <input value={query} placeholder="Type a brief description of what you want to do" data-testid="insert-function-search" onChange={(e) => setQuery(e.target.value)} />
      </Field>
      <Field label="Or select a category:">
        <select value={cat} data-testid="insert-function-category" onChange={(e) => setCat(e.target.value)}>
          <option value="All">All</option>
          {categories.map((c) => <option key={c} value={c}>{CATEGORY_LABELS[c] ?? c}</option>)}
        </select>
      </Field>
      <div className="sh-label">Select a function:</div>
      <ListBox items={list} value={current} keyOf={(f) => f.name} render={(f) => f.name} height={200} testId="insert-function-list" label="Functions" onChange={setSel} onActivate={(f) => { onClose(); window.setTimeout(() => insertFunction(ui, f.name), 0); }} />
      {current && (
        <div className="sh-fn-help" data-testid="insert-function-help">
          <b>{signature(current)}</b>
          <p>{current.description || 'No description available.'}</p>
        </div>
      )}
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Name Manager

type Filter = 'all' | 'sheet' | 'workbook' | 'errors' | 'noErrors' | 'defined' | 'tables';

interface NameRow {
  key: string;
  name: string;
  formula: string;
  localSheet?: number;
  comment?: string;
  table?: boolean;
}

export function NameManagerDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const [filter, setFilter] = useState<Filter>('all');
  const [sel, setSel] = useState<string | null>(null);
  const [edit, setEdit] = useState<{ row: NameRow | null } | null>(null);
  const [refersTo, setRefersTo] = useState<string | null>(null);
  const [version, setVersion] = useState(0);
  const rows = useMemo<NameRow[]>(() => {
    void version;
    const names: NameRow[] = doc.wb.names.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.')).map((n: DefinedName) => ({ key: `${n.name}|${n.localSheet ?? ''}`, name: n.name, formula: n.formula, localSheet: n.localSheet, comment: n.comment }));
    const tables: NameRow[] = doc.wb.sheets.flatMap((s) => s.tables.map((t) => ({ key: `table|${t.name}`, name: t.name, formula: `${s.name}!${rangeText(t.ref, s)}`, table: true })));
    return [...names, ...tables].sort((a, b) => a.name.localeCompare(b.name));
  }, [doc, version]);
  const value = (r: NameRow) => {
    try {
      const v = doc.engine.evaluateText(r.formula, doc.wb.sheets[r.localSheet ?? doc.activeIndex] ?? doc.sheet, 0, 0);
      if (v instanceof CellError) return v.code;
      if (v && typeof v === 'object' && 'areas' in (v as object)) return '{…}';
      return v === null ? '' : String(v);
    } catch {
      return '#NAME?';
    }
  };
  const shown = rows.filter((r) => {
    switch (filter) {
      case 'sheet': return r.localSheet !== undefined;
      case 'workbook': return r.localSheet === undefined;
      case 'errors': return value(r).startsWith('#');
      case 'noErrors': return !value(r).startsWith('#');
      case 'defined': return !r.table;
      case 'tables': return !!r.table;
      default: return true;
    }
  });
  const current = shown.find((r) => r.key === sel) ?? null;
  const scopeName = (r: NameRow) => (r.localSheet === undefined ? 'Workbook' : doc.wb.sheets[r.localSheet]?.name ?? '');
  return (
    <>
      <Dialog title="Name Manager" onClose={onClose} closeOnly width={720} testId="name-manager">
        <div className="sh-row wrap sh-toolbar-row">
          <button type="button" className="sh-btn" data-testid="name-new" onClick={() => setEdit({ row: null })}>New...</button>
          <button type="button" className="sh-btn" disabled={!current || current.table} onClick={() => current && setEdit({ row: current })}>Edit...</button>
          <button
            type="button"
            className="sh-btn"
            disabled={!current || current.table}
            data-testid="name-delete"
            onClick={async () => {
              if (!current) return;
              if (!(await ui.confirm(`Are you sure you want to delete the name ${current.name}?`, { title: 'Name Manager', ok: 'OK' }))) return;
              doc.removeName(current.name, current.localSheet);
              setSel(null);
              setVersion(version + 1);
            }}
          >
            Delete
          </button>
          <span className="grow" />
          <Field label="Filter:">
            <select value={filter} onChange={(e) => setFilter(e.target.value as Filter)}>
              <option value="all">Clear Filter</option>
              <option value="sheet">Names Scoped to Worksheet</option>
              <option value="workbook">Names Scoped to Workbook</option>
              <option value="errors">Names with Errors</option>
              <option value="noErrors">Names without Errors</option>
              <option value="defined">Defined Names</option>
              <option value="tables">Table Names</option>
            </select>
          </Field>
        </div>
        <div className="sh-table names" role="table" data-testid="name-list">
          <div className="sh-table-head" role="row">
            <span>Name</span>
            <span>Value</span>
            <span>Refers To</span>
            <span>Scope</span>
            <span>Comment</span>
          </div>
          <div className="sh-table-body">
            {shown.map((r) => (
              <div key={r.key} role="row" className={`sh-table-row${r.key === sel ? ' selected' : ''}`} onClick={() => { setSel(r.key); setRefersTo(null); }} onDoubleClick={() => !r.table && setEdit({ row: r })}>
                <span>{r.table ? '▦ ' : ''}{r.name}</span>
                <span>{value(r)}</span>
                <span>={r.formula}</span>
                <span>{scopeName(r)}</span>
                <span>{r.comment}</span>
              </div>
            ))}
            {!shown.length && <div className="sh-table-empty">No names.</div>}
          </div>
        </div>
        <div className="sh-row">
          <span className="sh-label">Refers to:</span>
          <input
            className="grow"
            value={refersTo ?? (current ? `=${current.formula}` : '')}
            disabled={!current || current.table}
            aria-label="Refers to"
            onChange={(e) => setRefersTo(e.target.value)}
          />
          <button type="button" className="sh-btn icon" aria-label="Commit" disabled={refersTo === null || !current} onClick={() => {
            if (!current || refersTo === null) return;
            const error = doc.defineName({ name: current.name, formula: refersTo, localSheet: current.localSheet, comment: current.comment }, { name: current.name, localSheet: current.localSheet });
            if (error) ui.notify(error);
            setRefersTo(null);
            setVersion(version + 1);
          }}>✓</button>
          <button type="button" className="sh-btn icon" aria-label="Cancel" disabled={refersTo === null} onClick={() => setRefersTo(null)}>✕</button>
        </div>
      </Dialog>
      {edit && (
        <NewNameDialog
          name={edit.row?.name}
          scope={edit.row?.localSheet}
          onClose={() => {
            setEdit(null);
            setVersion(version + 1);
          }}
        />
      )}
    </>
  );
}

export function NewNameDialog({ name, scope, onClose }: { name?: string; scope?: number; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const existing = name !== undefined ? doc.wb.names.find((n) => n.name === name && n.localSheet === scope) : undefined;
  const suggested = useMemo(() => {
    if (existing) return existing.name;
    const { r, c } = doc.selection.active;
    // Excel suggests the text of the cell above or to the left.
    const near = [doc.displayText(r, c), r > 0 ? doc.displayText(r - 1, c) : '', c > 0 ? doc.displayText(r, c - 1) : ''].find((t) => /^[A-Za-z_][\w.]*$/.test(t.replace(/\s+/g, '_')) && Number.isNaN(Number(t)));
    return near ? near.replace(/\s+/g, '_') : '';
  }, [existing, doc]);
  const [n, setN] = useState(suggested);
  const [sc, setSc] = useState<number | -1>(scope ?? -1);
  const [comment, setComment] = useState(existing?.comment ?? '');
  const [ref, setRef] = useState(existing ? `=${existing.formula}` : `=${doc.selection.ranges.map((x) => `${doc.sheet.name.includes(' ') || /[^\w]/.test(doc.sheet.name) ? `'${doc.sheet.name.replace(/'/g, "''")}'` : doc.sheet.name}!${rangeText(doc.bounded(x, doc.sheet) ?? x, doc.sheet)}`).join(',')}`);
  return (
    <Dialog
      title={existing ? 'Edit Name' : 'New Name'}
      onClose={onClose}
      width={420}
      testId="new-name"
      onOk={() => {
        const error = doc.defineName({ name: n, formula: ref, localSheet: sc === -1 ? undefined : sc, comment }, existing ? { name: existing.name, localSheet: existing.localSheet } : undefined);
        if (error) {
          ui.notify(error);
          return false;
        }
        return true;
      }}
    >
      <Field label="Name:" wide>
        <input value={n} autoFocus data-testid="new-name-name" onChange={(e) => setN(e.target.value)} />
      </Field>
      <Field label="Scope:">
        <select value={sc} data-testid="new-name-scope" onChange={(e) => setSc(Number(e.target.value))}>
          <option value={-1}>Workbook</option>
          {doc.wb.sheets.map((s, i) => <option key={s.uid} value={i}>{s.name}</option>)}
        </select>
      </Field>
      <Field label="Comment:" wide>
        <textarea rows={3} value={comment} onChange={(e) => setComment(e.target.value)} />
      </Field>
      <Field label="Refers to:" wide>
        <RangeField value={ref} onChange={(v) => setRef(v.startsWith('=') ? v : `=${v}`)} label="Refers to" testId="new-name-ref" multi />
      </Field>
    </Dialog>
  );
}

export function NamesFromSelectionDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const range = doc.bounded(doc.range, doc.sheet) ?? doc.range;
  const [top, setTop] = useState(range.r1 > range.r0);
  const [left, setLeft] = useState(range.c1 > range.c0);
  const [bottom, setBottom] = useState(false);
  const [right, setRight] = useState(false);
  return (
    <Dialog
      title="Create Names from Selection"
      onClose={onClose}
      width={320}
      testId="names-from-selection"
      onOk={() => {
        const sheet = doc.sheet;
        const q = /[^\w]/.test(sheet.name) ? `'${sheet.name.replace(/'/g, "''")}'` : sheet.name;
        const errors: string[] = [];
        const add = (label: string, r0: number, c0: number, r1: number, c1: number) => {
          const name = label.trim().replace(/\s+/g, '_').replace(/^(\d)/, '_$1');
          if (!name) return;
          const error = doc.defineName({ name, formula: `${q}!${rangeText(makeRange(r0, c0, r1, c1), sheet)}` }, doc.wb.names.some((x) => x.name.toUpperCase() === name.toUpperCase() && x.localSheet === undefined) ? { name } : undefined);
          if (error) errors.push(name);
        };
        const r0 = range.r0 + (top ? 1 : 0);
        const r1 = range.r1 - (bottom ? 1 : 0);
        const c0 = range.c0 + (left ? 1 : 0);
        const c1 = range.c1 - (right ? 1 : 0);
        doc.transact('Create names', () => {
          if (top) for (let c = c0; c <= c1; c++) add(doc.displayText(range.r0, c), r0, c, r1, c);
          if (bottom) for (let c = c0; c <= c1; c++) add(doc.displayText(range.r1, c), r0, c, r1, c);
          if (left) for (let r = r0; r <= r1; r++) add(doc.displayText(r, range.c0), r, c0, r, c1);
          if (right) for (let r = r0; r <= r1; r++) add(doc.displayText(r, range.c1), r, c0, r, c1);
        });
        if (errors.length) ui.notify(`Some names could not be created: ${errors.slice(0, 5).join(', ')}`);
        return true;
      }}
    >
      <p className="sh-lead">Create names from values in the:</p>
      <Check label="Top row" checked={top} onChange={setTop} testId="names-top" />
      <Check label="Left column" checked={left} onChange={setLeft} testId="names-left" />
      <Check label="Bottom row" checked={bottom} onChange={setBottom} />
      <Check label="Right column" checked={right} onChange={setRight} />
    </Dialog>
  );
}
