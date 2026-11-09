/** Pop-ups anchored to cells: the threaded comment card, the note editor and the column filter drop-down. */
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { avatarOf, cellName, type ColumnFilter, commentStamp, FILTER_OPERATORS, type FilterCondition, type FilterOperator, type MenuItem, NUMBER_CONDITIONS, placeNextTo, TEXT_CONDITIONS } from '@sheetline/core/views';
import { useUI } from './context';
import { Icon } from './Icon';
import { Menu } from './Menu';
import { Modal } from './Modal';

/** A fixed pop-up next to a rectangle, closed by clicks outside and Escape. */
export function Anchored({ rect, className, onClose, children, testId, label, width, below }: { rect: DOMRect; className: string; onClose: () => void; children: ReactNode; testId?: string; label?: string; width?: number; below?: boolean }) {
  const { layer } = useUI();
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    if (ref.current) placeNextTo(ref.current, rect, below);
  });
  useEffect(() => {
    const down = (e: PointerEvent) => {
      const t = e.target as HTMLElement;
      if (t.closest(`.${className.split(' ')[0]}`) || t.closest('.sh-popup') || t.closest('.sh-bottom-sheet') || t.closest('.sh-sheet-backdrop') || t.closest('.sh-modal-layer')) return;
      onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !layer.querySelector('.sh-modal-layer')) onClose();
    };
    const id = window.setTimeout(() => {
      window.addEventListener('pointerdown', down, true);
      window.addEventListener('keydown', key);
    }, 0);
    return () => {
      window.clearTimeout(id);
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('keydown', key);
    };
  }, [onClose, className, layer]);
  return createPortal(
    <div ref={ref} className={className} role="dialog" aria-label={label} data-testid={testId} style={width ? { width } : undefined}>
      {children}
    </div>,
    layer,
  );
}

export function Avatar({ name }: { name: string }) {
  const { initial, color } = avatarOf(name);
  return <span className="sh-avatar" style={{ background: color }} aria-hidden="true">{initial}</span>;
}

/** A threaded comment card: the thread, replies and the reply box (or a new comment box). */
export function CommentThreadCard({ r, c, startNew, onDone, inPanel }: { r: number; c: number; startNew?: boolean; onDone?: () => void; inPanel?: boolean }) {
  const ui = useUI();
  const { doc } = ui;
  const thread = doc.sheet.commentAt(r, c);
  const [text, setText] = useState('');
  const [editing, setEditing] = useState<{ id: string; text: string } | null>(null);
  const [menu, setMenu] = useState<{ id: string; at: DOMRect } | null>(null);
  const ro = ui.state.readOnly;
  const post = () => {
    if (!text.trim()) return;
    if (thread?.kind === 'threaded') doc.reply(thread.comments[0].id, text.trim());
    else doc.addComment(r, c, text.trim());
    setText('');
    if (startNew && !thread) onDone?.();
  };
  const isNew = !thread || thread.kind !== 'threaded';
  return (
    <div className={`sh-thread${thread?.resolved ? ' resolved' : ''}`} data-testid="comment-thread">
      {isNew ? (
        <div className="sh-thread-head">
          <Avatar name={doc.author} />
          <span className="who">{doc.author}</span>
          <span className="ref">{cellName(r, c)}</span>
        </div>
      ) : (
        thread.comments.map((cm, i) => (
          <div key={cm.id} className={`sh-comment${i ? ' reply' : ''}`} data-testid="comment">
            <div className="sh-thread-head">
              <Avatar name={cm.author || 'Guest'} />
              <span className="who">{cm.author || 'Guest'}</span>
              {i === 0 && <span className="ref">{cellName(r, c)}</span>}
              {!ro && (
                <button type="button" className="sh-icon-btn" aria-label="Thread actions" data-testid="comment-options" onClick={(e) => setMenu({ id: cm.id, at: e.currentTarget.getBoundingClientRect() })}>
                  <Icon name={i === 0 ? 'settings' : 'more_horiz'} />
                </button>
              )}
            </div>
            {editing?.id === cm.id ? (
              <div className="sh-thread-edit">
                <textarea className="sh-textarea" value={editing.text} onChange={(e) => setEditing({ ...editing, text: e.target.value })} />
                <div className="row">
                  <button type="button" className="sh-btn-primary" onClick={() => { doc.editComment(cm.id, editing.text); setEditing(null); }}>Save</button>
                  <button type="button" className="sh-btn-secondary" onClick={() => setEditing(null)}>Cancel</button>
                </div>
              </div>
            ) : (
              <div className="body">{cm.text}</div>
            )}
            <div className="when">{commentStamp(cm.created)}{i === 0 && thread.resolved ? ' · Resolved' : ''}</div>
          </div>
        ))
      )}
      {!ro && !thread?.resolved && (
        <div className="sh-thread-input">
          <textarea
            className="sh-textarea"
            value={text}
            autoFocus={!inPanel}
            placeholder={isNew ? 'Start a conversation' : '@mention or reply'}
            aria-label={isNew ? 'Comment' : 'Reply'}
            data-testid="comment-input"
            rows={isNew ? 3 : 1}
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
                e.preventDefault();
                post();
              }
            }}
          />
          {(text || isNew) && (
            <div className="row">
              <button type="button" className="sh-send" aria-label="Post" title="Post (Ctrl+Enter)" disabled={!text.trim()} data-testid="comment-save" onClick={post}>
                <Icon name="arrow_forward" />
              </button>
              <button type="button" className="sh-icon-btn" aria-label="Cancel" onClick={() => (isNew ? onDone?.() : setText(''))}>
                <Icon name="close" />
              </button>
            </div>
          )}
        </div>
      )}
      {menu && thread && (
        <Menu
          at={menu.at}
          onClose={() => setMenu(null)}
          items={[
            { title: 'Edit Comment', icon: 'edit', onClick: () => setEditing({ id: menu.id, text: thread.comments.find((x) => x.id === menu.id)?.text ?? '' }) },
            ...(menu.id === thread.comments[0].id ? [
              { title: thread.resolved ? 'Reopen Thread' : 'Resolve Thread', icon: 'check_circle' as const, onClick: () => doc.setResolved(menu.id, !thread.resolved) },
              { title: 'Delete Thread', icon: 'delete' as const, onClick: () => doc.deleteComment(menu.id) },
            ] : [{ title: 'Delete Reply', icon: 'delete' as const, onClick: () => doc.deleteComment(menu.id) }]),
          ]}
        />
      )}
    </div>
  );
}

export function CommentsBox({ r, c, rect, startNew, onClose }: { r: number; c: number; rect: DOMRect; startNew?: boolean; onClose: () => void }) {
  return (
    <Anchored rect={rect} className="sh-comment-card" onClose={onClose} testId="comments-box" label={`Comment ${cellName(r, c)}`}>
      <CommentThreadCard r={r} c={c} startNew={startNew} onDone={onClose} />
    </Anchored>
  );
}

export function NoteEditor({ r, c, rect, onClose }: { r: number; c: number; rect: DOMRect; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const existing = doc.sheet.commentAt(r, c);
  const [text, setText] = useState(existing?.kind === 'note' ? existing.comments.map((x) => x.text).join('\n') : '');
  const save = () => {
    doc.setNote(r, c, text.trim());
    onClose();
  };
  return (
    <Anchored rect={rect} className="sh-note" onClose={save} testId="note-editor" label={`Note ${cellName(r, c)}`}>
      <div className="sh-note-author">{doc.author}:</div>
      <textarea value={text} autoFocus aria-label="Note" data-testid="note-input" onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Escape') { e.stopPropagation(); onClose(); } }} />
      <div className="sh-note-actions">
        <button type="button" className="sh-btn-primary" data-testid="note-save" onClick={save}>OK</button>
      </div>
    </Anchored>
  );
}

// -----------------------------------------------------------------------------------------------
// Filter drop-down

/** Excel's Custom AutoFilter dialog: two conditions joined by And / Or. */
function CustomFilterDialog({ title, initial, onApply, onClose }: { title: string; initial: ColumnFilter | undefined; onApply: (f: ColumnFilter) => void; onClose: () => void }) {
  const [c1, setC1] = useState<FilterCondition>(initial?.conditions?.[0] ?? { op: 'equal', value: '' });
  const [c2, setC2] = useState<FilterCondition>(initial?.conditions?.[1] ?? { op: 'equal', value: '' });
  const [and, setAnd] = useState(initial?.and ?? true);
  const row = (c: FilterCondition, set: (x: FilterCondition) => void, n: number) => (
    <div className="sh-form-row">
      <select className="sh-select" value={c.op} aria-label={`Operator ${n}`} data-testid={`custom-filter-op${n}`} onChange={(e) => set({ ...c, op: e.target.value as FilterOperator })}>
        {FILTER_OPERATORS.map((o) => <option key={o.op} value={o.op}>{o.label}</option>)}
      </select>
      <input className="sh-input" value={c.value} aria-label={`Value ${n}`} data-testid={`custom-filter-value${n}`} onChange={(e) => set({ ...c, value: e.target.value })} />
    </div>
  );
  return (
    <Modal
      title="Custom AutoFilter"
      onClose={onClose}
      width={460}
      blocking
      testId="custom-filter-dialog"
      footer={
        <>
          <button type="button" className="sh-btn-primary" data-testid="custom-filter-ok" onClick={() => onApply({ conditions: [c1, c2].filter((x) => x.value !== ''), and })}>OK</button>
          <button type="button" className="sh-btn-secondary" onClick={onClose}>Cancel</button>
        </>
      }
    >
      <div className="sh-form">
        <div>Show rows where: <b>{title}</b></div>
        {row(c1, setC1, 1)}
        <div className="sh-form-row">
          <label><input type="radio" checked={and} onChange={() => setAnd(true)} /> And</label>
          <label><input type="radio" checked={!and} onChange={() => setAnd(false)} /> Or</label>
        </div>
        {row(c2, setC2, 2)}
        <div className="sh-hint">Use ? to represent any single character. Use * to represent any series of characters.</div>
      </div>
    </Modal>
  );
}

export function FilterPopup({ col, at, tableId, onClose }: { col: number; at: { x: number; y: number }; tableId?: number; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const table = tableId !== undefined ? doc.tableById(tableId)?.table : undefined;
  const ref = table ? { ...table.ref, r1: table.ref.r1 - (table.totalsRow ? 1 : 0) } : doc.sheet.autoFilter?.ref;
  const current: ColumnFilter | undefined = table ? table.filters.get(col) : doc.sheet.autoFilter?.columns.get(col);
  const values = useMemo(() => doc.filterValues(col, doc.sheet, ref), [doc, col, ref]);
  const numeric = values.filter((v) => v.text !== '').every((v) => v.numeric) && values.some((v) => v.numeric);
  const header = ref ? doc.displayText(ref.r0, col) || `Column ${col + 1}` : '';
  const [search, setSearch] = useState('');
  const [custom, setCustom] = useState<ColumnFilter | null>(null);
  const [sub, setSub] = useState<DOMRect | null>(null);
  const [checked, setChecked] = useState<Set<string>>(() => {
    if (current?.values) return new Set([...current.values.map((v) => v.toLowerCase()), ...(current.blanks ? [''] : [])]);
    return new Set(values.map((v) => v.text.toLowerCase()));
  });
  const shown = values.filter((v) => !search || v.text.toLowerCase().includes(search.toLowerCase()));
  const all = shown.every((v) => checked.has(v.text.toLowerCase()));
  const rect = new DOMRect(at.x, at.y, 0, 0);
  const setFilter = (f: ColumnFilter | null) => {
    if (table) doc.setTableFilter(table.id, col, f);
    else doc.setColumnFilter(col, f);
  };
  const apply = () => {
    const visible = search ? shown : values;
    const allChecked = values.every((v) => checked.has(v.text.toLowerCase())) && !search;
    if (allChecked) setFilter(null);
    else {
      const keep = visible.filter((v) => checked.has(v.text.toLowerCase()) && v.text !== '').map((v) => v.text);
      setFilter({ values: keep, blanks: visible.some((v) => v.text === '') && checked.has('') });
    }
    onClose();
  };
  const sort = (asc: boolean) => {
    if (table) doc.sortTable(table.id, col, asc);
    else doc.sortFilterColumn(col, asc);
    onClose();
  };
  const conditions = numeric ? NUMBER_CONDITIONS : TEXT_CONDITIONS;
  const subItems: MenuItem[] = [
    ...conditions.map((c) => ({ title: c.label, onClick: () => setCustom({ conditions: [{ op: c.op, value: '' }], and: true }) })),
    { type: 'line' },
    { title: numeric ? 'Custom Filter...' : 'Custom Filter...', onClick: () => setCustom(current ?? { conditions: [], and: true }) },
  ];
  if (custom) {
    return (
      <CustomFilterDialog
        title={header}
        initial={custom}
        onClose={onClose}
        onApply={(f) => {
          setFilter(f.conditions?.length ? f : null);
          onClose();
        }}
      />
    );
  }
  return (
    <Anchored rect={rect} below className="sh-filter" onClose={() => { if (!sub) onClose(); }} testId="filter-popup" label="Filter">
      <button type="button" className="sh-filter-item" data-testid="filter-sort-asc" onClick={() => sort(true)}>
        <Icon name="sort_by_alpha" /> {numeric ? 'Sort Smallest to Largest' : 'Sort A to Z'}
      </button>
      <button type="button" className="sh-filter-item" data-testid="filter-sort-desc" onClick={() => sort(false)}>
        <Icon name="sort" /> {numeric ? 'Sort Largest to Smallest' : 'Sort Z to A'}
      </button>
      <div className="sh-filter-sep" />
      <button type="button" className="sh-filter-item" disabled={!current} data-testid="filter-clear" onClick={() => { setFilter(null); onClose(); }}>
        <Icon name="filter_alt_off" /> Clear Filter From “{header}”
      </button>
      <button type="button" className="sh-filter-item" data-testid="filter-conditions" onClick={(e) => setSub(e.currentTarget.getBoundingClientRect())}>
        <Icon name="filter_list" /> {numeric ? 'Number Filters' : 'Text Filters'} <Icon name="arrow_right" className="sub" />
      </button>
      <input className="sh-input sh-filter-search" placeholder="Search" value={search} aria-label="Search" data-testid="filter-search" onChange={(e) => setSearch(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && apply()} />
      <div className="sh-filter-values" data-testid="filter-values">
        <label>
          <input
            type="checkbox"
            checked={all}
            onChange={(e) => {
              const next = new Set(checked);
              for (const v of shown) {
                if (e.target.checked) next.add(v.text.toLowerCase());
                else next.delete(v.text.toLowerCase());
              }
              setChecked(next);
            }}
          />
          {search ? '(Select All Search Results)' : '(Select All)'}
        </label>
        {shown.map((v) => (
          <label key={v.text || '(blank)'}>
            <input
              type="checkbox"
              checked={checked.has(v.text.toLowerCase())}
              onChange={(e) => {
                const next = new Set(checked);
                if (e.target.checked) next.add(v.text.toLowerCase());
                else next.delete(v.text.toLowerCase());
                setChecked(next);
              }}
            />
            {v.text === '' ? '(Blanks)' : v.text}
          </label>
        ))}
      </div>
      <div className="sh-filter-actions">
        <button type="button" className="sh-btn-primary" data-testid="filter-ok" onClick={apply}>OK</button>
        <button type="button" className="sh-btn-secondary" onClick={onClose}>Cancel</button>
      </div>
      {sub && <Menu at={sub} items={subItems} onClose={() => setSub(null)} />}
    </Anchored>
  );
}
