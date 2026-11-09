/**
 * Sheet tab strip under the grid: previous / next buttons, the list of all sheets, the tabs (click,
 * double click to rename, drag to reorder, right click or long press for the tab menu) and "+".
 */
import { useEffect, useRef, useState } from 'react';
import { useUI } from './context';
import { Icon } from './Icon';
import { Menu } from './Menu';
import { cmd } from '@sheetline/core/views';

export function SheetBar({ onTabMenu }: { onTabMenu: (index: number, at: { x: number; y: number }) => void }) {
  const ui = useUI();
  const { doc, state } = ui;
  const [editing, setEditing] = useState<number | null>(null);
  const [draft, setDraft] = useState('');
  const [drag, setDrag] = useState<{ from: number; over: number; side: 'before' | 'after' } | null>(null);
  const [list, setList] = useState<DOMRect | null>(null);
  const press = useRef<{ timer: number; index: number; x: number; y: number; moved: boolean; pointerId: number } | null>(null);
  const strip = useRef<HTMLDivElement>(null);
  const listBtn = useRef<HTMLButtonElement>(null);
  const locked = state.workbookProtected;

  // Keep the active tab in view.
  useEffect(() => {
    strip.current?.querySelector<HTMLElement>('.sh-sheet-tab.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [state.activeIndex]);

  const commitRename = () => {
    if (editing === null) return;
    const index = editing;
    const name = draft.trim();
    setEditing(null);
    if (!name || name === doc.wb.sheets[index]?.name) return;
    const error = doc.renameSheet(index, name);
    if (error) ui.notify(error);
  };

  const step = (dir: 1 | -1) => {
    const visible = state.sheets.filter((s) => s.state === 'visible').map((s) => s.index);
    const pos = visible.indexOf(state.activeIndex);
    const next = visible[pos + dir];
    if (next !== undefined) doc.setActiveSheet(next);
  };

  const onPointerDown = (e: React.PointerEvent, index: number) => {
    if (e.button !== 0 || editing !== null) return;
    const timer = window.setTimeout(() => {
      if (press.current && !press.current.moved) {
        doc.setActiveSheet(index);
        onTabMenu(index, { x: press.current.x, y: press.current.y });
        press.current = null;
      }
    }, 500);
    press.current = { timer, index, x: e.clientX, y: e.clientY, moved: false, pointerId: e.pointerId };
  };

  const onPointerMove = (e: React.PointerEvent) => {
    const p = press.current;
    if (!p || p.pointerId !== e.pointerId) return;
    if (!p.moved && Math.hypot(e.clientX - p.x, e.clientY - p.y) > 8) {
      p.moved = true;
      window.clearTimeout(p.timer);
      if (e.pointerType === 'mouse' && !state.readOnly && !locked) {
        (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
        setDrag({ from: p.index, over: p.index, side: 'after' });
      }
    }
    if (drag) {
      const el = document.elementFromPoint(e.clientX, e.clientY)?.closest<HTMLElement>('[data-tab-index]');
      if (el) {
        const over = Number(el.dataset.tabIndex);
        const rect = el.getBoundingClientRect();
        setDrag({ ...drag, over, side: e.clientX < rect.left + rect.width / 2 ? 'before' : 'after' });
      }
    }
  };

  const onPointerUp = (e: React.PointerEvent, index: number) => {
    const p = press.current;
    if (p) window.clearTimeout(p.timer);
    press.current = null;
    if (drag) {
      const { from, over, side } = drag;
      setDrag(null);
      let to = side === 'before' ? over : over + 1;
      if (from < to) to--;
      if (to !== from) doc.moveSheet(from, Math.max(0, Math.min(doc.wb.sheets.length - 1, to)));
      return;
    }
    if (p && !p.moved && e.button === 0) {
      if (ui.grid?.isEditing() && !ui.grid.editState.text.startsWith('=')) ui.grid.commitEdit();
      doc.setActiveSheet(index);
    }
  };

  return (
    <div className="sh-sheetbar" data-testid="sheet-tabs">
      <button type="button" className="sh-sheet-nav" aria-label="Previous sheet" data-testid="tab-prev" onClick={() => step(-1)}>
        <Icon name="chevron_left" />
      </button>
      <button type="button" className="sh-sheet-nav" aria-label="Next sheet" data-testid="tab-next" onClick={() => step(1)}>
        <Icon name="chevron_right" />
      </button>
      <button ref={listBtn} type="button" className="sh-sheet-nav" aria-label="All sheets" title="All sheets" data-testid="tab-list" onClick={() => setList(listBtn.current!.getBoundingClientRect())}>
        <Icon name="menu" />
      </button>
      <div ref={strip} className="sh-sheet-tabs" role="tablist" aria-label="Worksheets">
        {state.sheets.map((s) => {
          if (s.state !== 'visible') return null;
          const active = s.index === state.activeIndex;
          if (editing === s.index) {
            return (
              <input
                key={s.index}
                className="sh-sheet-rename"
                aria-label="Sheet name"
                data-testid="tab-rename-input"
                autoFocus
                value={draft}
                onFocus={(e) => e.target.select()}
                onChange={(e) => setDraft(e.target.value)}
                onBlur={commitRename}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') commitRename();
                  else if (e.key === 'Escape') setEditing(null);
                }}
              />
            );
          }
          const dropClass = drag && drag.over === s.index && drag.from !== s.index ? ` drop-${drag.side}` : '';
          return (
            <button
              key={s.index}
              type="button"
              role="tab"
              aria-selected={active}
              data-tab-index={s.index}
              data-testid={`tab-${s.name}`}
              className={`sh-sheet-tab${active ? ' active' : ''}${dropClass}`}
              style={s.color ? ({ '--tab-color': s.color } as React.CSSProperties) : undefined}
              onPointerDown={(e) => onPointerDown(e, s.index)}
              onPointerMove={onPointerMove}
              onPointerUp={(e) => onPointerUp(e, s.index)}
              onPointerCancel={() => {
                if (press.current) window.clearTimeout(press.current.timer);
                press.current = null;
                setDrag(null);
              }}
              onDoubleClick={() => {
                if (state.readOnly || locked) return;
                setDraft(s.name);
                setEditing(s.index);
              }}
              onContextMenu={(e) => {
                e.preventDefault();
                doc.setActiveSheet(s.index);
                onTabMenu(s.index, { x: e.clientX, y: e.clientY });
              }}
            >
              {s.name}
            </button>
          );
        })}
      </div>
      {!state.readOnly && ui.view.newTab && (
        <button type="button" className="sh-sheet-add" aria-label="New sheet" title="New sheet (Shift+F11)" data-testid="tab-add" disabled={locked} onClick={() => cmd.newWorksheet(ui)}>
          <Icon name="add" />
        </button>
      )}
      <div className="sh-sheetbar-rest" />
      {list && (
        <Menu
          at={list}
          sheet={ui.phone}
          title="Sheets"
          testId="tab-list-menu"
          onClose={() => setList(null)}
          items={state.sheets.filter((s) => s.state === 'visible').map((s) => ({ title: s.name, checked: s.index === state.activeIndex, onClick: () => doc.setActiveSheet(s.index) }))}
        />
      )}
    </div>
  );
}
