/**
 * Status bar: the cell mode (Ready, Enter, Edit), the summary of the selection (Average, Count,
 * Sum; more through the right-click menu) and the zoom slider.
 */
import { useMemo, useState } from 'react';
import { usePersistent, useUI } from './context';
import { Icon } from './Icon';
import { Menu } from './Menu';
import { cmd, type MenuItem, selectionStats, sliderToZoom, STATUS_STAT_LABELS as LABELS, type StatusStat as Stat, zoomToSlider } from '@sheetline/core/views';

export function StatusBar() {
  const ui = useUI();
  const { doc, state, edit } = ui;
  const [shown, setShown] = usePersistent<Stat[]>('status-stats', ['average', 'count', 'sum']);
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  const zoom = doc.sheet.zoom;

  // eslint-disable-next-line react-hooks/exhaustive-deps
  const stats = useMemo(() => (state.multiCell ? selectionStats(doc) : null), [doc, state.version, state.multiCell]);

  const mode = edit.active ? (edit.mode === 'edit' ? 'Edit' : 'Enter') : 'Ready';
  const menuItems: MenuItem[] = (Object.keys(LABELS) as Stat[]).map((k) => ({
    title: LABELS[k],
    checked: shown.includes(k),
    keepOpen: true,
    onClick: () => setShown(shown.includes(k) ? shown.filter((x) => x !== k) : [...shown, k]),
  }));

  return (
    <div className="sh-statusbar" data-testid="status-bar" onContextMenu={(e) => { e.preventDefault(); setMenu({ x: e.clientX, y: e.clientY }); }}>
      <span className="sh-status-mode" data-testid="status-mode">{mode}</span>
      {doc.calcPending && <button type="button" className="sh-status-calc" onClick={() => doc.calculateNow()}>Calculate</button>}
      {state.sheetProtected && <span className="sh-status-flag" title="This sheet is protected"><Icon name="lock" /></span>}
      <span className="sh-status-grow" />
      {stats && (
        <span className="sh-status-stats" data-testid="status-stats">
          {(Object.keys(LABELS) as Stat[]).filter((k) => shown.includes(k) && stats[k] !== null).map((k) => (
            <span key={k} data-testid={`status-${k}`}>{LABELS[k]}: {stats[k]}</span>
          ))}
        </span>
      )}
      {!ui.phone && (
        <span className="sh-zoom">
          <button type="button" aria-label="Zoom out" data-testid="zoom-out" onClick={() => cmd.zoomTo(ui, Math.max(0.1, Math.round((zoom - 0.1) * 10) / 10))}>
            <Icon name="remove" />
          </button>
          <input
            type="range"
            min={0}
            max={100}
            step={1}
            aria-label="Zoom"
            data-testid="zoom-slider"
            value={zoomToSlider(zoom)}
            onChange={(e) => cmd.zoomTo(ui, sliderToZoom(Number(e.target.value)))}
          />
          <button type="button" aria-label="Zoom in" data-testid="zoom-in" onClick={() => cmd.zoomTo(ui, Math.min(4, Math.round((zoom + 0.1) * 10) / 10))}>
            <Icon name="add" />
          </button>
        </span>
      )}
      <button type="button" className="sh-zoom-value" aria-label="Zoom level" data-testid="zoom-value" onClick={() => ui.openDialog({ kind: 'zoom' })}>
        {Math.round(zoom * 100)}%
      </button>
      {menu && <Menu items={menuItems} at={menu} title="Customize Status Bar" onClose={() => setMenu(null)} />}
    </div>
  );
}
