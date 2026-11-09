/** Keyboard shortcuts of the editor outside cell editing (Ctrl+1, Ctrl+Shift+L, F9, Alt+Q...). */
import type { UIContextValue } from './context';
import * as cmd from './commands';

/** Excel shortcuts handled by the interface (the grid handles navigation and editing keys). */
export function handleShortcut(ctx: UIContextValue, e: KeyboardEvent): boolean {
  const ctrl = e.ctrlKey || e.metaKey;
  const k = e.key.toLowerCase();
  if (ctrl && !e.shiftKey && (k === '1')) {
    ctx.openDialog({ kind: 'format-cells' });
    return true;
  }
  if (ctrl && e.shiftKey && (k === 'f' || k === 'p')) {
    ctx.openDialog({ kind: 'format-cells', tab: 'font' });
    return true;
  }
  if (ctrl && e.shiftKey && k === 'l') {
    cmd.toggleFilter(ctx);
    return true;
  }
  if (ctrl && !e.shiftKey && (k === 't' || k === 'l')) {
    if (ctx.state.tableId === null) ctx.openDialog({ kind: 'create-table' });
    return true;
  }
  if (ctrl && !e.shiftKey && k === 'g') {
    ctx.openDialog({ kind: 'find', tab: 'goto' });
    return true;
  }
  if (e.key === 'F5' && !ctrl && !e.altKey) {
    ctx.openDialog({ kind: 'find', tab: 'goto' });
    return true;
  }
  if (e.key === 'F9') {
    if (e.shiftKey) ctx.doc.calculateSheet();
    else ctx.doc.calculateNow();
    return true;
  }
  if (e.key === 'F3' && e.shiftKey) {
    ctx.openDialog({ kind: 'insert-function' });
    return true;
  }
  if (ctrl && e.key === 'F3') {
    ctx.openDialog({ kind: 'name-manager' });
    return true;
  }
  if (e.altKey && e.key === 'F1') {
    ctx.runCommand('insert-chart-default');
    return true;
  }
  if (ctrl && e.shiftKey && (e.key === '%' || e.key === '5')) {
    cmd.setNumberFormat(ctx, '0%');
    return true;
  }
  if (ctrl && e.shiftKey && (e.key === '$' || e.key === '4')) {
    cmd.setNumberFormat(ctx, '"$"#,##0.00_);("$"#,##0.00)');
    return true;
  }
  if (ctrl && e.shiftKey && (e.key === '!' || e.key === '1')) {
    cmd.setNumberFormat(ctx, '#,##0.00');
    return true;
  }
  if (ctrl && e.shiftKey && (e.key === '~' || e.key === '`')) {
    cmd.setNumberFormat(ctx, 'General');
    return true;
  }
  if (ctrl && (e.key === '`' || e.key === '~') && !e.shiftKey) {
    ctx.setView({ showFormulas: !ctx.view.showFormulas });
    return true;
  }
  if (ctrl && e.key === '9') {
    cmd.setHidden(ctx, 'row', !e.shiftKey);
    return true;
  }
  if (ctrl && e.key === '0') {
    cmd.setHidden(ctx, 'col', !e.shiftKey);
    return true;
  }
  if (ctrl && (e.key === '+' || (e.shiftKey && e.key === '='))) {
    ctx.openDialog({ kind: 'insert-cells' });
    return true;
  }
  if (ctrl && e.key === '-') {
    ctx.openDialog({ kind: 'delete-cells' });
    return true;
  }
  if (ctrl && e.altKey && k === 'v') {
    ctx.openDialog({ kind: 'paste-special' });
    return true;
  }
  // Alt+Q: the search box ("Tell me") of Excel.
  if (e.altKey && !ctrl && k === 'q') {
    ctx.openDialog({ kind: 'command-palette' });
    return true;
  }
  if (e.key === 'Escape' && ctx.painter) {
    ctx.setPainter(null);
    return true;
  }
  return false;
}
