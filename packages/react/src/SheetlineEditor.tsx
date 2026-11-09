/**
 * The editor: ribbon (or one-row toolbar), formula bar, grid, sheet tabs and status bar, with the
 * File backstage, side panels (comments, selection, format), context menus, dialogs and the cell
 * pop-ups (comments, notes, filters).
 */
import { useEffect, useRef, useState, useSyncExternalStore, type CSSProperties } from 'react';
import { createPortal } from 'react-dom';
import { type SheetlineOptions, SheetlineUI, RIBBON_TABS, rootClassOf, type UIState } from '@sheetline/core/views';
import { UIContext, useUI } from './context';
import { Menu } from './Menu';
import { FormulaAssist, FormulaBar } from './FormulaBar';
import { SheetBar } from './SheetBar';
import { StatusBar } from './StatusBar';
import { KeyBar } from './KeyBar';
import { CommentsBox, FilterPopup, NoteEditor } from './Popovers';
import { ConfirmDialog, PromptDialog } from './dialogs/SmallDialogs';
import { DialogHost } from './dialogs/DialogHost';
import { Backstage } from './backstage/Backstage';
import { Ribbon } from './ribbon/Ribbon';
import { SidePanels } from './panels/SidePanels';
import { Icon } from './Icon';

export interface SheetlineEditorProps extends SheetlineOptions {
  className?: string;
  style?: CSSProperties;
}

/** The grid's element; the grid of each workbook is created in it. */
function GridHost({ state }: { state: UIState }) {
  const { ui } = useUI();
  const host = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = host.current!;
    ui.attachGrid(el);
    return () => ui.detachGrid(el);
  }, [ui]);
  const { settings } = state;
  return <div className={`sh-grid-host${settings.showHScroll === false ? ' no-hscroll' : ''}${settings.showVScroll === false ? ' no-vscroll' : ''}`} ref={host} data-testid="grid" />;
}

/** What the editor shows for one workbook (drawn afresh for each workbook opened). */
function WorkbookView({ state }: { state: UIState }) {
  const ctx = useUI();
  const { ui } = ctx;
  const { phone, layout, theme, collapsed, view, panel, backstage, edit, menu, dialog, comments, note, filter, pending, toast } = state;
  return (
    <>
      <Ribbon
        tabs={RIBBON_TABS}
        active={ui.currentTab()}
        onTab={ui.showTab}
        onFile={() => ui.openBackstage('new')}
        layout={phone ? 'ribbon' : layout}
        collapsed={collapsed && !phone}
        onCollapse={ui.setCollapsed}
        extra={
          <>
            <button
              type="button"
              className="sh-theme-toggle"
              title={theme === 'dark' ? 'Light theme' : 'Dark theme'}
              aria-label="Switch Theme"
              data-testid="theme-toggle"
              onClick={() => ui.setTheme(theme === 'dark' ? 'light' : 'dark')}
            >
              <Icon name={theme === 'dark' ? 'light_mode' : 'dark_mode'} />
            </button>
            {!phone && (
              <label className="sh-switch" title="Switch Layout (ribbon / toolbar)">
                <input type="checkbox" role="switch" aria-label="Switch Layout" data-testid="layout-toggle" checked={layout === 'toolbar'} onChange={(e) => ui.setLayout(e.target.checked ? 'toolbar' : 'ribbon')} />
                <span />
              </label>
            )}
          </>
        }
      />
      {view.formulaBar && <FormulaBar />}
      <div className="sh-main">
        <GridHost state={state} />
        {panel && <SidePanels />}
      </div>
      {view.tabStrip && <SheetBar onTabMenu={ui.openTabMenu} />}
      {view.statusBar && <StatusBar />}
      {backstage && <Backstage page={backstage} onPage={ui.setBackstagePage} onClose={() => ui.closeBackstage()} />}
      <FormulaAssist />
      {phone && edit.active && createPortal(<KeyBar />, ctx.layer)}
      {menu && <Menu items={menu.items} at={menu.at} sheet={phone} title={menu.title} testId={menu.testId} onClose={ui.closeMenu} />}
      {dialog && <DialogHost spec={dialog} onClose={() => ui.closeDialog()} />}
      {comments && <CommentsBox r={comments.r} c={comments.c} rect={comments.rect} startNew={comments.startNew} onClose={ui.closeComments} />}
      {note && <NoteEditor r={note.r} c={note.c} rect={note.rect} onClose={ui.closeNote} />}
      {filter && <FilterPopup col={filter.col} at={filter.at} tableId={filter.tableId} onClose={ui.closeFilter} />}
      {pending?.kind === 'confirm' && (
        <ConfirmDialog message={pending.message} title={pending.opts?.title} ok={pending.opts?.ok} cancel={pending.opts?.cancel} danger={pending.opts?.danger} onAnswer={(v) => ui.answer(v)} />
      )}
      {pending?.kind === 'ask' && (
        <ConfirmDialog
          message={pending.message}
          title={pending.opts?.title}
          ok={pending.opts?.ok ?? 'Yes'}
          alt={pending.opts?.alt ?? 'No'}
          cancel={pending.opts?.cancel ?? 'Cancel'}
          onAnswer={(v, alt) => ui.answer(v ? 'ok' : alt ? 'alt' : 'cancel')}
        />
      )}
      {pending?.kind === 'prompt' && <PromptDialog message={pending.message} value={pending.value} title={pending.opts?.title} ok={pending.opts?.ok} onAnswer={(v) => ui.answer(v)} />}
      {toast &&
        createPortal(
          <div className="sh-toast" role="status" data-testid="toast">
            <span>{toast}</span>
            <button type="button" onClick={ui.dismissToast}>OK</button>
          </div>,
          ctx.layer,
        )}
    </>
  );
}

/** Messages of the editor itself: recovery, errors, dragged files and files being read. */
function Messages({ state }: { state: UIState }) {
  const { ui, layer } = useUI();
  const { restore, notice, dragging, busy } = state;
  return createPortal(
    <>
      {restore && (
        <div className="sh-restore" role="status" data-testid="restore-banner">
          <span>
            Recover “{restore.fileName}” from {new Date(restore.savedAt).toLocaleString()}?
          </span>
          <button type="button" className="sh-btn primary" data-testid="restore-recover" onClick={ui.recover}>
            Recover
          </button>
          <button type="button" className="sh-btn" data-testid="restore-discard" onClick={ui.discardRecovery}>
            Discard
          </button>
        </div>
      )}
      {notice && (
        <div className="sh-toast" role="alert" data-testid="notice">
          <span>{notice}</span>
          <button type="button" onClick={() => ui.notice(null)}>OK</button>
        </div>
      )}
      {dragging && <div className="sh-drop-hint">Drop a spreadsheet to open it, or a picture to insert it</div>}
      {busy && (
        <div className="sh-busy" role="status" aria-live="polite" data-testid="busy">
          <span className="sh-spinner" aria-hidden="true" />
          {busy}
        </div>
      )}
    </>,
    layer,
  );
}

/** The complete editor. */
export function SheetlineEditor({ className, style, ...options }: SheetlineEditorProps) {
  const [ui] = useState(() => new SheetlineUI(options));
  const state = useSyncExternalStore(ui.subscribe, ui.getState, ui.getState);
  const ctx = ui.getCtx();
  const root = useRef<HTMLDivElement>(null);

  // Later props are passed on (the first ones created the interface).
  const mounted = useRef(false);
  useEffect(() => {
    if (mounted.current) ui.setOptions(options);
  });

  useEffect(() => {
    mounted.current = true;
    ui.attach(root.current!);
    return () => ui.detach();
  }, [ui]);

  return (
    <UIContext.Provider value={ctx}>
      <div ref={root} className={`${rootClassOf(state)}${className ? ` ${className}` : ''}`} style={style} data-testid="editor" data-layout={state.layout} data-theme={state.theme}>
        <WorkbookView key={state.docKey} state={state} />
      </div>
      <Messages state={state} />
    </UIContext.Provider>
  );
}
