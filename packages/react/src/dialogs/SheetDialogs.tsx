/**
 * Sheet dialogs: Protect / Unprotect Sheet, Protect Workbook, Move or Copy, Unhide, and Page Setup
 * (Page, Margins, Header/Footer with custom sections, Sheet).
 */
import { useState } from 'react';
import { useUI } from '../context';
import { Check, Dialog, Field, ListBox, Radio, RangeField, Tabs, parseRangeText } from './Dialog';
import { colName, commit, DEFAULT_PROTECTION, HEADER_PRESETS, type HeaderFooterSet, hfJoin, hfSections, hfText, MARGIN_PRESETS, type PageSetup, PAPER, PROTECTION_LABELS, type ProtectionOptions } from '@sheetline/core/views';

// -----------------------------------------------------------------------------------------------
// Protection

export function ProtectSheetDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [password, setPassword] = useState('');
  const [opts, setOpts] = useState<ProtectionOptions>({ ...DEFAULT_PROTECTION });
  const [confirmPw, setConfirmPw] = useState<string | null>(null);
  if (confirmPw !== null) {
    return (
      <Dialog
        title="Confirm Password"
        onClose={() => setConfirmPw(null)}
        width={340}
        testId="protect-confirm"
        onOk={() => {
          if (confirmPw !== password) {
            ui.notify('Confirmation password is not identical.');
            setConfirmPw(null);
            return false;
          }
          commit(ui);
          ui.doc.protectSheet(opts, password);
          onClose();
          return false;
        }}
      >
        <Field label="Reenter password to proceed." wide>
          <input type="password" autoFocus value={confirmPw} data-testid="protect-confirm-input" onChange={(e) => setConfirmPw(e.target.value)} />
        </Field>
        <p className="sh-hint">Caution: If you lose or forget the password, it cannot be recovered.</p>
      </Dialog>
    );
  }
  return (
    <Dialog
      title="Protect Sheet"
      onClose={onClose}
      width={360}
      testId="protect-sheet"
      onOk={() => {
        if (password) {
          setConfirmPw('');
          return false;
        }
        commit(ui);
        ui.doc.protectSheet(opts, '');
        return true;
      }}
    >
      <Check label="Protect worksheet and contents of locked cells" checked onChange={() => undefined} />
      <Field label="Password to unprotect sheet:" wide>
        <input type="password" value={password} data-testid="protect-password" onChange={(e) => setPassword(e.target.value)} />
      </Field>
      <div className="sh-label">Allow all users of this worksheet to:</div>
      <div className="sh-checklist" data-testid="protect-options">
        {PROTECTION_LABELS.map((p) => (
          <Check key={p.key} label={p.label} checked={opts[p.key]} onChange={(v) => setOpts({ ...opts, [p.key]: v })} />
        ))}
      </div>
    </Dialog>
  );
}

export function UnprotectSheetDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [password, setPassword] = useState('');
  // Sheets protected without a password unprotect at once.
  const needs = !!ui.doc.sheet.protection?.password?.legacy || !!ui.doc.sheet.protection?.password?.hashValue;
  if (!needs) {
    return (
      <Dialog title="Unprotect Sheet" onClose={onClose} width={320} testId="unprotect-sheet" onOk={() => { ui.doc.unprotectSheet(''); }}>
        <p>Remove the protection of this sheet?</p>
      </Dialog>
    );
  }
  return (
    <Dialog
      title="Unprotect Sheet"
      onClose={onClose}
      width={320}
      testId="unprotect-sheet"
      onOk={() => {
        if (!ui.doc.unprotectSheet(password)) {
          ui.notify('The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.');
          return false;
        }
        return true;
      }}
    >
      <Field label="Password:" wide>
        <input type="password" autoFocus value={password} data-testid="unprotect-password" onChange={(e) => setPassword(e.target.value)} />
      </Field>
    </Dialog>
  );
}

export function ProtectWorkbookDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [password, setPassword] = useState('');
  return (
    <Dialog title="Protect Structure and Windows" onClose={onClose} width={340} testId="protect-workbook" onOk={() => { ui.doc.protectWorkbook(password); }}>
      <div className="sh-label">Protect workbook for</div>
      <Check label="Structure" checked onChange={() => undefined} />
      <Field label="Password (optional):" wide>
        <input type="password" value={password} data-testid="protect-workbook-password" onChange={(e) => setPassword(e.target.value)} />
      </Field>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Sheets

export function MoveCopySheetDialog({ index, onClose }: { index: number; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const sheets = doc.wb.sheets;
  const [before, setBefore] = useState<number>(Math.min(index + 1, sheets.length));
  const [copy, setCopy] = useState(false);
  const items = [...sheets.map((_, i) => i), sheets.length];
  return (
    <Dialog
      title="Move or Copy"
      onClose={onClose}
      width={320}
      testId="move-copy-sheet"
      onOk={() => {
        commit(ui);
        if (copy) {
          const dup = doc.duplicateSheet(index);
          const from = dup ? doc.wb.sheets.indexOf(dup) : -1;
          // "before" indexes the list without the copy; the copy sits right after the original.
          if (from >= 0 && from !== before) doc.moveSheet(from, Math.min(doc.wb.sheets.length - 1, before));
          return true;
        }
        let to = before;
        if (to > index) to--;
        if (to !== index) doc.moveSheet(index, Math.max(0, Math.min(sheets.length - 1, to)));
        return true;
      }}
    >
      <p className="sh-lead">Move selected sheets</p>
      <div className="sh-label">Before sheet:</div>
      <ListBox items={items} value={before} keyOf={(i) => String(i)} height={180} testId="move-copy-list" label="Before sheet" render={(i) => (i === sheets.length ? '(move to end)' : sheets[i].name)} onChange={setBefore} />
      <Check label="Create a copy" checked={copy} onChange={setCopy} testId="move-copy-copy" />
    </Dialog>
  );
}

export function UnhideSheetDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const hidden = ui.doc.wb.sheets.map((s, i) => ({ s, i })).filter((x) => x.s.state === 'hidden');
  const [sel, setSel] = useState<number | null>(hidden[0]?.i ?? null);
  return (
    <Dialog
      title="Unhide"
      onClose={onClose}
      width={300}
      testId="unhide-sheet"
      okDisabled={sel === null}
      onOk={() => {
        if (sel !== null) ui.doc.setSheetHidden(sel, false);
      }}
    >
      <div className="sh-label">Unhide sheet:</div>
      <ListBox items={hidden.map((x) => x.i)} value={sel} keyOf={(i) => String(i)} height={160} testId="unhide-list" label="Hidden sheets" render={(i) => ui.doc.wb.sheets[i].name} onChange={setSel} />
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Page Setup

type PageTab = 'page' | 'margins' | 'header' | 'sheet';

function HeaderFooterEditor({ title, code, onSave, onClose }: { title: string; code: string; onSave: (code: string) => void; onClose: () => void }) {
  const [s, setS] = useState(() => hfSections(code));
  const [focus, setFocus] = useState<'left' | 'center' | 'right'>('center');
  const insert = (token: string) => setS({ ...s, [focus]: s[focus] + token });
  return (
    <Dialog title={title} onClose={onClose} width={600} testId="header-footer-editor" onOk={() => onSave(hfJoin(s))}>
      <p className="sh-hint">To format text: select the text, then choose a code. Page number &amp;P, number of pages &amp;N, date &amp;D, time &amp;T, file &amp;F, sheet &amp;A.</p>
      <div className="sh-row wrap">
        {[
          ['Page Number', '&P'],
          ['Number of Pages', '&N'],
          ['Date', '&D'],
          ['Time', '&T'],
          ['File Name', '&F'],
          ['Sheet Name', '&A'],
        ].map(([label, token]) => (
          <button key={token} type="button" className="sh-btn" title={`Insert ${label}`} onClick={() => insert(token)}>{label}</button>
        ))}
      </div>
      <div className="sh-hf-sections">
        {(['left', 'center', 'right'] as const).map((k) => (
          <label key={k}>
            <span>{k === 'left' ? 'Left section:' : k === 'center' ? 'Center section:' : 'Right section:'}</span>
            <textarea rows={5} value={s[k]} data-testid={`hf-${k}`} onFocus={() => setFocus(k)} onChange={(e) => setS({ ...s, [k]: e.target.value })} />
          </label>
        ))}
      </div>
    </Dialog>
  );
}

export function PageSetupDialog({ tab: initial = 'page', onClose }: { tab?: PageTab; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const sheet = doc.sheet;
  const [tab, setTab] = useState<PageTab>(initial);
  const [p, setP] = useState<PageSetup>(() => ({ ...sheet.page, margins: { ...sheet.page.margins }, headerFooter: { ...sheet.page.headerFooter } }));
  const titles = doc.printTitles();
  const area = doc.printArea();
  const [printArea, setPrintArea] = useState(area ? `$${colName(area.c0)}$${area.r0 + 1}:$${colName(area.c1)}$${area.r1 + 1}` : '');
  const [rowsTop, setRowsTop] = useState(titles.rows ? `$${titles.rows[0] + 1}:$${titles.rows[1] + 1}` : '');
  const [colsLeft, setColsLeft] = useState(titles.cols ? `$${colName(titles.cols[0])}:$${colName(titles.cols[1])}` : '');
  const [hfEdit, setHfEdit] = useState<'header' | 'footer' | null>(null);
  const set = (patch: Partial<PageSetup>) => setP({ ...p, ...patch });
  const setM = (k: keyof PageSetup['margins'], v: string) => {
    const n = Number(v);
    if (Number.isFinite(n) && n >= 0) setP({ ...p, margins: { ...p.margins, [k]: n } });
  };
  const setHF = (patch: Partial<HeaderFooterSet>) => setP({ ...p, headerFooter: { ...p.headerFooter, ...patch } });
  const preview = (code: string | undefined) => {
    const s = hfSections(code);
    const ctx = { page: 1, pages: 1, file: doc.fileName || 'Book1', sheet: sheet.name, now: new Date() };
    return [hfText(s.left, ctx), hfText(s.center, ctx), hfText(s.right, ctx)];
  };
  const apply = (): boolean => {
    let areaRange = null;
    if (printArea.trim()) {
      const hit = parseRangeText(doc, printArea);
      if (!hit) {
        ui.notify('The print area reference is not valid.');
        return false;
      }
      areaRange = hit.range;
    }
    const rows = rowsTop.trim() ? /^\$?(\d+):\$?(\d+)$/.exec(rowsTop.trim().replace(/^.*!/, '')) : null;
    const cols = colsLeft.trim() ? /^\$?([A-Za-z]+):\$?([A-Za-z]+)$/.exec(colsLeft.trim().replace(/^.*!/, '')) : null;
    if ((rowsTop.trim() && !rows) || (colsLeft.trim() && !cols)) {
      ui.notify('Print titles must be whole rows ($1:$2) or whole columns ($A:$B).');
      return false;
    }
    const colIndex = (s: string) => [...s.toUpperCase()].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    doc.transact('Page Setup', () => {
      doc.setPageSetup(p);
      const before = doc.printArea();
      if (JSON.stringify(before) !== JSON.stringify(areaRange)) doc.setPrintArea(areaRange);
      doc.setPrintTitles(rows ? [Number(rows[1]) - 1, Number(rows[2]) - 1] : null, cols ? [colIndex(cols[1]), colIndex(cols[2])] : null);
    });
    return true;
  };
  const headerPresets = HEADER_PRESETS.some((h) => h.code === (p.headerFooter.oddHeader ?? '')) ? HEADER_PRESETS : [...HEADER_PRESETS, { label: preview(p.headerFooter.oddHeader).filter(Boolean).join('  ') || '(custom)', code: p.headerFooter.oddHeader ?? '' }];
  const footerPresets = HEADER_PRESETS.some((h) => h.code === (p.headerFooter.oddFooter ?? '')) ? HEADER_PRESETS : [...HEADER_PRESETS, { label: preview(p.headerFooter.oddFooter).filter(Boolean).join('  ') || '(custom)', code: p.headerFooter.oddFooter ?? '' }];
  return (
    <>
      <Dialog
        title="Page Setup"
        onClose={onClose}
        width={560}
        testId="page-setup"
        onOk={apply}
        footerLeft={
          <>
            <button type="button" className="sh-btn" onClick={() => { if (apply()) { onClose(); ui.openBackstage('print'); } }}>Print...</button>
            <button type="button" className="sh-btn" data-testid="page-setup-preview" onClick={() => { if (apply()) { onClose(); ui.openBackstage('print'); } }}>Print Preview</button>
          </>
        }
      >
        <Tabs value={tab} onChange={setTab} tabs={[{ id: 'page', label: 'Page' }, { id: 'margins', label: 'Margins' }, { id: 'header', label: 'Header/Footer' }, { id: 'sheet', label: 'Sheet' }]} testId="page-setup-tabs" />
        <div className="sh-tab-body" style={{ minHeight: 280 }}>
          {tab === 'page' && (
            <div className="sh-stack">
              <fieldset className="sh-group">
                <legend>Orientation</legend>
                <div className="sh-row">
                  <Radio name="ps-orient" label="Portrait" checked={p.orientation === 'portrait'} onChange={() => set({ orientation: 'portrait' })} testId="ps-portrait" />
                  <Radio name="ps-orient" label="Landscape" checked={p.orientation === 'landscape'} onChange={() => set({ orientation: 'landscape' })} testId="ps-landscape" />
                </div>
              </fieldset>
              <fieldset className="sh-group">
                <legend>Scaling</legend>
                <div className="sh-row">
                  <Radio name="ps-scale" label="Adjust to:" checked={!p.fitToPage} onChange={() => set({ fitToPage: false })} />
                  <input value={p.scale} inputMode="numeric" style={{ width: 60 }} disabled={p.fitToPage} aria-label="Scale" data-testid="ps-scale" onChange={(e) => set({ scale: Math.max(10, Math.min(400, Number(e.target.value) || 100)) })} />
                  <span>% normal size</span>
                </div>
                <div className="sh-row">
                  <Radio name="ps-scale" label="Fit to:" checked={p.fitToPage} onChange={() => set({ fitToPage: true })} testId="ps-fit" />
                  <input value={p.fitToWidth || ''} inputMode="numeric" style={{ width: 50 }} disabled={!p.fitToPage} aria-label="Pages wide" onChange={(e) => set({ fitToWidth: Number(e.target.value) || 0 })} />
                  <span>page(s) wide by</span>
                  <input value={p.fitToHeight || ''} inputMode="numeric" style={{ width: 50 }} disabled={!p.fitToPage} aria-label="Pages tall" onChange={(e) => set({ fitToHeight: Number(e.target.value) || 0 })} />
                  <span>tall</span>
                </div>
              </fieldset>
              <Field label="Paper size:">
                <select value={p.paperSize} data-testid="ps-paper" onChange={(e) => set({ paperSize: Number(e.target.value) })}>
                  {PAPER.map((x) => <option key={x.code} value={x.code}>{x.label} {x.size}</option>)}
                </select>
              </Field>
              <Field label="Print quality:">
                <select value={p.printQuality ?? 600} onChange={(e) => set({ printQuality: Number(e.target.value) })}>
                  {[150, 300, 600, 1200].map((q) => <option key={q} value={q}>{q} dpi</option>)}
                </select>
              </Field>
              <Field label="First page number:">
                <input value={p.firstPageNumber ?? 'Auto'} style={{ width: 80 }} onChange={(e) => set({ firstPageNumber: Number(e.target.value) || undefined })} />
              </Field>
            </div>
          )}
          {tab === 'margins' && (
            <div className="sh-stack">
              <div className="sh-margins-grid">
                {(['top', 'header', 'left', 'right', 'bottom', 'footer'] as const).map((k) => (
                  <Field key={k} label={`${k[0].toUpperCase()}${k.slice(1)}:`}>
                    <input value={p.margins[k]} inputMode="decimal" style={{ width: 70 }} data-testid={`ps-margin-${k}`} onChange={(e) => setM(k, e.target.value)} />
                  </Field>
                ))}
              </div>
              <div className="sh-row wrap">
                {MARGIN_PRESETS.map((m) => <button key={m.name} type="button" className="sh-btn" onClick={() => set({ margins: { ...m.margins } })}>{m.label}</button>)}
              </div>
              <fieldset className="sh-group">
                <legend>Center on page</legend>
                <Check label="Horizontally" checked={p.centerH} onChange={(v) => set({ centerH: v })} />
                <Check label="Vertically" checked={p.centerV} onChange={(v) => set({ centerV: v })} />
              </fieldset>
              <p className="sh-hint">Margins are in inches.</p>
            </div>
          )}
          {tab === 'header' && (
            <div className="sh-stack">
              <div className="sh-hf-preview">{preview(p.headerFooter.oddHeader).map((t, i) => <span key={i}>{t}</span>)}</div>
              <Field label="Header:" wide>
                <select value={p.headerFooter.oddHeader ?? ''} data-testid="ps-header" onChange={(e) => setHF({ oddHeader: e.target.value || undefined })}>
                  {headerPresets.map((h, i) => <option key={i} value={h.code}>{h.label}</option>)}
                </select>
              </Field>
              <div className="sh-row">
                <button type="button" className="sh-btn" data-testid="ps-custom-header" onClick={() => setHfEdit('header')}>Custom Header...</button>
                <button type="button" className="sh-btn" onClick={() => setHfEdit('footer')}>Custom Footer...</button>
              </div>
              <Field label="Footer:" wide>
                <select value={p.headerFooter.oddFooter ?? ''} data-testid="ps-footer" onChange={(e) => setHF({ oddFooter: e.target.value || undefined })}>
                  {footerPresets.map((h, i) => <option key={i} value={h.code}>{h.label}</option>)}
                </select>
              </Field>
              <div className="sh-hf-preview">{preview(p.headerFooter.oddFooter).map((t, i) => <span key={i}>{t}</span>)}</div>
              <Check label="Different odd and even pages" checked={!!p.headerFooter.differentOddEven} onChange={(v) => setHF({ differentOddEven: v || undefined })} />
              <Check label="Different first page" checked={!!p.headerFooter.differentFirst} onChange={(v) => setHF({ differentFirst: v || undefined })} />
            </div>
          )}
          {tab === 'sheet' && (
            <div className="sh-stack">
              <Field label="Print area:" wide>
                <RangeField value={printArea} onChange={setPrintArea} label="Print area" testId="ps-print-area" />
              </Field>
              <fieldset className="sh-group">
                <legend>Print titles</legend>
                <Field label="Rows to repeat at top:" wide>
                  <RangeField value={rowsTop} onChange={(v) => setRowsTop(v.replace(/^\$?[A-Z]+(\$?\d+):\$?[A-Z]+(\$?\d+)$/i, '$1:$2'))} label="Rows to repeat at top" testId="ps-rows-top" />
                </Field>
                <Field label="Columns to repeat at left:" wide>
                  <RangeField value={colsLeft} onChange={(v) => setColsLeft(v.replace(/^(\$?[A-Z]+)\$?\d+:(\$?[A-Z]+)\$?\d+$/i, '$1:$2'))} label="Columns to repeat at left" testId="ps-cols-left" />
                </Field>
              </fieldset>
              <fieldset className="sh-group">
                <legend>Print</legend>
                <Check label="Gridlines" checked={p.gridLines} onChange={(v) => set({ gridLines: v })} testId="ps-gridlines" />
                <Check label="Black and white" checked={p.blackAndWhite} onChange={(v) => set({ blackAndWhite: v })} />
                <Check label="Row and column headings" checked={p.headings} onChange={(v) => set({ headings: v })} />
              </fieldset>
              <fieldset className="sh-group">
                <legend>Page order</legend>
                <Radio name="ps-order" label="Down, then over" checked={p.pageOrder === 'downThenOver'} onChange={() => set({ pageOrder: 'downThenOver' })} />
                <Radio name="ps-order" label="Over, then down" checked={p.pageOrder === 'overThenDown'} onChange={() => set({ pageOrder: 'overThenDown' })} />
              </fieldset>
            </div>
          )}
        </div>
      </Dialog>
      {hfEdit && (
        <HeaderFooterEditor
          title={hfEdit === 'header' ? 'Header' : 'Footer'}
          code={(hfEdit === 'header' ? p.headerFooter.oddHeader : p.headerFooter.oddFooter) ?? ''}
          onClose={() => setHfEdit(null)}
          onSave={(code) => setHF(hfEdit === 'header' ? { oddHeader: code || undefined } : { oddFooter: code || undefined })}
        />
      )}
    </>
  );
}
