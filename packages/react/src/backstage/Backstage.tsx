/**
 * The File backstage: New, Open, Save, Import, Export, Print (settings and live preview) and Info
 * (document properties), shown over the whole window.
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import { useUI } from '../context';
import { Icon } from '../Icon';
import { BACKSTAGE_PAGES as PAGES, buildPages, canvasMeasure, DOC_PROPERTIES as PROPS, type ExportFormat, fileBaseName as baseName, type IconName, loadPrintFonts, MARGIN_PRESETS, type Measure, type Page, pagesToPdf, pageSvg, PAPER, paperOf, type PrintChoice, printOptionsFor, printPages, printScalingKind as scalingKind, rasterize, saveFile } from '@sheetline/core/views';

export function Backstage({ page, onPage, onClose }: { page: string; onPage: (p: string) => void; onClose: () => void }) {
  const ui = useUI();
  useEffect(() => {
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !ui.layer.querySelector('.sh-modal-layer')) onClose();
    };
    window.addEventListener('keydown', key);
    return () => window.removeEventListener('keydown', key);
  }, [onClose, ui.layer]);
  return (
    <div className="sh-backstage" data-testid="backstage" role="dialog" aria-label="File">
      <nav className="sh-bs-nav" role="menu" aria-label="File">
        <button type="button" className="sh-bs-back" aria-label="Back" data-testid="backstage-back" onClick={onClose}>
          <Icon name="arrow_back" />
        </button>
        {PAGES.map((p, i) => (
          <div key={p.id}>
            {i > 0 && PAGES[i - 1].group !== p.group && <div className="sh-bs-sep" />}
            <button type="button" role="menuitemradio" aria-checked={page === p.id} className={`sh-bs-item${page === p.id ? ' active' : ''}`} data-testid={`backstage-${p.id}`} onClick={() => onPage(p.id)}>
              {p.label}
            </button>
          </div>
        ))}
      </nav>
      <section className={`sh-bs-content${page === 'print' ? ' print' : ''}`} role="tabpanel" aria-label={PAGES.find((p) => p.id === page)?.label}>
        {page === 'new' && <NewPage onClose={onClose} />}
        {page === 'open' && <OpenPage onClose={onClose} />}
        {page === 'save' && <SavePage onClose={onClose} />}
        {page === 'import' && <ImportPage onClose={onClose} />}
        {page === 'export' && <ExportPage onClose={onClose} />}
        {page === 'print' && <PrintPage onClose={onClose} />}
        {page === 'info' && <InfoPage />}
      </section>
      {ui.phone && <div className="sh-bs-phone-pad" />}
    </div>
  );
}

function NewPage({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const create = async () => {
    if (ui.doc.modified) {
      const answer = await ui.ask('Warning: Current file has been modified. Do you want to save your changes?', { title: 'Spreadsheet', ok: 'Yes', alt: 'No', cancel: 'Cancel' });
      if (answer === 'cancel') return;
      if (answer === 'ok') ui.files.save();
    }
    onClose();
    ui.files.newWorkbook();
  };
  return (
    <>
      <h1>New</h1>
      <div className="sh-bs-tiles">
        <button type="button" className="sh-bs-tile" data-testid="new-blank" aria-label="Blank workbook" onClick={() => void create()}>
          <span className="sh-blank-preview" aria-hidden="true">
            <svg viewBox="0 0 200 150" width="200" height="150">
              <rect x="0.5" y="0.5" width="199" height="149" fill="#fff" stroke="#9e9e9e" />
              <rect x="0.5" y="0.5" width="199" height="18" fill="#f3f3f3" stroke="#d4d4d4" />
              <rect x="0.5" y="0.5" width="24" height="149" fill="#f3f3f3" stroke="#d4d4d4" />
              {[1, 2, 3, 4, 5, 6].map((i) => <line key={`h${i}`} x1="24" y1={18 + i * 19} x2="200" y2={18 + i * 19} stroke="#e0e0e0" />)}
              {[1, 2].map((i) => <line key={`v${i}`} x1={24 + i * 62} y1="18" x2={24 + i * 62} y2="150" stroke="#e0e0e0" />)}
              <rect x="25" y="19" width="61" height="18" fill="none" stroke="#217346" strokeWidth="2" />
              {['A', 'B', 'C'].map((l, i) => <text key={l} x={55 + i * 62} y="13" fontSize="9" textAnchor="middle" fill="#555">{l}</text>)}
              {[1, 2, 3, 4, 5, 6, 7].map((i) => <text key={i} x="12" y={32 + (i - 1) * 19} fontSize="9" textAnchor="middle" fill="#555">{i}</text>)}
            </svg>
          </span>
          <span>Blank workbook</span>
        </button>
      </div>
    </>
  );
}

function OpenPage({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  return (
    <>
      <h1>Open</h1>
      <p className="sh-bs-lead">Open an Excel workbook (.xlsx, .xlsm, .xltx, .xltm, .xls, .xlsb) or a text file (.csv, .tsv, .txt) from this device. You can also drop a file anywhere on the window.</p>
      <button type="button" className="sh-bs-action" data-testid="open-browse" onClick={() => { onClose(); ui.files.open(); }}>
        <Icon name="folder_open" />
        <span>Browse</span>
      </button>
    </>
  );
}

function SavePage({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [name, setName] = useState(baseName(ui.doc.fileName));
  const ext = /\.xlsm$/i.test(ui.doc.fileName) ? '.xlsm' : '.xlsx';
  return (
    <>
      <h1>Save</h1>
      <p className="sh-bs-lead">Save the workbook as an Excel file. Where the browser allows it, the file you opened is updated; otherwise a copy is downloaded.</p>
      <div className="sh-bs-form">
        <label className="sh-field">
          <span>File name</span>
          <span className="sh-inline">
            <input className="sh-input" value={name} data-testid="save-name" onChange={(e) => setName(e.target.value)} />
            <span>{ext}</span>
          </span>
        </label>
        <button
          type="button"
          className="sh-bs-action"
          data-testid="save-button"
          onClick={() => {
            if (name.trim() && name.trim() !== baseName(ui.doc.fileName)) ui.files.rename(name.trim());
            onClose();
            ui.files.save();
          }}
        >
          <Icon name="save" />
          <span>Save</span>
        </button>
      </div>
    </>
  );
}

function ImportPage({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [kind, setKind] = useState<'excel' | 'csv'>('excel');
  return (
    <>
      <h1>Import</h1>
      <div className="sh-bs-columns">
        <div className="sh-bs-list" role="menu" aria-label="Import">
          <button type="button" role="menuitemradio" aria-checked={kind === 'excel'} className={kind === 'excel' ? 'active' : ''} data-testid="import-excel" onClick={() => setKind('excel')}>Excel File</button>
          <button type="button" role="menuitemradio" aria-checked={kind === 'csv'} className={kind === 'csv' ? 'active' : ''} data-testid="import-csv" onClick={() => setKind('csv')}>CSV File</button>
        </div>
        <div className="sh-bs-detail">
          <h2>{kind === 'excel' ? 'Excel File' : 'CSV File'}</h2>
          <p className="sh-bs-lead">
            {kind === 'excel'
              ? 'Opens an Excel workbook (.xlsx, .xlsm, .xltx, .xltm, Excel 97-2003 .xls, binary .xlsb) with its formulas, styles, tables, charts, pictures, comments and conditional formats.'
              : 'Opens a comma, semicolon or tab separated text file. The separator and the encoding are detected.'}
          </p>
          <button type="button" className="sh-bs-action" data-testid="import-file" onClick={() => { onClose(); ui.files.importFile(kind); }}>
            <Icon name="upload_file" />
            <span>{kind === 'excel' ? 'Import Excel File' : 'Import CSV File'}</span>
          </button>
        </div>
      </div>
    </>
  );
}

const EXCEL_TYPES: { value: ExportFormat; label: string }[] = [
  { value: 'xlsx', label: 'Excel Workbook (*.xlsx)' },
  { value: 'xltx', label: 'Excel Template (*.xltx)' },
  { value: 'xlsm', label: 'Excel Macro-Enabled Workbook (*.xlsm)' },
  { value: 'xltm', label: 'Excel Macro-Enabled Template (*.xltm)' },
];

function ExportPage({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [kind, setKind] = useState<'excel' | 'csv' | 'pdf'>('excel');
  const [type, setType] = useState<ExportFormat>('xlsx');
  const [name, setName] = useState(baseName(ui.doc.fileName));
  const run = (format: ExportFormat) => {
    onClose();
    if (format === 'pdf') ui.openBackstage('print');
    else ui.files.exportAs(format, name.trim() || 'Book1');
  };
  return (
    <>
      <h1>Export</h1>
      <div className="sh-bs-columns">
        <div className="sh-bs-list" role="menu" aria-label="Export">
          {(['excel', 'csv', 'pdf'] as const).map((k) => (
            <button key={k} type="button" role="menuitemradio" aria-checked={kind === k} className={kind === k ? 'active' : ''} data-testid={`export-${k}`} onClick={() => setKind(k)}>
              {k === 'excel' ? 'Excel File' : k === 'csv' ? 'CSV File' : 'PDF File'}
            </button>
          ))}
        </div>
        <div className="sh-bs-detail">
          <h2>{kind === 'excel' ? 'Excel File' : kind === 'csv' ? 'CSV File' : 'PDF File'}</h2>
          {kind !== 'pdf' && (
            <label className="sh-field">
              <span>File name</span>
              <input className="sh-input" value={name} data-testid="export-name" onChange={(e) => setName(e.target.value)} />
            </label>
          )}
          {kind === 'excel' && (
            <label className="sh-field">
              <span>File type</span>
              <select className="sh-select" value={type} data-testid="export-type" onChange={(e) => setType(e.target.value as ExportFormat)}>
                {EXCEL_TYPES.map((t) => <option key={t.value} value={t.value}>{t.label}</option>)}
              </select>
            </label>
          )}
          <p className="sh-bs-lead">
            {kind === 'excel' && 'Saves a copy of the workbook. Parts of an opened file that the editor does not change are kept as they were.'}
            {kind === 'csv' && 'Saves the values of the active sheet as comma separated text (UTF-8).'}
            {kind === 'pdf' && 'Creates a PDF from the print settings and the page setup of the sheet (File > Print shows the pages first).'}
          </p>
          <button type="button" className="sh-bs-action" data-testid="export-file" onClick={() => run(kind === 'excel' ? type : kind)}>
            <Icon name={kind === 'pdf' ? 'picture_as_pdf' : 'file_download'} />
            <span>{kind === 'excel' ? 'Export Excel File' : kind === 'csv' ? 'Export CSV File' : 'Export PDF File'}</span>
          </button>
        </div>
      </div>
    </>
  );
}

// -----------------------------------------------------------------------------------------------
// Print

const PT_TO_PX = 96 / 72;

function PrintPage({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const sheet = doc.sheet;
  const page = sheet.page;
  const [choice, setChoice] = useState<PrintChoice>({ scope: 'sheet' });
  const [pages, setPages] = useState<Page[] | null>(null);
  const [current, setCurrent] = useState(0);
  const [busy, setBusy] = useState(false);
  const [fit, setFit] = useState(true);
  const measure = useRef<Measure | null>(null);
  const pane = useRef<HTMLDivElement>(null);
  const [paneSize, setPaneSize] = useState({ w: 600, h: 700 });
  const title = baseName(doc.fileName);
  const version = ui.state.version;

  const opts = useMemo(() => printOptionsFor(doc, sheet, choice), [doc, sheet, choice, version]);
  useEffect(() => {
    let alive = true;
    const t = setTimeout(async () => {
      await loadPrintFonts();
      if (!alive) return;
      measure.current ??= canvasMeasure();
      const built = buildPages(doc, opts, measure.current, { title: doc.fileName });
      setPages(built);
      setCurrent((c) => Math.min(c, built.length - 1));
    }, 30);
    return () => {
      alive = false;
      clearTimeout(t);
    };
  }, [doc, opts]);

  useEffect(() => {
    const el = pane.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setPaneSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setPaneSize({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  const pg = pages?.[current];
  const scale = pg ? (fit ? Math.min((paneSize.w - 40) / (pg.width * PT_TO_PX), (paneSize.h - 40) / (pg.height * PT_TO_PX)) : 1) : 1;
  const markup = useMemo(() => (pg ? pageSvg(pg, { width: '100%', height: '100%' }) : ''), [pg]);

  const download = async () => {
    if (!pages || busy) return;
    setBusy(true);
    try {
      const bytes = await pagesToPdf(pages, { rasterize, measure: measure.current ?? canvasMeasure() }, { title });
      saveFile(bytes, `${title}.pdf`, 'application/pdf');
    } catch (e) {
      ui.notify(`The PDF could not be created: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  const scaling = scalingKind(page.fitToPage, page.fitToWidth, page.fitToHeight);
  const marginPreset = MARGIN_PRESETS.find((m) => JSON.stringify(m.margins) === JSON.stringify(page.margins))?.name ?? 'custom';

  return (
    <div className="sh-print">
      <div className="sh-print-settings" data-testid="print-options">
        <h1>Print</h1>
        <div className="sh-print-actions">
          <button type="button" className="sh-bs-action big" data-testid="print-print" disabled={!pages} onClick={() => pages && printPages(pages)}>
            <Icon name="print" />
            <span>Print</span>
          </button>
          <button type="button" className="sh-bs-action big" data-testid="print-download" disabled={!pages || busy} onClick={() => void download()}>
            <Icon name={busy ? 'hourglass_empty' : 'picture_as_pdf'} />
            <span>{busy ? 'Creating…' : 'Save as PDF'}</span>
          </button>
        </div>
        <h2>Settings</h2>
        <PrintSelect label="What to print" icon="grid_on" testId="print-scope" value={choice.scope} onChange={(v) => setChoice({ ...choice, scope: v as PrintChoice['scope'] })} options={[
          { value: 'sheet', label: 'Print Active Sheets', hint: 'Only print the active sheets' },
          { value: 'workbook', label: 'Print Entire Workbook', hint: 'Print the entire workbook' },
          { value: 'selection', label: 'Print Selection', hint: 'Only print the current selection' },
        ]} />
        <div className="sh-print-pages">
          <span>Pages:</span>
          <input type="number" min={1} className="sh-input" aria-label="Pages from" data-testid="print-from" value={choice.pageFrom ?? ''} onChange={(e) => setChoice({ ...choice, pageFrom: e.target.value ? Number(e.target.value) : undefined })} />
          <span>to</span>
          <input type="number" min={1} className="sh-input" aria-label="Pages to" data-testid="print-to" value={choice.pageTo ?? ''} onChange={(e) => setChoice({ ...choice, pageTo: e.target.value ? Number(e.target.value) : undefined })} />
        </div>
        <PrintSelect label="Orientation" icon={page.orientation === 'landscape' ? 'crop_landscape' : 'crop_portrait'} testId="print-orientation" value={page.orientation} onChange={(v) => doc.setPageSetup({ orientation: v as 'portrait' | 'landscape' })} options={[
          { value: 'portrait', label: 'Portrait Orientation' },
          { value: 'landscape', label: 'Landscape Orientation' },
        ]} />
        <PrintSelect label="Paper size" icon="description" testId="print-paper" value={String(page.paperSize)} onChange={(v) => doc.setPageSetup({ paperSize: Number(v) })} options={PAPER.map((p) => ({ value: String(p.code), label: p.label, hint: p.size }))} />
        <PrintSelect label="Margins" icon="margin" testId="print-margins" value={marginPreset} onChange={(v) => {
          if (v === 'custom') ui.openDialog({ kind: 'page-setup', tab: 'margins' });
          else doc.setPageSetup({ margins: { ...MARGIN_PRESETS.find((m) => m.name === v)!.margins } });
        }} options={[
          ...MARGIN_PRESETS.map((m) => ({ value: m.name, label: `${m.label} Margins`, hint: `Top: ${m.margins.top}" Bottom: ${m.margins.bottom}" Left: ${m.margins.left}" Right: ${m.margins.right}"` })),
          { value: 'custom', label: 'Custom Margins...', hint: marginPreset === 'custom' ? `Top: ${page.margins.top}" Left: ${page.margins.left}"` : '' },
        ]} />
        <PrintSelect label="Scaling" icon="fit_screen" testId="print-scaling" value={scaling} onChange={(v) => {
          if (v === 'none') doc.setPageSetup({ fitToPage: false, scale: 100 });
          else if (v === 'sheet') doc.setPageSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 1 });
          else if (v === 'columns') doc.setPageSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 0 });
          else if (v === 'rows') doc.setPageSetup({ fitToPage: true, fitToWidth: 0, fitToHeight: 1 });
          else ui.openDialog({ kind: 'page-setup', tab: 'page' });
        }} options={[
          { value: 'none', label: page.fitToPage || page.scale === 100 ? 'No Scaling' : `Custom Scaling (${page.scale}%)`, hint: 'Print sheets at their actual size' },
          { value: 'sheet', label: 'Fit Sheet on One Page', hint: 'Shrink the printout so that it fits on one page' },
          { value: 'columns', label: 'Fit All Columns on One Page', hint: 'Shrink the printout so that it is one page wide' },
          { value: 'rows', label: 'Fit All Rows on One Page', hint: 'Shrink the printout so that it is one page high' },
          { value: 'custom', label: 'Custom Scaling Options...' },
        ]} />
        <button type="button" className="sh-link" data-testid="print-page-setup" onClick={() => ui.openDialog({ kind: 'page-setup', tab: 'page' })}>Page Setup</button>
        <div className="sh-hint" style={{ marginTop: 12 }}>{paperOf(page.paperSize).label} · {pages?.length ?? 0} page{pages?.length === 1 ? '' : 's'}</div>
        <button type="button" className="sh-link" onClick={onClose}>Back to the sheet</button>
      </div>
      <div className="sh-print-preview">
        <div className="sh-print-canvas" ref={pane} data-testid="print-preview">
          {!pages && <div className="sh-print-wait">Preparing pages…</div>}
          {pg && (
            <div
              className="sh-print-page"
              data-testid="print-page"
              style={{ width: pg.width * PT_TO_PX * scale, height: pg.height * PT_TO_PX * scale }}
              // Generated from the document by the print layout (escaped text).
              dangerouslySetInnerHTML={{ __html: markup }}
            />
          )}
        </div>
        <div className="sh-print-nav">
          <button type="button" aria-label="Previous page" data-testid="print-prev" disabled={current <= 0} onClick={() => setCurrent(current - 1)}><Icon name="arrow_left" /></button>
          <input className="sh-input" aria-label="Current page" value={pages ? current + 1 : 0} onChange={(e) => { const n = Number(e.target.value); if (pages && n >= 1 && n <= pages.length) setCurrent(n - 1); }} />
          <span data-testid="print-page-count">of {pages?.length ?? 0}</span>
          <button type="button" aria-label="Next page" data-testid="print-next" disabled={!pages || current >= pages.length - 1} onClick={() => setCurrent(current + 1)}><Icon name="arrow_right" /></button>
          <span className="grow" />
          <button type="button" aria-label={fit ? 'Actual size' : 'Zoom to page'} title={fit ? 'Actual size' : 'Zoom to page'} className={fit ? 'on' : ''} onClick={() => setFit(!fit)}><Icon name="fit_screen" /></button>
        </div>
      </div>
    </div>
  );
}

function PrintSelect({ label, icon, value, options, onChange, testId }: { label: string; icon: IconName; value: string; options: { value: string; label: string; hint?: string }[]; onChange: (v: string) => void; testId: string }) {
  const sel = options.find((o) => o.value === value) ?? options[0];
  return (
    <label className="sh-print-select" title={label}>
      <span className="sr-only">{label}</span>
      <span className="face" aria-hidden="true">
        <span className="ic"><Icon name={icon} /></span>
        <span className="txt">
          <span className="t">{sel.label}</span>
          {sel.hint && <span className="h">{sel.hint}</span>}
        </span>
      </span>
      <select aria-label={label} data-testid={testId} value={value} onChange={(e) => onChange(e.target.value)}>
        {options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    </label>
  );
}

// -----------------------------------------------------------------------------------------------
// Info

function InfoPage() {
  const ui = useUI();
  const { doc } = ui;
  const [, force] = useState(0);
  const date = (iso?: string) => (iso ? new Date(iso).toLocaleString() : '');
  return (
    <>
      <h1>Info</h1>
      <div className="sh-info">
        <div className="sh-info-props">
          <h2>Properties</h2>
          {PROPS.map((p) => (
            <label key={p.key} className="sh-info-row">
              <span>{p.label}</span>
              <input
                className="sh-input"
                placeholder={p.placeholder}
                data-testid={`info-${p.key}`}
                defaultValue={doc.wb.props[p.key] ?? ''}
                onBlur={(e) => {
                  const v = e.target.value.trim();
                  if ((doc.wb.props[p.key] ?? '') === v) return;
                  if (v) doc.wb.props[p.key] = v;
                  else delete doc.wb.props[p.key];
                  doc.modified = true;
                  force((n) => n + 1);
                }}
              />
            </label>
          ))}
        </div>
        <div className="sh-info-side">
          <h2>Related Dates</h2>
          <div className="sh-info-row"><span>Last Modified</span><span>{date(doc.wb.props.modified)}</span></div>
          <div className="sh-info-row"><span>Created</span><span>{date(doc.wb.props.created)}</span></div>
          <h2>Related People</h2>
          <div className="sh-info-row"><span>Last Modified By</span><span>{doc.wb.props.lastModifiedBy ?? ''}</span></div>
          <h2>File</h2>
          <div className="sh-info-row"><span>Name</span><span data-testid="file-name">{doc.fileName}</span></div>
          <div className="sh-info-row"><span>Sheets</span><span>{doc.wb.sheets.length}</span></div>
          <h2>Protection</h2>
          <button type="button" className="sh-btn-secondary" data-testid="info-protect-workbook" onClick={() => ui.openDialog({ kind: 'protect-workbook' })}>
            {doc.wb.protection?.structure ? 'Unprotect Workbook...' : 'Protect Workbook...'}
          </button>
        </div>
      </div>
    </>
  );
}
