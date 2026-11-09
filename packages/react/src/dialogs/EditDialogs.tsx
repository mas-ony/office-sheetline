/**
 * Editing dialogs: Find and Replace (with Go To), Go To Special, Zoom, Insert / Delete cells,
 * row and column sizes, Paste Special and Insert Hyperlink.
 */
import { useEffect, useMemo, useState } from 'react';
import { useUI } from '../context';
import { Check, Dialog, Field, ListBox, NumberField, Radio, Tabs } from './Dialog';
import { cellName, cmd, GOTO_SPECIALS as SPECIALS, PASTE_SPECIAL_WHAT as PASTE_WHAT, type PasteOptions, type PasteWhat, ptToPx, pxToPt, type SearchHit, type SpecialKind, widthCharsToPx, widthPxToChars } from '@sheetline/core/views';

// -----------------------------------------------------------------------------------------------
// Find and Replace / Go To

export function FindDialog({ tab: initial = 'find', onClose }: { tab?: 'find' | 'replace' | 'goto'; onClose: () => void }) {
  const ui = useUI();
  const { doc, grid } = ui;
  const [tab, setTab] = useState(initial);
  const [query, setQuery] = useState(() => {
    const r = doc.range;
    return r.r0 === r.r1 && r.c0 === r.c1 ? '' : '';
  });
  const [replacement, setReplacement] = useState('');
  const [options, setOptions] = useState(false);
  const [within, setWithin] = useState<'sheet' | 'workbook'>('sheet');
  const [byCols, setByCols] = useState(false);
  const [lookIn, setLookIn] = useState<'formulas' | 'values'>('formulas');
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [hits, setHits] = useState<SearchHit[] | null>(null);
  const [index, setIndex] = useState(-1);
  const [all, setAll] = useState<SearchHit[] | null>(null);
  const [goto, setGoto] = useState('');
  const opts = useMemo(() => ({ matchCase, wholeCell: whole, inFormulas: lookIn === 'formulas', allSheets: within === 'workbook' }), [matchCase, whole, lookIn, within]);

  useEffect(() => {
    setHits(null);
    setIndex(-1);
  }, [query, opts, byCols]);

  useEffect(() => () => {
    if (grid) {
      grid.searchHit = undefined;
      grid.requestRender();
    }
  }, [grid]);

  const search = () => {
    let list = doc.find(query, opts);
    if (byCols) list = [...list].sort((a, b) => doc.wb.sheets.indexOf(a.sheet) - doc.wb.sheets.indexOf(b.sheet) || a.c - b.c || a.r - b.r);
    setHits(list);
    return list;
  };

  const show = (hit: SearchHit) => {
    const si = doc.wb.sheets.indexOf(hit.sheet);
    if (si !== doc.activeIndex) doc.setActiveSheet(si);
    doc.selectCell(hit.r, hit.c);
    grid?.scrollIntoView(hit.r, hit.c);
    if (grid) {
      grid.searchHit = { r: hit.r, c: hit.c };
      grid.requestRender();
    }
  };

  const next = () => {
    if (!query) return;
    const list = hits ?? search();
    if (!list.length) {
      ui.notify(`We couldn't find what you were looking for.`);
      return;
    }
    // Start after the active cell like Excel.
    let k = index + 1;
    if (index < 0) {
      const { r, c } = doc.selection.active;
      const pos = list.findIndex((h) => h.sheet === doc.sheet && (byCols ? h.c > c || (h.c === c && h.r > r) : h.r > r || (h.r === r && h.c > c)));
      k = pos < 0 ? 0 : pos;
    }
    k %= list.length;
    setIndex(k);
    show(list[k]);
  };

  const replaceOne = () => {
    const list = hits ?? search();
    const cur = list[index];
    const { r, c } = doc.selection.active;
    if (cur && cur.sheet === doc.sheet && cur.r === r && cur.c === c) {
      doc.replace(cur, query, replacement, opts);
      const fresh = search();
      setIndex(index - 1);
      if (fresh.length) {
        const k = Math.min(Math.max(0, index), fresh.length - 1);
        setIndex(k);
        show(fresh[k]);
      }
      return;
    }
    next();
  };

  const replaceAll = () => {
    cmd.commit(ui);
    const n = doc.replaceAll(query, replacement, opts);
    setHits(null);
    ui.notify(n ? `All done. We made ${n} replacement${n > 1 ? 's' : ''}.` : `We couldn't find anything to replace.`);
  };

  const names = doc.wb.names.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.')).map((n) => n.name);

  return (
    <Dialog title={tab === 'goto' ? 'Go To' : 'Find and Replace'} onClose={onClose} closeOnly blocking={false} width={520} testId="find-dialog">
      <Tabs value={tab} onChange={setTab} tabs={[{ id: 'find', label: 'Find' }, { id: 'replace', label: 'Replace' }, { id: 'goto', label: 'Go To' }]} testId="find-tabs" />
      <div className="sh-tab-body">
        {tab !== 'goto' ? (
          <div className="sh-stack">
            <Field label="Find what:" wide>
              <input
                value={query}
                autoFocus
                data-testid="find-input"
                onChange={(e) => setQuery(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    next();
                  }
                }}
              />
            </Field>
            {tab === 'replace' && (
              <Field label="Replace with:" wide>
                <input value={replacement} data-testid="replace-input" onChange={(e) => setReplacement(e.target.value)} />
              </Field>
            )}
            {options && (
              <div className="sh-row wrap top">
                <div className="sh-stack">
                  <Field label="Within:">
                    <select value={within} onChange={(e) => setWithin(e.target.value as typeof within)}>
                      <option value="sheet">Sheet</option>
                      <option value="workbook">Workbook</option>
                    </select>
                  </Field>
                  <Field label="Search:">
                    <select value={byCols ? 'cols' : 'rows'} onChange={(e) => setByCols(e.target.value === 'cols')}>
                      <option value="rows">By Rows</option>
                      <option value="cols">By Columns</option>
                    </select>
                  </Field>
                  <Field label="Look in:">
                    <select value={lookIn} onChange={(e) => setLookIn(e.target.value as typeof lookIn)}>
                      <option value="formulas">Formulas</option>
                      <option value="values">Values</option>
                    </select>
                  </Field>
                </div>
                <div className="sh-stack">
                  <Check label="Match case" checked={matchCase} onChange={setMatchCase} testId="find-match-case" />
                  <Check label="Match entire cell contents" checked={whole} onChange={setWhole} testId="find-whole" />
                </div>
              </div>
            )}
            <div className="sh-row wrap">
              <button type="button" className="sh-btn" onClick={() => setOptions(!options)}>{options ? 'Options <<' : 'Options >>'}</button>
              <span className="grow" />
              {tab === 'replace' && (
                <>
                  <button type="button" className="sh-btn" data-testid="replace-all" disabled={!query} onClick={replaceAll}>Replace All</button>
                  <button type="button" className="sh-btn" data-testid="replace-one" disabled={!query} onClick={replaceOne}>Replace</button>
                </>
              )}
              <button type="button" className="sh-btn" data-testid="find-all" disabled={!query} onClick={() => setAll(search())}>Find All</button>
              <button type="button" className="sh-btn primary" data-testid="find-next" disabled={!query} onClick={next}>Find Next</button>
            </div>
            {all && (
              <div className="sh-table finds" role="table" data-testid="find-results">
                <div className="sh-table-head" role="row"><span>Sheet</span><span>Cell</span><span>Value</span></div>
                <div className="sh-table-body">
                  {all.map((h, i) => (
                    <div key={i} role="row" className="sh-table-row" onClick={() => show(h)}>
                      <span>{h.sheet.name}</span>
                      <span>{cellName(h.r, h.c, true, true)}</span>
                      <span>{h.text}</span>
                    </div>
                  ))}
                </div>
                <div className="sh-hint">{all.length} cell(s) found</div>
              </div>
            )}
          </div>
        ) : (
          <div className="sh-stack">
            <div className="sh-label">Go to:</div>
            <ListBox items={names} value={names.includes(goto) ? goto : null} keyOf={(x) => x} render={(x) => x} height={140} label="Names" onChange={setGoto} onActivate={(x) => { if (doc.goTo(x)) onClose(); }} />
            <Field label="Reference:" wide>
              <input
                value={goto}
                data-testid="goto-input"
                onChange={(e) => setGoto(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    if (doc.goTo(goto)) onClose();
                    else ui.notify('Reference isn\'t valid.');
                  }
                }}
              />
            </Field>
            <div className="sh-row">
              <button type="button" className="sh-btn" onClick={() => ui.openDialog({ kind: 'goto-special' })}>Special...</button>
              <span className="grow" />
              <button type="button" className="sh-btn primary" data-testid="goto-ok" onClick={() => { if (doc.goTo(goto)) { grid?.scrollIntoView(doc.selection.active.r, doc.selection.active.c); onClose(); } else ui.notify('Reference isn\'t valid.'); }}>OK</button>
            </div>
          </div>
        )}
      </div>
    </Dialog>
  );
}

export function GotoSpecialDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [kind, setKind] = useState<SpecialKind>('comments');
  const [types, setTypes] = useState({ numbers: true, text: true, logicals: true, errors: true });
  return (
    <Dialog
      title="Go To Special"
      onClose={onClose}
      width={420}
      testId="goto-special"
      onOk={() => {
        const n = ui.doc.selectSpecial(kind, types);
        if (!n) {
          ui.notify('No cells were found.');
          return false;
        }
        return true;
      }}
    >
      <div className="sh-special-grid">
        {SPECIALS.map((s) => (
          <div key={s.kind}>
            <Radio name="goto-special" label={s.label} checked={kind === s.kind} onChange={() => setKind(s.kind)} testId={`special-${s.kind}`} />
            {(s.kind === 'formulas' || s.kind === 'constants') && kind === s.kind && (
              <div className="sh-indent">
                <Check label="Numbers" checked={types.numbers} onChange={(v) => setTypes({ ...types, numbers: v })} />
                <Check label="Text" checked={types.text} onChange={(v) => setTypes({ ...types, text: v })} />
                <Check label="Logicals" checked={types.logicals} onChange={(v) => setTypes({ ...types, logicals: v })} />
                <Check label="Errors" checked={types.errors} onChange={(v) => setTypes({ ...types, errors: v })} />
              </div>
            )}
          </div>
        ))}
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Zoom

export function ZoomDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const current = Math.round(ui.doc.sheet.zoom * 100);
  const presets = [200, 100, 75, 50, 25];
  const [choice, setChoice] = useState<number | 'fit' | 'custom'>(presets.includes(current) ? current : 'custom');
  const [custom, setCustom] = useState<number | ''>(current);
  return (
    <Dialog
      title="Zoom"
      onClose={onClose}
      width={260}
      testId="zoom-dialog"
      onOk={() => {
        if (choice === 'fit') cmd.zoomToSelection(ui);
        else {
          const z = choice === 'custom' ? Number(custom) : choice;
          if (!Number.isFinite(z) || z < 10 || z > 400) {
            ui.notify('Enter a whole number between 10 and 400.');
            return false;
          }
          cmd.zoomTo(ui, z / 100);
        }
        return true;
      }}
    >
      <fieldset className="sh-group">
        <legend>Magnification</legend>
        {presets.map((p) => <Radio key={p} name="zoom" label={`${p}%`} checked={choice === p} onChange={() => setChoice(p)} testId={`zoom-${p}`} />)}
        <Radio name="zoom" label="Fit selection" checked={choice === 'fit'} onChange={() => setChoice('fit')} />
        <div className="sh-row">
          <Radio name="zoom" label="Custom:" checked={choice === 'custom'} onChange={() => setChoice('custom')} />
          <NumberField value={custom} min={10} max={400} label="Custom zoom" testId="zoom-custom" onChange={(v) => { setCustom(v); setChoice('custom'); }} />
          <span>%</span>
        </div>
      </fieldset>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Insert / Delete cells

export function InsertCellsDialog({ remove, onClose }: { remove?: boolean; onClose: () => void }) {
  const ui = useUI();
  const [mode, setMode] = useState<'shift' | 'shift2' | 'rows' | 'cols'>('shift');
  return (
    <Dialog
      title={remove ? 'Delete' : 'Insert'}
      onClose={onClose}
      width={260}
      testId={remove ? 'delete-cells' : 'insert-cells'}
      onOk={() => {
        cmd.commit(ui);
        const range = ui.doc.range;
        if (mode === 'rows') {
          if (remove) cmd.deleteRows(ui);
          else cmd.insertRows(ui);
          return true;
        }
        if (mode === 'cols') {
          if (remove) cmd.deleteCols(ui);
          else cmd.insertCols(ui);
          return true;
        }
        const error = remove ? ui.doc.deleteCells(range, mode === 'shift' ? 'left' : 'up') : ui.doc.insertCells(range, mode === 'shift' ? 'right' : 'down');
        if (error) {
          ui.notify(error);
          return false;
        }
        return true;
      }}
    >
      <fieldset className="sh-group">
        <legend>{remove ? 'Delete' : 'Insert'}</legend>
        <Radio name="cells" label={remove ? 'Shift cells left' : 'Shift cells right'} checked={mode === 'shift'} onChange={() => setMode('shift')} testId="cells-shift-h" />
        <Radio name="cells" label={remove ? 'Shift cells up' : 'Shift cells down'} checked={mode === 'shift2'} onChange={() => setMode('shift2')} testId="cells-shift-v" />
        <Radio name="cells" label="Entire row" checked={mode === 'rows'} onChange={() => setMode('rows')} testId="cells-rows" />
        <Radio name="cells" label="Entire column" checked={mode === 'cols'} onChange={() => setMode('cols')} testId="cells-cols" />
      </fieldset>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Row height, column width, defaults

export function SizeDialog({ what, onClose }: { what: 'row-height' | 'col-width' | 'default-height' | 'default-width'; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const sheet = doc.sheet;
  const { r, c } = doc.selection.active;
  const initial =
    what === 'row-height' ? pxToPt(sheet.rowHeightPx(r)) : what === 'default-height' ? sheet.defaultRowHeightPt : what === 'col-width' ? widthPxToChars(sheet.colWidthPx(c)) : widthPxToChars(sheet.defaultColWidthPx);
  const [value, setValue] = useState(String(Math.round(initial * 100) / 100));
  const title = { 'row-height': 'Row Height', 'col-width': 'Column Width', 'default-height': 'Standard Height', 'default-width': 'Standard Width' }[what];
  return (
    <Dialog
      title={title}
      onClose={onClose}
      width={280}
      testId="size-dialog"
      onOk={() => {
        const v = Number(value);
        const max = what.includes('height') ? 409 : 255;
        if (!Number.isFinite(v) || v < 0 || v > max) {
          ui.notify(`Enter a number between 0 and ${max}.`);
          return false;
        }
        cmd.commit(ui);
        const lines = (axis: 'row' | 'col') => {
          const out = new Set<number>();
          for (const x of doc.selection.ranges) {
            const [a, b] = axis === 'row' ? [x.r0, Math.min(x.r1, x.r0 + 100000)] : [x.c0, x.c1];
            for (let i = a; i <= b; i++) out.add(i);
          }
          return [...out];
        };
        if (what === 'row-height') doc.setRowHeight(lines('row'), v);
        else if (what === 'col-width') doc.setColWidth(lines('col'), widthCharsToPx(v));
        else if (what === 'default-height') doc.setDefaultSize({ rowHeightPt: v });
        else doc.setDefaultSize({ colWidthPx: widthCharsToPx(v) });
        return true;
      }}
    >
      <Field label={`${title}:`}>
        <input value={value} inputMode="decimal" autoFocus data-testid="size-input" style={{ width: 90 }} onChange={(e) => setValue(e.target.value)} />
      </Field>
      {what === 'row-height' && <p className="sh-hint">Points ({Math.round(ptToPx(Number(value) || 0))} pixels).</p>}
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Paste Special

export function PasteSpecialDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [what, setWhat] = useState<PasteWhat>('all');
  const [op, setOp] = useState<NonNullable<PasteOptions['op']>>('none');
  const [skipBlanks, setSkipBlanks] = useState(false);
  const [transpose, setTranspose] = useState(false);
  const has = ui.doc.hasClip();
  const paste = (link = false) => {
    void cmd.paste(ui, { what, op, skipBlanks, transpose, link });
  };
  return (
    <Dialog
      title="Paste Special"
      onClose={onClose}
      width={460}
      testId="paste-special"
      okDisabled={!has}
      onOk={() => paste()}
      footerLeft={<button type="button" className="sh-btn" disabled={!has} onClick={() => { paste(true); onClose(); }}>Paste Link</button>}
    >
      {!has && <p className="sh-hint">Copy cells in this workbook first (Ctrl+C).</p>}
      <div className="sh-row top wrap">
        <fieldset className="sh-group">
          <legend>Paste</legend>
          <div className="sh-two-col even">
            {PASTE_WHAT.map((p) => <Radio key={p.what} name="paste-what" label={p.label} checked={what === p.what} onChange={() => setWhat(p.what)} testId={`paste-what-${p.what}`} />)}
          </div>
        </fieldset>
      </div>
      <fieldset className="sh-group">
        <legend>Operation</legend>
        <div className="sh-two-col even">
          {(['none', 'add', 'subtract', 'multiply', 'divide'] as const).map((o) => (
            <Radio key={o} name="paste-op" label={o === 'none' ? 'None' : o[0].toUpperCase() + o.slice(1)} checked={op === o} onChange={() => setOp(o)} testId={`paste-op-${o}`} />
          ))}
        </div>
      </fieldset>
      <div className="sh-row">
        <Check label="Skip blanks" checked={skipBlanks} onChange={setSkipBlanks} />
        <Check label="Transpose" checked={transpose} onChange={setTranspose} testId="paste-transpose" />
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Hyperlink

export function HyperlinkDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const { r, c } = doc.selection.active;
  const link = doc.sheet.hyperlinkAt(r, c);
  const [tab, setTab] = useState<'web' | 'doc' | 'email'>(link?.location ? 'doc' : link?.target?.startsWith('mailto:') ? 'email' : 'web');
  const [text, setText] = useState(doc.displayText(r, c) || link?.display || '');
  const [tip, setTip] = useState(link?.tooltip ?? '');
  const [url, setUrl] = useState(link?.target && !link.target.startsWith('mailto:') ? link.target : '');
  const [sheetName, setSheetName] = useState(link?.location?.includes('!') ? link.location.split('!')[0].replace(/^'|'$/g, '') : doc.sheet.name);
  const [ref, setRef] = useState(link?.location?.includes('!') ? link.location.split('!')[1] : 'A1');
  const [email, setEmail] = useState(link?.target?.startsWith('mailto:') ? decodeURIComponent(link.target.slice(7).split('?')[0]) : '');
  const [subject, setSubject] = useState(link?.target?.includes('subject=') ? decodeURIComponent(link.target.split('subject=')[1]) : '');
  return (
    <Dialog
      title="Insert Hyperlink"
      onClose={onClose}
      width={480}
      testId="hyperlink-dialog"
      footerLeft={link ? <button type="button" className="sh-btn" data-testid="hyperlink-remove" onClick={() => { doc.removeLink(); onClose(); }}>Remove</button> : undefined}
      onOk={() => {
        let target: string;
        if (tab === 'web') {
          if (!url.trim()) {
            ui.notify('Enter an address.');
            return false;
          }
          target = /^[a-z][a-z0-9+.-]*:/i.test(url.trim()) ? url.trim() : `http://${url.trim()}`;
        } else if (tab === 'doc') {
          const q = /[^\w]/.test(sheetName) ? `'${sheetName.replace(/'/g, "''")}'` : sheetName;
          target = `#${q}!${ref.trim() || 'A1'}`;
        } else {
          if (!email.trim()) {
            ui.notify('Enter an e-mail address.');
            return false;
          }
          target = `mailto:${email.trim()}${subject ? `?subject=${encodeURIComponent(subject)}` : ''}`;
        }
        cmd.commit(ui);
        const current = doc.displayText(r, c);
        doc.setLink(target, text && text !== current ? text : current ? undefined : text || target.replace(/^#/, ''), r, c, doc.sheet, tip || undefined);
        return true;
      }}
    >
      <Field label="Text to display:" wide>
        <input value={text} data-testid="hyperlink-text" onChange={(e) => setText(e.target.value)} />
      </Field>
      <Field label="ScreenTip:" wide>
        <input value={tip} onChange={(e) => setTip(e.target.value)} />
      </Field>
      <Tabs value={tab} onChange={setTab} tabs={[{ id: 'web', label: 'Web Page' }, { id: 'doc', label: 'This Document' }, { id: 'email', label: 'Email Address' }]} testId="hyperlink-tabs" />
      <div className="sh-tab-body">
        {tab === 'web' && (
          <Field label="Address:" wide>
            <input value={url} placeholder="URL, like http://www.example.com" inputMode="url" data-testid="hyperlink-url" onChange={(e) => setUrl(e.target.value)} />
          </Field>
        )}
        {tab === 'doc' && (
          <div className="sh-stack">
            <Field label="Type the cell reference:" wide>
              <input value={ref} data-testid="hyperlink-ref" onChange={(e) => setRef(e.target.value)} />
            </Field>
            <div className="sh-label">Or select a place in this document:</div>
            <ListBox items={doc.wb.sheets.filter((s) => s.state === 'visible').map((s) => s.name)} value={sheetName} keyOf={(x) => x} render={(x) => x} height={120} label="Sheets" onChange={setSheetName} />
          </div>
        )}
        {tab === 'email' && (
          <div className="sh-stack">
            <Field label="E-mail address:" wide>
              <input value={email} inputMode="email" data-testid="hyperlink-email" onChange={(e) => setEmail(e.target.value)} />
            </Field>
            <Field label="Subject:" wide>
              <input value={subject} onChange={(e) => setSubject(e.target.value)} />
            </Field>
          </div>
        )}
      </div>
    </Dialog>
  );
}

