/**
 * Data dialogs: Sort, Text to Columns, Data Validation, Remove Duplicates, Group / Ungroup,
 * Subtotal, Outline settings, Goal Seek, Data Table, Scenario Manager and Fill Series.
 */
import { useMemo, useState } from 'react';
import { useUI } from '../context';
import { Check, Dialog, Field, ListBox, Radio, RangeField, Tabs, parseRangeText, rangeText } from './Dialog';
import { cellName, cellsOfRangeText, colName, commit, CUSTOM_LISTS, dataRegion, type DataValidation, looksLikeHeader, makeRange, parseCellRef, parseScenarioCells, type Scenario, scenarioCellsText, type SeriesOptions, setSelectionValidation, type SortLevel, sortLineName, sortLineNumeric, sortLines, splitText, SUBTOTAL_FUNCTIONS, type TextToColumnsOptions, VALIDATION_ALLOW_OPTIONS as ALLOW, VALIDATION_OPERATOR_OPTIONS as OPERATORS, validationAt, validationFormula, type ValidationOperator } from '@sheetline/core/views';

// -----------------------------------------------------------------------------------------------
// Sort

let levelKey = 0;

export function SortDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const region = useMemo(() => dataRegion(ui) ?? doc.range, []); // eslint-disable-line react-hooks/exhaustive-deps
  const [headers, setHeaders] = useState(() => looksLikeHeader(ui, region, Math.max(region.c0, Math.min(region.c1, doc.selection.active.c))));
  const [byColumns, setByColumns] = useState(false);
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [options, setOptions] = useState(false);
  const [levels, setLevels] = useState<(SortLevel & { key: number })[]>([{ key: levelKey++, index: Math.max(region.c0, Math.min(region.c1, doc.selection.active.c)), order: 'asc', sortOn: 'values' }]);
  const [sel, setSel] = useState(0);
  const lines = sortLines(region, byColumns);
  const lineName = (i: number) => sortLineName(doc, region, byColumns, headers, i);
  const numeric = (i: number) => sortLineNumeric(doc, region, byColumns, headers, i);
  const set = (k: number, patch: Partial<SortLevel>) => setLevels(levels.map((l) => (l.key === k ? { ...l, ...patch } : l)));
  return (
    <>
      <Dialog
        title="Sort"
        onClose={onClose}
        width={640}
        testId="sort-dialog"
        onOk={() => {
          commit(ui);
          doc.sortLevels(region, levels, { hasHeader: headers && !byColumns, byColumns, caseSensitive });
          return true;
        }}
      >
        <div className="sh-row wrap sh-toolbar-row">
          <button type="button" className="sh-btn" data-testid="sort-add-level" onClick={() => { const unused = lines.find((i) => !levels.some((l) => l.index === i)) ?? lines[0]; setLevels([...levels, { key: levelKey++, index: unused, order: 'asc', sortOn: 'values' }]); setSel(levels.length); }}>+ Add Level</button>
          <button type="button" className="sh-btn" disabled={levels.length <= 1} onClick={() => { setLevels(levels.filter((_, i) => i !== sel)); setSel(Math.max(0, sel - 1)); }}>✕ Delete Level</button>
          <button type="button" className="sh-btn" onClick={() => { const l = levels[sel]; if (l) { setLevels([...levels.slice(0, sel + 1), { ...l, key: levelKey++ }, ...levels.slice(sel + 1)]); setSel(sel + 1); } }}>Copy Level</button>
          <button type="button" className="sh-btn icon" aria-label="Move up" disabled={sel <= 0} onClick={() => { const n = levels.slice(); [n[sel - 1], n[sel]] = [n[sel], n[sel - 1]]; setLevels(n); setSel(sel - 1); }}>▲</button>
          <button type="button" className="sh-btn icon" aria-label="Move down" disabled={sel >= levels.length - 1} onClick={() => { const n = levels.slice(); [n[sel + 1], n[sel]] = [n[sel], n[sel + 1]]; setLevels(n); setSel(sel + 1); }}>▼</button>
          <button type="button" className="sh-btn" onClick={() => setOptions(true)}>Options...</button>
          <span className="grow" />
          <Check label="My data has headers" checked={headers} disabled={byColumns} onChange={setHeaders} testId="sort-headers" />
        </div>
        <div className="sh-table sort" role="table">
          <div className="sh-table-head" role="row">
            <span>{byColumns ? 'Row' : 'Column'}</span>
            <span>Sort On</span>
            <span>Order</span>
          </div>
          <div className="sh-table-body">
            {levels.map((l, i) => (
              <div key={l.key} role="row" className={`sh-table-row${i === sel ? ' selected' : ''}`} onClick={() => setSel(i)}>
                <span>
                  <span className="sh-sort-by">{i === 0 ? 'Sort by' : 'Then by'}</span>
                  <select value={l.index} data-testid={`sort-by-${i}`} aria-label="Sort by" onChange={(e) => set(l.key, { index: Number(e.target.value) })}>
                    {lines.map((x) => <option key={x} value={x}>{lineName(x)}</option>)}
                  </select>
                </span>
                <span>
                  <select value={l.sortOn} aria-label="Sort on" onChange={(e) => set(l.key, { sortOn: e.target.value as SortLevel['sortOn'], order: e.target.value === 'values' ? 'asc' : 'asc', color: e.target.value === 'values' ? undefined : '' })}>
                    <option value="values">Cell Values</option>
                    <option value="cellColor">Cell Color</option>
                    <option value="fontColor">Font Color</option>
                  </select>
                </span>
                <span>
                  {l.sortOn === 'values' ? (
                    <select
                      value={l.order === 'custom' ? `custom:${CUSTOM_LISTS.findIndex((x) => x === l.list)}` : l.order}
                      aria-label="Order"
                      data-testid={`sort-order-${i}`}
                      onChange={(e) => {
                        const v = e.target.value;
                        if (v.startsWith('custom:')) set(l.key, { order: 'custom', list: CUSTOM_LISTS[Number(v.slice(7))] });
                        else set(l.key, { order: v as 'asc' | 'desc', list: undefined });
                      }}
                    >
                      <option value="asc">{numeric(l.index) ? 'Smallest to Largest' : 'A to Z'}</option>
                      <option value="desc">{numeric(l.index) ? 'Largest to Smallest' : 'Z to A'}</option>
                      {CUSTOM_LISTS.map((list, k) => (
                        <option key={k} value={`custom:${k}`}>{list.slice(0, 3).join(', ')}…</option>
                      ))}
                    </select>
                  ) : (
                    <span className="sh-row">
                      <input value={l.color ?? ''} placeholder="#RRGGBB" style={{ width: 80 }} aria-label="Color" onChange={(e) => set(l.key, { color: e.target.value.toUpperCase() })} />
                      <select value={l.order} aria-label="Position" onChange={(e) => set(l.key, { order: e.target.value as 'asc' | 'desc' })}>
                        <option value="asc">On Top</option>
                        <option value="desc">On Bottom</option>
                      </select>
                    </span>
                  )}
                </span>
              </div>
            ))}
          </div>
        </div>
      </Dialog>
      {options && (
        <Dialog title="Sort Options" onClose={() => setOptions(false)} width={300} testId="sort-options">
          <Check label="Case sensitive" checked={caseSensitive} onChange={setCaseSensitive} />
          <fieldset className="sh-group">
            <legend>Orientation</legend>
            <Radio name="sort-orientation" label="Sort top to bottom" checked={!byColumns} onChange={() => setByColumns(false)} />
            <Radio name="sort-orientation" label="Sort left to right" checked={byColumns} onChange={() => { setByColumns(true); setLevels(levels.map((l) => ({ ...l, index: region.r0 }))); }} />
          </fieldset>
        </Dialog>
      )}
    </>
  );
}

// -----------------------------------------------------------------------------------------------
// Text to Columns

export function TextToColumnsDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const range = doc.bounded(doc.range, doc.sheet) ?? doc.range;
  const [step, setStep] = useState(1);
  const [opts, setOpts] = useState<TextToColumnsOptions>({ type: 'delimited', delimiters: { tab: true, semicolon: false, comma: false, space: false, other: '' }, consecutive: false, qualifier: '"', breaks: [10] });
  const [dest, setDest] = useState(rangeText(makeRange(range.r0, range.c0), doc.sheet));
  const [formats, setFormats] = useState<('general' | 'text' | 'skip')[]>([]);
  const [col, setCol] = useState(0);
  const sample = useMemo(() => {
    const rows: string[] = [];
    for (let r = range.r0; r <= Math.min(range.r1, range.r0 + 7); r++) rows.push(doc.displayText(r, range.c0));
    return rows;
  }, [doc, range.r0, range.r1, range.c0]);
  const split = sample.map((t) => (t === '' ? [] : splitText(t, opts)));
  const columns = Math.max(1, ...split.map((s) => s.length));
  if (range.c0 !== range.c1) {
    return (
      <Dialog title="Convert Text to Columns Wizard" onClose={onClose} closeOnly width={420} testId="text-to-columns">
        <p>This can convert only one column at a time. The range can be many rows tall but no more than one column wide. Try again by selecting cells in one column only.</p>
      </Dialog>
    );
  }
  const set = (p: Partial<TextToColumnsOptions>) => setOpts({ ...opts, ...p });
  const finish = async () => {
    const at = parseCellRef(dest.replace(/\$/g, '').replace(/^.*!/, ''));
    const full: TextToColumnsOptions = { ...opts, destination: at ? { r: at.r, c: at.c } : undefined, formats };
    if (doc.textToColumnsOverwrites(range, full) && !(await ui.confirm('Do you want to replace the contents of the destination cells?', { title: 'Text to Columns', ok: 'OK' }))) return false;
    commit(ui);
    const error = doc.textToColumns(range, full);
    if (error) {
      ui.notify(error);
      return false;
    }
    return true;
  };
  return (
    <Dialog
      title={`Convert Text to Columns Wizard - Step ${step} of 3`}
      onClose={onClose}
      width={600}
      testId="text-to-columns"
      ok={step < 3 ? 'Next >' : 'Finish'}
      footerLeft={
        <>
          <button type="button" className="sh-btn" disabled={step === 1} onClick={() => setStep(step - 1)}>&lt; Back</button>
          {step < 3 && <button type="button" className="sh-btn" data-testid="text-to-columns-finish" onClick={() => void finish().then((done) => done && onClose())}>Finish</button>}
        </>
      }
      onOk={async () => {
        if (step < 3) {
          setStep(step + 1);
          return false;
        }
        return finish();
      }}
    >
      {step === 1 && (
        <fieldset className="sh-group">
          <legend>Original data type</legend>
          <Radio name="t2c-type" label={<span><b>Delimited</b> - Characters such as commas or tabs separate each field.</span>} checked={opts.type === 'delimited'} onChange={() => set({ type: 'delimited' })} testId="t2c-delimited" />
          <Radio name="t2c-type" label={<span><b>Fixed width</b> - Fields are aligned in columns with spaces between each field.</span>} checked={opts.type === 'fixed'} onChange={() => set({ type: 'fixed' })} testId="t2c-fixed" />
        </fieldset>
      )}
      {step === 2 && opts.type === 'delimited' && (
        <div className="sh-row wrap top">
          <fieldset className="sh-group">
            <legend>Delimiters</legend>
            <Check label="Tab" checked={opts.delimiters.tab} onChange={(v) => set({ delimiters: { ...opts.delimiters, tab: v } })} />
            <Check label="Semicolon" checked={opts.delimiters.semicolon} onChange={(v) => set({ delimiters: { ...opts.delimiters, semicolon: v } })} />
            <Check label="Comma" checked={opts.delimiters.comma} testId="t2c-comma" onChange={(v) => set({ delimiters: { ...opts.delimiters, comma: v } })} />
            <Check label="Space" checked={opts.delimiters.space} testId="t2c-space" onChange={(v) => set({ delimiters: { ...opts.delimiters, space: v } })} />
            <div className="sh-row">
              <Check label="Other:" checked={!!opts.delimiters.other} onChange={(v) => set({ delimiters: { ...opts.delimiters, other: v ? opts.delimiters.other || '|' : '' } })} />
              <input value={opts.delimiters.other} maxLength={1} style={{ width: 32 }} aria-label="Other delimiter" onChange={(e) => set({ delimiters: { ...opts.delimiters, other: e.target.value } })} />
            </div>
          </fieldset>
          <div className="sh-stack">
            <Check label="Treat consecutive delimiters as one" checked={opts.consecutive} onChange={(v) => set({ consecutive: v })} />
            <Field label="Text qualifier:">
              <select value={opts.qualifier} onChange={(e) => set({ qualifier: e.target.value as TextToColumnsOptions['qualifier'] })}>
                <option value={'"'}>"</option>
                <option value="'">'</option>
                <option value="">{'{none}'}</option>
              </select>
            </Field>
          </div>
        </div>
      )}
      {step === 2 && opts.type === 'fixed' && (
        <Field label="Break positions (characters, separated by commas):" wide>
          <input value={opts.breaks.join(', ')} onChange={(e) => set({ breaks: e.target.value.split(/[,\s]+/).map(Number).filter((n) => n > 0) })} />
        </Field>
      )}
      {step === 3 && (
        <div className="sh-row wrap top">
          <fieldset className="sh-group">
            <legend>Column data format</legend>
            <Radio name="t2c-format" label="General" checked={(formats[col] ?? 'general') === 'general'} onChange={() => { const f = formats.slice(); f[col] = 'general'; setFormats(f); }} />
            <Radio name="t2c-format" label="Text" checked={formats[col] === 'text'} onChange={() => { const f = formats.slice(); f[col] = 'text'; setFormats(f); }} />
            <Radio name="t2c-format" label="Do not import column (skip)" checked={formats[col] === 'skip'} onChange={() => { const f = formats.slice(); f[col] = 'skip'; setFormats(f); }} />
          </fieldset>
          <Field label="Destination:">
            <RangeField value={dest} onChange={setDest} label="Destination" testId="t2c-destination" />
          </Field>
        </div>
      )}
      <div className="sh-label">Data preview</div>
      <div className="sh-t2c-preview" data-testid="t2c-preview">
        <table>
          {step === 3 && (
            <thead>
              <tr>
                {Array.from({ length: columns }, (_, i) => (
                  <th key={i} className={i === col ? 'selected' : ''} onClick={() => setCol(i)}>{formats[i] === 'skip' ? 'Skip' : formats[i] === 'text' ? 'Text' : 'General'}</th>
                ))}
              </tr>
            </thead>
          )}
          <tbody>
            {split.map((cells, r) => (
              <tr key={r}>
                {step === 1 ? <td>{sample[r]}</td> : Array.from({ length: columns }, (_, i) => <td key={i} className={step === 3 && i === col ? 'selected' : ''} onClick={() => setCol(i)}>{cells[i] ?? ''}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Data Validation

export function ValidationDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const { r, c } = doc.selection.active;
  const existing = validationAt(doc.sheet, r, c);
  const [tab, setTab] = useState<'settings' | 'input' | 'error'>('settings');
  const [type, setType] = useState(existing?.type ?? 'none');
  const [op, setOp] = useState<ValidationOperator>((existing?.operator as ValidationOperator) ?? 'between');
  const [f1, setF1] = useState(existing?.formula1 ? (existing.type === 'list' && existing.formula1.startsWith('"') ? existing.formula1.slice(1, -1) : `=${existing.formula1}`) : '');
  const [f2, setF2] = useState(existing?.formula2 ? `=${existing.formula2}` : '');
  const [ignoreBlank, setIgnoreBlank] = useState(existing?.allowBlank ?? true);
  const [dropDown, setDropDown] = useState(!existing?.hideDropDown);
  const [showInput, setShowInput] = useState(existing?.showInputMessage ?? true);
  const [promptTitle, setPromptTitle] = useState(existing?.promptTitle ?? '');
  const [prompt, setPrompt] = useState(existing?.prompt ?? '');
  const [showError, setShowError] = useState(existing?.showErrorMessage ?? true);
  const [errorStyle, setErrorStyle] = useState<'stop' | 'warning' | 'information'>(existing?.errorStyle ?? 'stop');
  const [errorTitle, setErrorTitle] = useState(existing?.errorTitle ?? '');
  const [error, setError] = useState(existing?.error ?? '');
  const twoValues = op === 'between' || op === 'notBetween';
  const usesOp = type !== 'none' && type !== 'list' && type !== 'custom';

  const toFormula = (text: string) => validationFormula(doc, r, c, type, text);

  const save = () => {
    const sel = doc.selection.ranges.map((x) => doc.bounded(x, doc.sheet) ?? x);
    let rule: DataValidation | null = null;
    if (type !== 'none' || prompt || error) {
      if (type !== 'none' && type !== 'custom' && !f1.trim()) {
        ui.notify(type === 'list' ? 'Enter the list items or a reference (Source).' : 'Enter a value.');
        return false;
      }
      rule = {
        sqref: sel,
        type,
        operator: usesOp ? op : undefined,
        formula1: type === 'none' ? undefined : type === 'list' && !f1.trim().startsWith('=') ? `"${f1.replace(/"/g, '')}"` : toFormula(f1),
        formula2: usesOp && twoValues ? toFormula(f2) : undefined,
        allowBlank: ignoreBlank,
        hideDropDown: type === 'list' ? !dropDown : undefined,
        showInputMessage: showInput,
        promptTitle: promptTitle || undefined,
        prompt: prompt || undefined,
        showErrorMessage: showError,
        errorStyle,
        errorTitle: errorTitle || undefined,
        error: error || undefined,
      };
    }
    commit(ui);
    setSelectionValidation(doc, sel, rule);
    return true;
  };

  return (
    <Dialog
      title="Data Validation"
      onClose={onClose}
      onOk={save}
      width={460}
      testId="validation-dialog"
      footerLeft={
        <button type="button" className="sh-btn" data-testid="validation-clear" onClick={() => { setType('none'); setF1(''); setF2(''); setPrompt(''); setPromptTitle(''); setError(''); setErrorTitle(''); }}>
          Clear All
        </button>
      }
    >
      <Tabs value={tab} onChange={setTab} tabs={[{ id: 'settings', label: 'Settings' }, { id: 'input', label: 'Input Message' }, { id: 'error', label: 'Error Alert' }]} testId="validation-tabs" />
      <div className="sh-tab-body">
        {tab === 'settings' && (
          <div className="sh-stack">
            <div className="sh-label">Validation criteria</div>
            <div className="sh-row wrap">
              <Field label="Allow:">
                <select value={type} data-testid="validation-allow" onChange={(e) => setType(e.target.value)}>
                  {ALLOW.map((a) => <option key={a.value} value={a.value}>{a.label}</option>)}
                </select>
              </Field>
              <Check label="Ignore blank" checked={ignoreBlank} disabled={type === 'none'} onChange={setIgnoreBlank} />
            </div>
            {usesOp && (
              <Field label="Data:">
                <select value={op} data-testid="validation-operator" onChange={(e) => setOp(e.target.value as ValidationOperator)}>
                  {OPERATORS.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
                </select>
              </Field>
            )}
            {type === 'list' && <Check label="In-cell dropdown" checked={dropDown} onChange={setDropDown} />}
            {type !== 'none' && (
              <Field label={type === 'list' ? 'Source:' : type === 'custom' ? 'Formula:' : usesOp && twoValues ? 'Minimum:' : 'Value:'} wide>
                <RangeField value={f1} onChange={setF1} label="Value" testId="validation-f1" />
              </Field>
            )}
            {usesOp && twoValues && (
              <Field label="Maximum:" wide>
                <RangeField value={f2} onChange={setF2} label="Maximum" testId="validation-f2" />
              </Field>
            )}
            {type === 'list' && <p className="sh-hint">Separate items with commas (Yes,No,Maybe) or refer to cells (=$A$1:$A$5).</p>}
          </div>
        )}
        {tab === 'input' && (
          <div className="sh-stack">
            <Check label="Show input message when cell is selected" checked={showInput} onChange={setShowInput} />
            <Field label="Title:" wide>
              <input value={promptTitle} maxLength={32} onChange={(e) => setPromptTitle(e.target.value)} />
            </Field>
            <Field label="Input message:" wide>
              <textarea rows={5} value={prompt} maxLength={255} data-testid="validation-prompt" onChange={(e) => setPrompt(e.target.value)} />
            </Field>
          </div>
        )}
        {tab === 'error' && (
          <div className="sh-stack">
            <Check label="Show error alert after invalid data is entered" checked={showError} onChange={setShowError} />
            <div className="sh-row wrap top">
              <Field label="Style:">
                <select value={errorStyle} data-testid="validation-style" onChange={(e) => setErrorStyle(e.target.value as typeof errorStyle)}>
                  <option value="stop">Stop</option>
                  <option value="warning">Warning</option>
                  <option value="information">Information</option>
                </select>
              </Field>
              <div className="sh-stack grow">
                <Field label="Title:" wide>
                  <input value={errorTitle} maxLength={32} onChange={(e) => setErrorTitle(e.target.value)} />
                </Field>
                <Field label="Error message:" wide>
                  <textarea rows={4} value={error} maxLength={255} data-testid="validation-error" onChange={(e) => setError(e.target.value)} />
                </Field>
              </div>
            </div>
          </div>
        )}
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Remove Duplicates

export function RemoveDuplicatesDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const region = useMemo(() => dataRegion(ui) ?? doc.range, []); // eslint-disable-line react-hooks/exhaustive-deps
  const [headers, setHeaders] = useState(() => looksLikeHeader(ui, region, region.c0));
  const cols = Array.from({ length: region.c1 - region.c0 + 1 }, (_, i) => region.c0 + i);
  const [chosen, setChosen] = useState<Set<number>>(new Set(cols));
  const name = (c: number) => (headers ? doc.displayText(region.r0, c) || `Column ${colName(c)}` : `Column ${colName(c)}`);
  return (
    <Dialog
      title="Remove Duplicates"
      onClose={onClose}
      width={420}
      testId="remove-duplicates"
      onOk={() => {
        if (!chosen.size) {
          ui.notify('Select at least one column.');
          return false;
        }
        commit(ui);
        const res = doc.removeDuplicates(region, cols.filter((c) => chosen.has(c)), headers);
        ui.notify(res.removed ? `${res.removed} duplicate value${res.removed > 1 ? 's' : ''} found and removed; ${res.remaining} unique value${res.remaining === 1 ? '' : 's'} remain.` : 'No duplicate values found.');
        return true;
      }}
    >
      <p className="sh-lead">To delete duplicate values, select one or more columns that contain duplicates.</p>
      <div className="sh-row wrap">
        <button type="button" className="sh-btn" onClick={() => setChosen(new Set(cols))}>Select All</button>
        <button type="button" className="sh-btn" onClick={() => setChosen(new Set())}>Unselect All</button>
        <span className="grow" />
        <Check label="My data has headers" checked={headers} onChange={setHeaders} testId="remove-duplicates-headers" />
      </div>
      <div className="sh-checklist" data-testid="remove-duplicates-columns">
        {cols.map((c) => (
          <Check key={c} label={name(c)} checked={chosen.has(c)} onChange={(v) => { const n = new Set(chosen); if (v) n.add(c); else n.delete(c); setChosen(n); }} />
        ))}
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Group / Ungroup, Subtotal, Outline settings

export function GroupDialog({ ungroup, onClose }: { ungroup?: boolean; onClose: () => void }) {
  const ui = useUI();
  const [axis, setAxis] = useState<'row' | 'col'>('row');
  return (
    <Dialog
      title={ungroup ? 'Ungroup' : 'Group'}
      onClose={onClose}
      width={260}
      testId="group-dialog"
      onOk={() => {
        commit(ui);
        const r = ui.doc.range;
        const [from, to] = axis === 'row' ? [r.r0, r.r1] : [r.c0, r.c1];
        if (ungroup) {
          if (!ui.doc.ungroupLines(axis, from, to)) ui.notify('The selection is not grouped.');
          return true;
        }
        const error = ui.doc.groupLines(axis, from, to);
        if (error) ui.notify(error);
        return true;
      }}
    >
      <Radio name="group-axis" label="Rows" checked={axis === 'row'} onChange={() => setAxis('row')} testId="group-rows" />
      <Radio name="group-axis" label="Columns" checked={axis === 'col'} onChange={() => setAxis('col')} testId="group-cols" />
    </Dialog>
  );
}

export function SubtotalDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const region = useMemo(() => dataRegion(ui) ?? doc.range, []); // eslint-disable-line react-hooks/exhaustive-deps
  const cols = Array.from({ length: region.c1 - region.c0 + 1 }, (_, i) => region.c0 + i);
  const name = (c: number) => doc.displayText(region.r0, c) || `Column ${colName(c)}`;
  const [groupBy, setGroupBy] = useState(region.c0);
  const [code, setCode] = useState(9);
  const [columns, setColumns] = useState<Set<number>>(new Set([region.c1]));
  const [replace, setReplace] = useState(true);
  const [pageBreaks, setPageBreaks] = useState(false);
  const [below, setBelow] = useState(true);
  return (
    <Dialog
      title="Subtotal"
      onClose={onClose}
      width={380}
      testId="subtotal-dialog"
      footerLeft={
        <button type="button" className="sh-btn" data-testid="subtotal-remove" onClick={() => { commit(ui); doc.removeSubtotals(region); onClose(); }}>
          Remove All
        </button>
      }
      onOk={() => {
        commit(ui);
        const error = doc.applySubtotals(region, { groupBy, code, columns: cols.filter((c) => columns.has(c)), replace, pageBreaks, summaryBelow: below });
        if (error) {
          ui.notify(error);
          return false;
        }
        return true;
      }}
    >
      <Field label="At each change in:" wide>
        <select value={groupBy} data-testid="subtotal-groupby" onChange={(e) => setGroupBy(Number(e.target.value))}>
          {cols.map((c) => <option key={c} value={c}>{name(c)}</option>)}
        </select>
      </Field>
      <Field label="Use function:" wide>
        <select value={code} data-testid="subtotal-function" onChange={(e) => setCode(Number(e.target.value))}>
          {SUBTOTAL_FUNCTIONS.map((f) => <option key={f.code} value={f.code}>{f.label}</option>)}
        </select>
      </Field>
      <div className="sh-label">Add subtotal to:</div>
      <div className="sh-checklist" data-testid="subtotal-columns">
        {cols.map((c) => (
          <Check key={c} label={name(c)} checked={columns.has(c)} onChange={(v) => { const n = new Set(columns); if (v) n.add(c); else n.delete(c); setColumns(n); }} />
        ))}
      </div>
      <Check label="Replace current subtotals" checked={replace} onChange={setReplace} />
      <Check label="Page break between groups" checked={pageBreaks} onChange={setPageBreaks} />
      <Check label="Summary below data" checked={below} onChange={setBelow} />
    </Dialog>
  );
}

export function OutlineSettingsDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const sheet = ui.doc.sheet;
  const [below, setBelow] = useState(sheet.outlineBelow);
  const [right, setRight] = useState(sheet.outlineRight);
  return (
    <Dialog
      title="Settings"
      onClose={onClose}
      width={320}
      testId="outline-settings"
      onOk={() => ui.doc.setOutlineSettings({ summaryBelow: below, summaryRight: right })}
      footerLeft={
        <button type="button" className="sh-btn" onClick={() => { ui.doc.setOutlineSettings({ summaryBelow: below, summaryRight: right }); if (!ui.doc.autoOutline()) ui.notify('Cannot create an outline.'); onClose(); }}>
          Create
        </button>
      }
    >
      <fieldset className="sh-group">
        <legend>Direction</legend>
        <Check label="Summary rows below detail" checked={below} onChange={setBelow} />
        <Check label="Summary columns to right of detail" checked={right} onChange={setRight} />
      </fieldset>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// What-If Analysis

export function GoalSeekDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const [setCell, setSetCell] = useState(rangeText(makeRange(doc.selection.active.r, doc.selection.active.c), doc.sheet));
  const [to, setTo] = useState('');
  const [by, setBy] = useState('');
  const [status, setStatus] = useState<{ text: string; target: number; result: number } | null>(null);
  if (status) {
    return (
      <Dialog
        title="Goal Seek Status"
        onClose={() => {
          // Cancel restores the input cell.
          doc.undo();
          onClose();
        }}
        onOk={() => {
          // OK keeps the value found (closing directly skips the Cancel path).
          onClose();
          return false;
        }}
        width={340}
        testId="goal-seek-status"
      >
        <p>{status.text}</p>
        <div className="sh-kv"><span>Target value:</span><b>{status.target}</b></div>
        <div className="sh-kv"><span>Current value:</span><b data-testid="goal-seek-result">{Number(status.result.toPrecision(12))}</b></div>
      </Dialog>
    );
  }
  return (
    <Dialog
      title="Goal Seek"
      onClose={onClose}
      width={340}
      testId="goal-seek"
      onOk={() => {
        const set = parseRangeText(doc, setCell);
        const changing = parseRangeText(doc, by);
        const target = Number(to);
        if (!set || set.range.r0 !== set.range.r1 || set.range.c0 !== set.range.c1) {
          ui.notify('Set cell must contain a single cell reference.');
          return false;
        }
        if (!Number.isFinite(target) || to.trim() === '') {
          ui.notify('To value must contain a number.');
          return false;
        }
        if (!changing || changing.range.r0 !== changing.range.r1 || changing.range.c0 !== changing.range.c1 || changing.sheet !== set.sheet) {
          ui.notify('By changing cell must contain a single cell reference on the same sheet.');
          return false;
        }
        commit(ui);
        const res = doc.goalSeek({ r: set.range.r0, c: set.range.c0 }, target, { r: changing.range.r0, c: changing.range.c0 }, set.sheet);
        if (typeof res === 'string') {
          ui.notify(res);
          return false;
        }
        const ref = cellName(set.range.r0, set.range.c0);
        setStatus({ text: res.found ? `Goal Seeking with Cell ${ref} found a solution.` : `Goal Seeking with Cell ${ref} may not have found a solution.`, target, result: res.result });
        return false;
      }}
    >
      <Field label="Set cell:" wide>
        <RangeField value={setCell} onChange={setSetCell} label="Set cell" testId="goal-seek-set" />
      </Field>
      <Field label="To value:" wide>
        <input value={to} inputMode="decimal" data-testid="goal-seek-to" onChange={(e) => setTo(e.target.value)} />
      </Field>
      <Field label="By changing cell:" wide>
        <RangeField value={by} onChange={setBy} label="By changing cell" testId="goal-seek-by" />
      </Field>
    </Dialog>
  );
}

export function DataTableDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const [rowInput, setRowInput] = useState('');
  const [colInput, setColInput] = useState('');
  return (
    <Dialog
      title="Data Table"
      onClose={onClose}
      width={340}
      testId="data-table"
      onOk={() => {
        const cell = (t: string) => {
          if (!t.trim()) return null;
          const hit = parseRangeText(doc, t);
          return hit && hit.range.r0 === hit.range.r1 && hit.range.c0 === hit.range.c1 && hit.sheet === doc.sheet ? { r: hit.range.r0, c: hit.range.c0 } : undefined;
        };
        const ri = cell(rowInput);
        const ci = cell(colInput);
        if (ri === undefined || ci === undefined) {
          ui.notify('Input cell reference is not valid.');
          return false;
        }
        commit(ui);
        const error = doc.dataTable(doc.range, ri, ci);
        if (error) {
          ui.notify(error);
          return false;
        }
        return true;
      }}
    >
      <Field label="Row input cell:" wide>
        <RangeField value={rowInput} onChange={setRowInput} label="Row input cell" testId="data-table-row" />
      </Field>
      <Field label="Column input cell:" wide>
        <RangeField value={colInput} onChange={setColInput} label="Column input cell" testId="data-table-col" />
      </Field>
      <p className="sh-hint">The results are calculated once and stored as values.</p>
    </Dialog>
  );
}

export function ScenariosDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const sheet = doc.sheet;
  const [list, setList] = useState<Scenario[]>(() => sheet.scenarios?.list.map((s) => ({ ...s, cells: s.cells.map((c) => ({ ...c })), values: [...s.values] })) ?? []);
  const [sel, setSel] = useState<number | null>(list.length ? 0 : null);
  const [edit, setEdit] = useState<{ index: number | null; name: string; cells: string; comment: string; step: 1 | 2; values: string[] } | null>(null);
  const [summary, setSummary] = useState(false);
  const save = (next: Scenario[]) => {
    setList(next);
    doc.setScenarios({ list: next, current: sheet.scenarios?.current });
  };
  const cellsText = scenarioCellsText;
  const parseCells = (text: string) => parseScenarioCells(doc, sheet, text);
  return (
    <>
      <Dialog title="Scenario Manager" onClose={onClose} closeOnly width={480} testId="scenario-manager">
        <div className="sh-two-col wide-left">
          <div className="sh-stack">
            <div className="sh-label">Scenarios:</div>
            <ListBox items={list.map((_, i) => i)} value={sel} keyOf={(i) => String(i)} height={150} testId="scenario-list" label="Scenarios" render={(i) => list[i].name} onChange={setSel} onActivate={(i) => doc.showScenario(i)} />
            {!list.length && <p className="sh-hint">No Scenarios defined. Choose Add to add scenarios.</p>}
          </div>
          <div className="sh-stack">
            <button type="button" className="sh-btn" data-testid="scenario-add" onClick={() => setEdit({ index: null, name: '', cells: doc.selection.ranges.map((r) => rangeText(r, sheet)).join(','), comment: `Created by ${doc.author} on ${new Date().toLocaleDateString()}`, step: 1, values: [] })}>Add...</button>
            <button type="button" className="sh-btn" disabled={sel === null} onClick={() => { if (sel === null) return; save(list.filter((_, i) => i !== sel)); setSel(list.length > 1 ? 0 : null); }}>Delete</button>
            <button type="button" className="sh-btn" disabled={sel === null} onClick={() => sel !== null && setEdit({ index: sel, name: list[sel].name, cells: cellsText(list[sel]), comment: list[sel].comment ?? '', step: 1, values: list[sel].values })}>Edit...</button>
            <button type="button" className="sh-btn" disabled={!list.length} onClick={() => setSummary(true)}>Summary...</button>
            <button type="button" className="sh-btn primary" disabled={sel === null} data-testid="scenario-show" onClick={() => sel !== null && doc.showScenario(sel)}>Show</button>
          </div>
        </div>
        {sel !== null && list[sel] && (
          <div className="sh-stack">
            <div className="sh-kv"><span>Changing cells:</span><b>{cellsText(list[sel])}</b></div>
            <div className="sh-kv"><span>Comment:</span><span>{list[sel].comment}</span></div>
          </div>
        )}
      </Dialog>
      {edit && edit.step === 1 && (
        <Dialog
          title={edit.index === null ? 'Add Scenario' : 'Edit Scenario'}
          onClose={() => setEdit(null)}
          width={400}
          testId="scenario-edit"
          onOk={() => {
            const cells = parseCells(edit.cells);
            if (!edit.name.trim()) {
              ui.notify('Enter a scenario name.');
              return false;
            }
            if (!cells) {
              ui.notify('Changing cells must be references on this sheet (up to 32 cells).');
              return false;
            }
            const values = cells.map((c, i) => edit.values[i] ?? doc.editText(c.r, c.c));
            setEdit({ ...edit, step: 2, values });
            return false;
          }}
        >
          <Field label="Scenario name:" wide>
            <input value={edit.name} data-testid="scenario-name" onChange={(e) => setEdit({ ...edit, name: e.target.value })} />
          </Field>
          <Field label="Changing cells:" wide>
            <RangeField value={edit.cells} onChange={(v) => setEdit({ ...edit, cells: v })} label="Changing cells" testId="scenario-cells" multi />
          </Field>
          <Field label="Comment:" wide>
            <textarea rows={3} value={edit.comment} onChange={(e) => setEdit({ ...edit, comment: e.target.value })} />
          </Field>
        </Dialog>
      )}
      {edit && edit.step === 2 && (
        <Dialog
          title="Scenario Values"
          onClose={() => setEdit(null)}
          width={360}
          testId="scenario-values"
          onOk={() => {
            const cells = parseCells(edit.cells)!;
            const sc: Scenario = { name: edit.name.trim(), cells, values: edit.values, comment: edit.comment || undefined, user: doc.author, locked: true };
            const next = edit.index === null ? [...list, sc] : list.map((x, i) => (i === edit.index ? sc : x));
            save(next);
            setSel(edit.index ?? next.length - 1);
            setEdit(null);
            return false;
          }}
        >
          <p className="sh-lead">Enter values for each of the changing cells.</p>
          {(parseCells(edit.cells) ?? []).map((c, i) => (
            <Field key={i} label={`${i + 1}: ${cellName(c.r, c.c, true, true)}`}>
              <input value={edit.values[i] ?? ''} data-testid={`scenario-value-${i}`} onChange={(e) => { const v = edit.values.slice(); v[i] = e.target.value; setEdit({ ...edit, values: v }); }} />
            </Field>
          ))}
        </Dialog>
      )}
      {summary && <ScenarioSummary onClose={() => setSummary(false)} />}
    </>
  );
}

function ScenarioSummary({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [result, setResult] = useState(rangeText(ui.doc.range, ui.doc.sheet));
  return (
    <Dialog
      title="Scenario Summary"
      onClose={onClose}
      width={360}
      testId="scenario-summary"
      onOk={() => {
        const cells = cellsOfRangeText(ui.doc, result);
        const sheet = ui.doc.sheet;
        const index = ui.doc.scenarioSummary(cells.slice(0, 32), sheet);
        if (index >= 0) ui.doc.setActiveSheet(index);
        return true;
      }}
    >
      <fieldset className="sh-group">
        <legend>Report type</legend>
        <Radio name="summary-type" label="Scenario summary" checked onChange={() => undefined} />
      </fieldset>
      <Field label="Result cells:" wide>
        <RangeField value={result} onChange={setResult} label="Result cells" multi />
      </Field>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Fill Series

export function SeriesDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const range = ui.doc.range;
  const [opts, setOpts] = useState<SeriesOptions>({ by: range.r1 - range.r0 >= range.c1 - range.c0 ? 'columns' : 'rows', type: 'linear', unit: 'day', step: 1, trend: false });
  const [step, setStep] = useState('1');
  const [stop, setStop] = useState('');
  const set = (p: Partial<SeriesOptions>) => setOpts({ ...opts, ...p });
  return (
    <Dialog
      title="Series"
      onClose={onClose}
      width={420}
      testId="series-dialog"
      onOk={() => {
        const s = Number(step);
        const st = stop.trim() === '' ? undefined : Number(stop);
        if (!Number.isFinite(s) || (st !== undefined && !Number.isFinite(st))) {
          ui.notify('Enter numbers for the step and stop values.');
          return false;
        }
        commit(ui);
        ui.doc.fillSeriesRange(range, { ...opts, step: s, stop: st });
        return true;
      }}
    >
      <div className="sh-row top wrap">
        <fieldset className="sh-group">
          <legend>Series in</legend>
          <Radio name="series-in" label="Rows" checked={opts.by === 'rows'} onChange={() => set({ by: 'rows' })} />
          <Radio name="series-in" label="Columns" checked={opts.by === 'columns'} onChange={() => set({ by: 'columns' })} />
        </fieldset>
        <fieldset className="sh-group">
          <legend>Type</legend>
          <Radio name="series-type" label="Linear" checked={opts.type === 'linear'} onChange={() => set({ type: 'linear' })} />
          <Radio name="series-type" label="Growth" checked={opts.type === 'growth'} onChange={() => set({ type: 'growth' })} />
          <Radio name="series-type" label="Date" checked={opts.type === 'date'} onChange={() => set({ type: 'date' })} />
          <Radio name="series-type" label="AutoFill" checked={opts.type === 'autofill'} onChange={() => set({ type: 'autofill' })} />
        </fieldset>
        <fieldset className="sh-group">
          <legend>Date unit</legend>
          {(['day', 'weekday', 'month', 'year'] as const).map((u) => (
            <Radio key={u} name="series-unit" label={u === 'weekday' ? 'Weekday' : u[0].toUpperCase() + u.slice(1)} checked={opts.unit === u} disabled={opts.type !== 'date'} onChange={() => set({ unit: u })} />
          ))}
        </fieldset>
      </div>
      <Check label="Trend" checked={opts.trend} disabled={opts.type === 'date' || opts.type === 'autofill'} onChange={(v) => set({ trend: v })} />
      <div className="sh-row">
        <Field label="Step value:">
          <input value={step} inputMode="decimal" style={{ width: 80 }} data-testid="series-step" onChange={(e) => setStep(e.target.value)} />
        </Field>
        <Field label="Stop value:">
          <input value={stop} inputMode="decimal" style={{ width: 80 }} data-testid="series-stop" onChange={(e) => setStop(e.target.value)} />
        </Field>
      </div>
    </Dialog>
  );
}
