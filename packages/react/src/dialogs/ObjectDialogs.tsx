/**
 * Tables, charts, sparklines and other objects: Create Table, Resize Table, Insert Chart (and
 * Change Chart Type), Select Data, Move Chart, Create Sparklines, Alt Text, Edit Text, Symbol.
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import { useUI } from '../context';
import { Check, Dialog, Field, ListBox, Radio, RangeField, parseRangeText, rangeText, selectionText } from './Dialog';
import { Icon } from '../Icon';
import { absRef, applySubtype, CHART_CATEGORIES, chartFromRange, chartPreviewSvg, type ChartSeries, chartSeriesName, chartSource, type ChartSubtype, COMBO_SERIES_KINDS as KINDS, comboKinds, commit, forcedChart, insertChart, localRef, makeRange, ONE_WAY_CHARTS as ONE_WAY, readChartData, resolveRef, type SeriesKind, subtypeOf, switchedSeries, SYMBOL_CHARS as SYMBOLS } from '@sheetline/core/views';

// -----------------------------------------------------------------------------------------------
// Tables

export function CreateTableDialog({ style, onClose }: { style?: string; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const initial = useMemo(() => {
    const src = chartSource(ui) ?? doc.range;
    return rangeText(src, doc.sheet);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const [ref, setRef] = useState(initial);
  const [headers, setHeaders] = useState(true);
  return (
    <Dialog
      title="Create Table"
      onClose={onClose}
      width={360}
      testId="create-table"
      onOk={() => {
        const hit = parseRangeText(doc, ref);
        if (!hit || hit.sheet !== doc.sheet) {
          ui.notify('The reference isn\'t valid. Select a range on this sheet.');
          return false;
        }
        commit(ui);
        const t = doc.createTable(hit.range, { hasHeaders: headers, style });
        if (typeof t === 'string') {
          ui.notify(t);
          return false;
        }
        doc.selectCell(t.ref.r0 + (t.headerRow ? 1 : 0), t.ref.c0);
        ui.showTab('table-design');
        return true;
      }}
    >
      <Field label="Where is the data for your table?" wide>
        <RangeField value={ref} onChange={setRef} label="Table range" testId="create-table-range" autoFocus />
      </Field>
      <Check label="My table has headers" checked={headers} onChange={setHeaders} testId="create-table-headers" />
    </Dialog>
  );
}

export function ResizeTableDialog({ tableId, onClose }: { tableId: number; onClose: () => void }) {
  const ui = useUI();
  const hit = ui.doc.tableById(tableId);
  const [ref, setRef] = useState(hit ? rangeText(hit.table.ref, hit.sheet) : '');
  if (!hit) return null;
  return (
    <Dialog
      title="Resize Table"
      onClose={onClose}
      width={360}
      testId="resize-table"
      onOk={() => {
        const r = parseRangeText(ui.doc, ref, hit.sheet);
        if (!r || r.sheet !== hit.sheet) {
          ui.notify('The reference isn\'t valid.');
          return false;
        }
        const error = ui.doc.resizeTable(tableId, r.range);
        if (error) {
          ui.notify(error);
          return false;
        }
        return true;
      }}
    >
      <Field label="Select the new data range for your table:" wide>
        <RangeField value={ref} onChange={setRef} label="Table range" testId="resize-table-range" home={hit.sheet} autoFocus />
      </Field>
      <p className="sh-hint">Note: The headers must remain in the same row, and the resulting table range must overlap the original table range.</p>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Insert Chart / Change Chart Type

export function InsertChartDialog({ category, changeId, onClose }: { category?: string; changeId?: string; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const existing = changeId ? doc.sheet.charts.find((c) => c.id === changeId) ?? null : null;
  const [cat, setCat] = useState(() => (existing ? CHART_CATEGORIES.find((c) => c.subtypes.includes(subtypeOf(existing)))?.id : category) ?? 'column');
  const catDef = CHART_CATEGORIES.find((c) => c.id === cat) ?? CHART_CATEGORIES[0];
  const [subId, setSubId] = useState<string>(() => (existing ? subtypeOf(existing).id : catDef.subtypes[0].id));
  const sub: ChartSubtype = catDef.subtypes.find((s) => s.id === subId) ?? catDef.subtypes[0];
  const source = useMemo(() => (existing ? null : chartSource(ui)), []); // eslint-disable-line react-hooks/exhaustive-deps
  const [orientation, setOrientation] = useState<0 | 1>(0);
  // The chosen category stays in view (the list is a sideways strip on phones).
  const catsRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = catsRef.current?.querySelector<HTMLElement>('[aria-selected="true"]');
    const strip = catsRef.current;
    if (el && strip && strip.scrollWidth > strip.clientWidth) strip.scrollLeft = Math.max(0, el.offsetLeft - strip.offsetLeft - 8);
  }, [cat]);
  const variants = useMemo(() => {
    if (existing || !source) return null;
    const a = chartFromRange(doc, doc.sheet, source, sub.type);
    if (ONE_WAY.has(sub.type)) return [a];
    return [a, forcedChart(doc, doc.sheet, source, sub.type, a.seriesIn === 'rows' ? 'columns' : 'rows')];
  }, [existing, source, sub.type, doc]);
  const seriesCount = (existing ?? variants?.[orientation] ?? variants?.[0])?.series.length ?? 0;
  const [custom, setCustom] = useState<{ kind: SeriesKind; secondary?: boolean }[]>(() =>
    existing?.type === 'combo' ? existing.series.map((s, i) => ({ kind: s.kind ?? (i === existing.series.length - 1 ? 'line' : 'column'), secondary: s.secondaryAxis })) : comboKinds('column-line', Math.max(1, seriesCount)),
  );
  const kinds = sub.combo === 'custom' ? custom : undefined;
  const chosen = existing ?? variants?.[Math.min(orientation, (variants?.length ?? 1) - 1)] ?? null;
  const candidate = chosen ? applySubtype(chosen, sub, kinds) : null;
  const previews = useMemo(() => variants?.map((v) => chartPreviewSvg(ui.doc, applySubtype(v, sub, kinds), 220, 150)) ?? null, [variants, sub, kinds, ui]);

  const ok = () => {
    if (!candidate) {
      ui.notify('Select the data for the chart first.');
      return false;
    }
    if (existing) {
      doc.updateChart(existing.id, { type: candidate.type, series: candidate.series, markers: candidate.markers, scatterStyle: candidate.scatterStyle, stockKind: candidate.stockKind });
      return true;
    }
    const id = insertChart(ui, sub, kinds);
    if (id && orientation === 1 && variants?.[1]) doc.updateChart(id, { seriesIn: candidate.seriesIn, headers: candidate.headers, categories: candidate.categories, series: candidate.series });
    return id !== null;
  };

  return (
    <Dialog title={existing ? 'Change Chart Type' : 'Insert Chart'} onClose={onClose} onOk={ok} width={760} testId="insert-chart" bodyClassName="sh-chart-dialog">
      <div className="sh-chart-dialog-grid">
        <div ref={catsRef} className="sh-chart-cats" role="listbox" aria-label="Chart categories" data-testid="chart-categories">
          {CHART_CATEGORIES.map((c) => (
            <button key={c.id} type="button" role="option" aria-selected={c.id === cat} className={c.id === cat ? 'selected' : ''} data-testid={`chart-cat-${c.id}`} onClick={() => { setCat(c.id); setSubId(c.subtypes[0].id); setOrientation(0); }}>
              <Icon name={c.icon} />
              <span>{c.label}</span>
            </button>
          ))}
        </div>
        <div className="sh-chart-main">
          <div className="sh-chart-subtypes" role="listbox" aria-label="Chart subtypes">
            {catDef.subtypes.map((s) => (
              <button key={s.id} type="button" role="option" aria-selected={s.id === sub.id} title={s.label} className={s.id === sub.id ? 'selected' : ''} data-testid={`chart-sub-${s.id}`} onClick={() => setSubId(s.id)}>
                <Icon name={s.icon} size={28} />
              </button>
            ))}
          </div>
          <div className="sh-chart-subtitle">{sub.label}</div>
          {sub.combo === 'custom' && chosen && (
            <div className="sh-combo-table">
              <div className="sh-combo-head"><span>Series Name</span><span>Chart Type</span><span>Secondary Axis</span></div>
              {chosen.series.map((s, i) => (
                <div key={i} className="sh-combo-row">
                  <span>{chartSeriesName(ui.doc, s, i)}</span>
                  <select value={custom[i]?.kind ?? 'column'} onChange={(e) => { const next = custom.slice(); next[i] = { ...(next[i] ?? {}), kind: e.target.value as SeriesKind }; setCustom(next); }}>
                    {KINDS.map((k) => <option key={k.kind} value={k.kind}>{k.label}</option>)}
                  </select>
                  <input type="checkbox" aria-label="Secondary Axis" checked={!!custom[i]?.secondary} onChange={(e) => { const next = custom.slice(); next[i] = { kind: next[i]?.kind ?? 'column', secondary: e.target.checked }; setCustom(next); }} />
                </div>
              ))}
            </div>
          )}
          {previews ? (
            <div className="sh-chart-previews">
              {previews.map((svg, i) => (
                <button
                  key={i}
                  type="button"
                  className={`sh-chart-preview${orientation === i ? ' selected' : ''}`}
                  data-testid={`chart-preview-${i}`}
                  title={variants?.[i]?.seriesIn === 'rows' ? 'Series in rows' : 'Series in columns'}
                  onClick={() => setOrientation(i as 0 | 1)}
                  // Generated by the chart renderer (texts escaped).
                  dangerouslySetInnerHTML={{ __html: svg }}
                />
              ))}
            </div>
          ) : candidate ? (
            <div className="sh-chart-preview single" dangerouslySetInnerHTML={{ __html: chartPreviewSvg(ui.doc, candidate, 420, 240) }} />
          ) : (
            <p className="sh-hint">Select the data you want to chart first: a range with labels in the first row or column works best.</p>
          )}
        </div>
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Select Data

export function SelectDataDialog({ chartId, onClose }: { chartId: string; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const chart = doc.sheet.charts.find((c) => c.id === chartId);
  const [range, setRange] = useState(chart?.range ? localRef(chart.range, doc.sheet) : '');
  const [series, setSeries] = useState<ChartSeries[]>(() => chart?.series.map((s) => ({ ...s })) ?? []);
  const [categories, setCategories] = useState(chart?.categories ? localRef(chart.categories, doc.sheet) : '');
  const [seriesIn, setSeriesIn] = useState(chart?.seriesIn ?? 'columns');
  const [headers, setHeaders] = useState(chart?.headers ?? true);
  const [sel, setSel] = useState<number | null>(series.length ? 0 : null);
  const [edit, setEdit] = useState<{ index: number | null; name: string; values: string } | null>(null);
  const [empty, setEmpty] = useState<'gap' | 'zero' | 'span'>('gap');
  if (!chart) return null;

  const qualify = (text: string) => {
    const hit = parseRangeText(doc, text);
    return hit ? absRef(hit.sheet, hit.range) : null;
  };

  const useRange = (text: string) => {
    setRange(text);
    const hit = parseRangeText(doc, text);
    if (!hit) return;
    const def = chartFromRange(doc, hit.sheet, hit.range, chart.type);
    setSeries(applySubtype(def, subtypeOf(chart)).series);
    setCategories(def.categories ? localRef(def.categories, doc.sheet) : '');
    setSeriesIn(def.seriesIn);
    setHeaders(def.headers);
    setSel(0);
  };

  return (
    <>
      <Dialog
        title="Select Data Source"
        onClose={onClose}
        width={600}
        testId="select-data"
        onOk={() => {
          const cat = categories.trim() ? qualify(categories) : undefined;
          if (categories.trim() && !cat) {
            ui.notify('The category labels reference isn\'t valid.');
            return false;
          }
          const hit = parseRangeText(doc, range);
          doc.updateChart(chart.id, { range: hit ? absRef(hit.sheet, hit.range) : chart.range, series, categories: cat ?? undefined, seriesIn, headers });
          void empty;
          return true;
        }}
      >
        <Field label="Chart data range:" wide>
          <RangeField value={range} onChange={useRange} label="Chart data range" testId="select-data-range" />
        </Field>
        <div className="sh-row center">
          <button
            type="button"
            className="sh-btn"
            data-testid="select-data-switch"
            onClick={() => {
              const next = seriesIn === 'rows' ? 'columns' : 'rows';
              const sw = switchedSeries(ui.doc, { ...chart, range: qualify(range) ?? chart.range, seriesIn, headers, categories: categories ? qualify(categories) ?? undefined : undefined, series }, next);
              if (sw.series) {
                setSeries(applySubtype({ ...chart, series: sw.series }, subtypeOf(chart)).series);
                setCategories(sw.categories ? localRef(sw.categories, doc.sheet) : '');
                setSeriesIn(next);
                setHeaders(!!sw.headers);
              }
            }}
          >
            <Icon name="swap_horiz" /> Switch Row/Column
          </button>
        </div>
        <div className="sh-two-col even">
          <fieldset className="sh-group">
            <legend>Legend Entries (Series)</legend>
            <div className="sh-row">
              <button type="button" className="sh-btn" data-testid="series-add" onClick={() => setEdit({ index: null, name: '', values: '' })}>Add</button>
              <button type="button" className="sh-btn" disabled={sel === null} onClick={() => sel !== null && setEdit({ index: sel, name: series[sel].nameRef ? localRef(series[sel].nameRef, doc.sheet) : series[sel].name ?? '', values: localRef(series[sel].values, doc.sheet) })}>Edit</button>
              <button type="button" className="sh-btn" disabled={sel === null} data-testid="series-remove" onClick={() => { if (sel === null) return; setSeries(series.filter((_, i) => i !== sel)); setSel(series.length > 1 ? Math.max(0, sel - 1) : null); }}>Remove</button>
              <button type="button" className="sh-btn icon" aria-label="Move up" disabled={!sel} onClick={() => { if (!sel) return; const n = series.slice(); [n[sel - 1], n[sel]] = [n[sel], n[sel - 1]]; setSeries(n); setSel(sel - 1); }}>▲</button>
              <button type="button" className="sh-btn icon" aria-label="Move down" disabled={sel === null || sel >= series.length - 1} onClick={() => { if (sel === null) return; const n = series.slice(); [n[sel + 1], n[sel]] = [n[sel], n[sel + 1]]; setSeries(n); setSel(sel + 1); }}>▼</button>
            </div>
            <ListBox items={series.map((_, i) => i)} value={sel} keyOf={(i) => String(i)} height={150} testId="series-list" label="Series" render={(i) => chartSeriesName(ui.doc, series[i], i)} onChange={setSel} />
          </fieldset>
          <fieldset className="sh-group">
            <legend>Horizontal (Category) Axis Labels</legend>
            <RangeField value={categories} onChange={setCategories} label="Axis label range" testId="select-data-categories" />
            <div className="sh-hint">Values: {readChartData(doc, { ...chart, categories: qualify(categories) ?? undefined }).categories.slice(0, 6).join(', ')}{readChartData(doc, chart).categories.length > 6 ? ', …' : ''}</div>
          </fieldset>
        </div>
        <Field label="Show empty cells as:">
          <select value={empty} onChange={(e) => setEmpty(e.target.value as typeof empty)}>
            <option value="gap">Gaps</option>
            <option value="zero">Zero</option>
            <option value="span">Connect data points with line</option>
          </select>
        </Field>
      </Dialog>
      {edit && (
        <Dialog
          title="Edit Series"
          onClose={() => setEdit(null)}
          width={380}
          testId="edit-series"
          onOk={() => {
            const values = qualify(edit.values);
            if (!values) {
              ui.notify('The series values reference isn\'t valid.');
              return false;
            }
            const nameRef = edit.name.trim() && parseRangeText(doc, edit.name) ? qualify(edit.name) ?? undefined : undefined;
            const s: ChartSeries = { ...(edit.index !== null ? series[edit.index] : {}), values, nameRef, name: nameRef ? undefined : edit.name || `Series ${series.length + 1}` };
            if (edit.index === null) {
              setSeries([...series, s]);
              setSel(series.length);
            } else setSeries(series.map((x, i) => (i === edit.index ? s : x)));
            return true;
          }}
        >
          <Field label="Series name:" wide>
            <RangeField value={edit.name} onChange={(v) => setEdit({ ...edit, name: v })} label="Series name" testId="edit-series-name" />
          </Field>
          <Field label="Series values:" wide>
            <RangeField value={edit.values} onChange={(v) => setEdit({ ...edit, values: v })} label="Series values" testId="edit-series-values" />
          </Field>
        </Dialog>
      )}
    </>
  );
}

// -----------------------------------------------------------------------------------------------
// Move Chart

export function MoveChartDialog({ chartId, onClose }: { chartId: string; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const sheets = doc.wb.sheets.filter((s) => s.state === 'visible' && !s.chartsheet);
  const [where, setWhere] = useState<'new' | 'object'>('object');
  const [newName, setNewName] = useState(() => {
    let k = 1;
    while (doc.wb.sheets.some((s) => s.name.toLowerCase() === `chart${k}`)) k++;
    return `Chart${k}`;
  });
  const [target, setTarget] = useState(doc.sheet.name);
  return (
    <Dialog
      title="Move Chart"
      onClose={onClose}
      width={420}
      testId="move-chart"
      onOk={() => {
        const from = doc.sheet;
        if (where === 'new') {
          const index = doc.activeIndex;
          const created = doc.addSheet(newName.trim() || undefined, index);
          if (!created) {
            ui.notify('That sheet name is not valid or is already used.');
            return false;
          }
          doc.moveChartToSheet(chartId, created, from);
          doc.setActiveSheet(doc.wb.sheets.indexOf(created));
          return true;
        }
        const t = doc.wb.sheetByName(target);
        if (t && t !== from) {
          doc.moveChartToSheet(chartId, t, from);
          doc.setActiveSheet(doc.wb.sheets.indexOf(t));
        }
        ui.selectObject(null);
        return true;
      }}
    >
      <p className="sh-lead">Choose where you want the chart to be placed:</p>
      <div className="sh-row">
        <Radio name="move-chart" label="New sheet:" checked={where === 'new'} onChange={() => setWhere('new')} testId="move-chart-new" />
        <input value={newName} disabled={where !== 'new'} aria-label="New sheet name" onChange={(e) => setNewName(e.target.value)} />
      </div>
      <div className="sh-row">
        <Radio name="move-chart" label="Object in:" checked={where === 'object'} onChange={() => setWhere('object')} />
        <select value={target} disabled={where !== 'object'} aria-label="Sheet" data-testid="move-chart-sheet" onChange={(e) => setTarget(e.target.value)}>
          {sheets.map((s) => <option key={s.uid} value={s.name}>{s.name}</option>)}
        </select>
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Sparklines

export function SparklineDialog({ type, groupId, single, onClose }: { type: 'line' | 'column' | 'stacked'; groupId?: string; single?: boolean; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const group = groupId ? doc.sheet.sparklineGroups.find((g) => g.id === groupId) : undefined;
  const { r, c } = doc.selection.active;
  const current = group?.sparklines.find((s) => s.r === r && s.c === c);
  const [data, setData] = useState(() => {
    if (single && current) return localRef(current.ref, doc.sheet);
    if (group) return groupData(group);
    return selectionText(doc);
  });
  const [location, setLocation] = useState(() => {
    if (single && current) return rangeText(makeRange(current.r, current.c), doc.sheet);
    if (group) return groupLocation(group);
    const sel = doc.range;
    // Default location: the empty column right of the data.
    return rangeText(makeRange(sel.r0, sel.c1 + 1, sel.r1, sel.c1 + 1), doc.sheet);
  });

  function groupData(g: NonNullable<typeof group>): string {
    const refs = g.sparklines.map((s) => resolveRef(doc, s.ref)).filter((x): x is NonNullable<ReturnType<typeof resolveRef>> => !!x);
    if (!refs.length) return '';
    const box = refs.reduce((b, x) => ({ r0: Math.min(b.r0, x.range.r0), c0: Math.min(b.c0, x.range.c0), r1: Math.max(b.r1, x.range.r1), c1: Math.max(b.c1, x.range.c1) }), refs[0].range);
    return rangeText(box, refs[0].sheet, doc.sheet);
  }
  function groupLocation(g: NonNullable<typeof group>): string {
    const box = g.sparklines.reduce((b, s) => ({ r0: Math.min(b.r0, s.r), c0: Math.min(b.c0, s.c), r1: Math.max(b.r1, s.r), c1: Math.max(b.c1, s.c) }), { r0: Infinity, c0: Infinity, r1: -Infinity, c1: -Infinity });
    return rangeText(box, doc.sheet);
  }

  return (
    <Dialog
      title={group ? (single ? "Edit Sparkline Data" : 'Edit Sparklines') : 'Create Sparklines'}
      onClose={onClose}
      width={420}
      testId="sparkline-dialog"
      onOk={() => {
        const d = parseRangeText(doc, data);
        if (!d) {
          ui.notify('Data Range reference is not valid.');
          return false;
        }
        if (single && group && current) {
          doc.updateSparklineGroup(group.id, { sparklines: group.sparklines.map((s) => (s === current ? { ...s, ref: absRef(d.sheet, d.range) } : s)) });
          return true;
        }
        const loc = parseRangeText(doc, location);
        if (!loc || loc.sheet !== doc.sheet) {
          ui.notify('Location reference is not valid. Select cells on this sheet.');
          return false;
        }
        commit(ui);
        const error = doc.addSparklines(group?.type ?? type, d, loc.range);
        if (error) {
          ui.notify(error);
          return false;
        }
        if (group) {
          // Keep the group's formatting on the rebuilt group.
          const added = doc.sheet.sparklineGroups[doc.sheet.sparklineGroups.length - 1];
          const { id: _id, sparklines: _s, ...look } = group;
          void _id;
          void _s;
          doc.updateSparklineGroup(added.id, look);
        }
        return true;
      }}
    >
      <p className="sh-lead">Choose the data that you want</p>
      <Field label="Data Range:" wide>
        <RangeField value={data} onChange={setData} label="Data Range" testId="sparkline-data" autoFocus />
      </Field>
      {!single && (
        <>
          <p className="sh-lead">Choose where you want the sparklines to be placed</p>
          <Field label="Location Range:" wide>
            <RangeField value={location} onChange={setLocation} label="Location Range" testId="sparkline-location" />
          </Field>
        </>
      )}
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Alt Text, shape text, symbols

export function AltTextDialog({ id, onClose }: { id: string; onClose: () => void }) {
  const ui = useUI();
  const hit = ui.doc.objectById(id);
  const [title, setTitle] = useState(hit && hit.kind !== 'image' ? hit.obj.altTitle ?? '' : '');
  const [descr, setDescr] = useState(hit ? hit.obj.descr ?? '' : '');
  if (!hit) return null;
  return (
    <Dialog title="Alt Text" onClose={onClose} width={400} testId="alt-text" onOk={() => ui.doc.setAltText(id, title, descr)}>
      <p className="sh-hint">How would you describe this object and its context to someone who is blind or low vision?</p>
      {hit.kind !== 'image' && (
        <Field label="Title" wide>
          <input value={title} data-testid="alt-text-title" onChange={(e) => setTitle(e.target.value)} />
        </Field>
      )}
      <Field label="Description" wide>
        <textarea rows={5} value={descr} data-testid="alt-text-description" onChange={(e) => setDescr(e.target.value)} />
      </Field>
    </Dialog>
  );
}

export function ShapeTextDialog({ id, onClose }: { id: string; onClose: () => void }) {
  const ui = useUI();
  const shape = ui.doc.sheet.shapes.find((s) => s.id === id);
  const [text, setText] = useState(shape?.text ?? '');
  const [size, setSize] = useState(String(shape?.fontSize ?? 11));
  const [bold, setBold] = useState(!!shape?.bold);
  const [italic, setItalic] = useState(!!shape?.italic);
  if (!shape) return null;
  return (
    <Dialog
      title={shape.textBox ? 'Text Box' : 'Edit Text'}
      onClose={onClose}
      width={420}
      testId="shape-text"
      onOk={() => {
        const n = Number(size);
        ui.doc.updateShape(id, { text: text || undefined, fontSize: Number.isFinite(n) && n >= 1 && n <= 409 ? n : shape.fontSize, bold: bold || undefined, italic: italic || undefined });
      }}
    >
      <textarea rows={6} autoFocus value={text} data-testid="shape-text-input" aria-label="Text" onChange={(e) => setText(e.target.value)} />
      <div className="sh-row">
        <Field label="Size:">
          <input value={size} inputMode="decimal" style={{ width: 60 }} onChange={(e) => setSize(e.target.value)} />
        </Field>
        <Check label="Bold" checked={bold} onChange={setBold} />
        <Check label="Italic" checked={italic} onChange={setItalic} />
      </div>
    </Dialog>
  );
}

export function SymbolDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [pick, setPick] = useState<string | null>(null);
  const insert = (ch: string) => {
    const g = ui.grid;
    if (!g) return;
    if (g.isEditing()) {
      const { text, caret } = g.editState;
      g.replaceEditText(text.slice(0, caret) + ch + text.slice(caret), caret + ch.length);
    } else g.startEdit(`${ui.doc.editText(ui.doc.selection.active.r, ui.doc.selection.active.c)}${ch}`, 'edit');
  };
  return (
    <Dialog title="Symbol" onClose={onClose} width={420} ok="Insert" blocking={false} testId="symbol-dialog" onOk={() => { if (pick) insert(pick); }}>
      <div className="sh-symbols" role="listbox" aria-label="Symbols">
        {SYMBOLS.map((ch) => (
          <button key={ch} type="button" role="option" aria-selected={pick === ch} className={pick === ch ? 'selected' : ''} onClick={() => setPick(ch)} onDoubleClick={() => { insert(ch); onClose(); }}>
            {ch}
          </button>
        ))}
      </div>
      <div className="sh-hint">Unicode: {pick ? `U+${pick.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}` : '—'}</div>
    </Dialog>
  );
}
