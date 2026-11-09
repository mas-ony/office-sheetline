/**
 * Side panels at the right of the grid: Threaded Comments, the Selection Pane (objects, show/hide,
 * order, names) and the Format pane of the selected chart, shape or picture.
 */
import { useEffect, useState, type ReactNode } from 'react';
import { useUI } from '../context';
import { Icon } from '../Icon';
import { CommentThreadCard } from '../Popovers';
import { Check, ColorButton, Field } from '../dialogs/Dialog';
import { cellName, cmd, type ColorSpec, type LineDash, objectRect, readChartData, resolveColor, type SeriesKind, setObjectSize, type SheetChart, type SheetShape, type Side, type Trendline } from '@sheetline/core/views';

function Panel({ title, children, testId, actions }: { title: string; children: ReactNode; testId: string; actions?: ReactNode }) {
  const ui = useUI();
  return (
    <aside className="sh-side-panel" data-testid={testId} aria-label={title}>
      <div className="sh-side-head">
        <span className="sh-side-title">{title}</span>
        {actions}
        <button type="button" className="sh-icon-btn" aria-label="Close panel" data-testid="side-panel-close" onClick={() => ui.setPanel(null)}>
          <Icon name="close" />
        </button>
      </div>
      <div className="sh-side-body">{children}</div>
    </aside>
  );
}

export function SidePanels() {
  const ui = useUI();
  switch (ui.panel) {
    case 'comments':
      return <CommentsPanel />;
    case 'selection':
      return <SelectionPane />;
    case 'format':
      return <FormatPane />;
    default:
      return null;
  }
}

// -----------------------------------------------------------------------------------------------
// Threaded Comments

function CommentsPanel() {
  const ui = useUI();
  const { doc, state } = ui;
  const [filter, setFilter] = useState<'all' | 'open' | 'resolved'>('all');
  const threads = doc.sheet.comments
    .filter((t) => t.kind === 'threaded')
    .filter((t) => (filter === 'all' ? true : filter === 'open' ? !t.resolved : t.resolved))
    .sort((a, b) => a.r - b.r || a.c - b.c);
  void state.version;
  return (
    <Panel
      title="Comments"
      testId="comments-panel"
      actions={
        <button type="button" className="sh-btn small" disabled={state.readOnly} data-testid="comments-panel-new" onClick={() => cmd.newComment(ui)}>
          <Icon name="add_comment" /> New
        </button>
      }
    >
      <div className="sh-row">
        <span className="sh-hint grow">{threads.length} items</span>
        <select value={filter} aria-label="Filter comments" onChange={(e) => setFilter(e.target.value as typeof filter)}>
          <option value="all">All</option>
          <option value="open">Active</option>
          <option value="resolved">Resolved</option>
        </select>
      </div>
      {threads.map((t) => (
        <div
          key={t.id}
          className={`sh-panel-thread${state.activeRef === cellName(t.r, t.c) ? ' active' : ''}`}
          onClick={() => {
            doc.selectCell(t.r, t.c);
            ui.grid?.scrollIntoView(t.r, t.c);
          }}
        >
          <div className="sh-panel-thread-ref">{cellName(t.r, t.c)}</div>
          <CommentThreadCard r={t.r} c={t.c} inPanel />
        </div>
      ))}
      {!threads.length && <p className="sh-hint">No comments yet. Select a cell and choose New to start a conversation.</p>}
    </Panel>
  );
}

// -----------------------------------------------------------------------------------------------
// Selection Pane

function SelectionPane() {
  const ui = useUI();
  const { doc } = ui;
  const sheet = doc.sheet;
  const [rename, setRename] = useState<{ id: string; name: string } | null>(null);
  void ui.state.version;
  const objects = [
    ...sheet.charts.map((o) => ({ id: o.id, name: o.name, z: o.z, hidden: !!o.hidden, kind: 'chart' as const })),
    ...sheet.shapes.map((o) => ({ id: o.id, name: o.name, z: o.z, hidden: !!o.hidden, kind: 'shape' as const })),
    ...sheet.images.map((o, i) => ({ id: o.id, name: o.name, z: o.z ?? i, hidden: !!o.hidden, kind: 'image' as const })),
  ].sort((a, b) => b.z - a.z);
  const sel = ui.selectedObject;
  return (
    <Panel title="Selection" testId="selection-pane">
      <div className="sh-row wrap">
        <button type="button" className="sh-btn small" onClick={() => doc.transact('Show all', () => objects.forEach((o) => o.hidden && doc.setObjectProps(o.id, { hidden: false })))}>Show All</button>
        <button type="button" className="sh-btn small" onClick={() => { ui.selectObject(null); doc.transact('Hide all', () => objects.forEach((o) => !o.hidden && doc.setObjectProps(o.id, { hidden: true }))); }}>Hide All</button>
        <span className="grow" />
        <button type="button" className="sh-icon-btn" aria-label="Bring Forward" disabled={!sel} onClick={() => sel && doc.reorderObject(sel, 'forward')}><Icon name="arrow_upward" /></button>
        <button type="button" className="sh-icon-btn" aria-label="Send Backward" disabled={!sel} onClick={() => sel && doc.reorderObject(sel, 'backward')}><Icon name="arrow_downward" /></button>
      </div>
      <div className="sh-object-list" role="listbox" aria-label="Objects on this sheet">
        {objects.map((o) => (
          <div key={o.id} role="option" aria-selected={sel === o.id} className={`sh-object-row${sel === o.id ? ' selected' : ''}${o.hidden ? ' hidden' : ''}`} data-testid={`object-row-${o.name}`} onClick={() => !o.hidden && ui.selectObject(o.id)} onDoubleClick={() => setRename({ id: o.id, name: o.name })}>
            <Icon name={o.kind === 'chart' ? 'insert_chart' : o.kind === 'shape' ? 'interests' : 'image'} />
            {rename?.id === o.id ? (
              <input
                autoFocus
                value={rename.name}
                aria-label="Object name"
                onChange={(e) => setRename({ ...rename, name: e.target.value })}
                onBlur={() => { if (rename.name.trim()) doc.setObjectProps(o.id, { name: rename.name.trim() }); setRename(null); }}
                onKeyDown={(e) => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur(); if (e.key === 'Escape') setRename(null); }}
              />
            ) : (
              <span className="grow">{o.name}</span>
            )}
            <button type="button" className="sh-icon-btn" aria-label={o.hidden ? 'Show' : 'Hide'} data-testid={`object-eye-${o.name}`} onClick={(e) => { e.stopPropagation(); if (!o.hidden && sel === o.id) ui.selectObject(null); doc.setObjectProps(o.id, { hidden: !o.hidden }); }}>
              <Icon name={o.hidden ? 'visibility_off' : 'visibility'} />
            </button>
          </div>
        ))}
        {!objects.length && <p className="sh-hint">There are no shapes, charts or pictures on this sheet.</p>}
      </div>
    </Panel>
  );
}

// -----------------------------------------------------------------------------------------------
// Format pane

const hexOf = (ui: ReturnType<typeof useUI>, c: ColorSpec | null) => (c ? `#${resolveColor(c, ui.doc.wb.theme, '000000')}` : null);
const specOf = (hex: string | null | undefined): ColorSpec | null => (hex ? { rgb: `FF${hex.replace('#', '').toUpperCase()}` } : null);

const DASHES: { dash: LineDash; label: string }[] = [
  { dash: 'solid', label: 'Solid' },
  { dash: 'dashed', label: 'Dash' },
  { dash: 'dotted', label: 'Round Dot' },
  { dash: 'dash-dot', label: 'Dash Dot' },
];

function Section({ title, children, open: initial = true }: { title: string; children: ReactNode; open?: boolean }) {
  const [open, setOpen] = useState(initial);
  return (
    <section className={`sh-pane-section${open ? ' open' : ''}`}>
      <button type="button" className="sh-pane-section-head" aria-expanded={open} onClick={() => setOpen(!open)}>
        <Icon name={open ? 'expand_more' : 'chevron_right'} />
        {title}
      </button>
      {open && <div className="sh-pane-section-body">{children}</div>}
    </section>
  );
}

/** Text box that commits on blur or Enter. */
function LazyInput({ value, onCommit, label, testId, inputMode }: { value: string; onCommit: (v: string) => void; label: string; testId?: string; inputMode?: 'decimal' | 'text' }) {
  const [v, setV] = useState(value);
  useEffect(() => setV(value), [value]);
  return (
    <input
      value={v}
      aria-label={label}
      data-testid={testId}
      inputMode={inputMode}
      onChange={(e) => setV(e.target.value)}
      onBlur={() => v !== value && onCommit(v)}
      onKeyDown={(e) => e.key === 'Enter' && (e.target as HTMLInputElement).blur()}
    />
  );
}

function FormatPane() {
  const ui = useUI();
  const id = ui.selectedObject;
  const hit = id ? ui.doc.objectById(id) : null;
  void ui.state.version;
  if (!hit) {
    return (
      <Panel title="Format" testId="format-pane">
        <p className="sh-hint">Select a chart, shape or picture to format it.</p>
      </Panel>
    );
  }
  if (hit.kind === 'chart') return <ChartFormat chart={hit.obj} />;
  if (hit.kind === 'shape') return <ShapeFormat shape={hit.obj} />;
  return (
    <Panel title="Format Picture" testId="format-pane">
      <SizeSection id={hit.obj.id} />
      <Section title="Alt Text">
        <button type="button" className="sh-btn" onClick={() => ui.openDialog({ kind: 'alt-text', id: hit.obj.id })}>Edit Alt Text...</button>
      </Section>
    </Panel>
  );
}

function SizeSection({ id }: { id: string }) {
  const ui = useUI();
  const r = objectRect(ui, id);
  if (!r) return null;
  const inches = (px: number) => String(Math.round((px / 96) * 100) / 100);
  return (
    <Section title="Size">
      <Field label="Height (in):">
        <LazyInput value={inches(r.h)} label="Height" inputMode="decimal" onCommit={(v) => Number(v) > 0 && setObjectSize(ui, id, r.w, Number(v) * 96)} />
      </Field>
      <Field label="Width (in):">
        <LazyInput value={inches(r.w)} label="Width" inputMode="decimal" onCommit={(v) => Number(v) > 0 && setObjectSize(ui, id, Number(v) * 96, r.h)} />
      </Field>
    </Section>
  );
}

function LineEditor({ line, onChange, noneLabel = 'No line' }: { line: { color: string; width: number; dash?: LineDash } | null | undefined; onChange: (l: { color: string; width: number; dash?: LineDash } | null) => void; noneLabel?: string }) {
  const ui = useUI();
  const on = !!line;
  return (
    <div className="sh-stack tight">
      <div className="sh-row">
        <label className="sh-check"><input type="radio" checked={!on} onChange={() => onChange(null)} /> <span>{noneLabel}</span></label>
        <label className="sh-check"><input type="radio" checked={on} onChange={() => onChange({ color: line?.color ?? '#595959', width: line?.width ?? 1 })} /> <span>Solid line</span></label>
      </div>
      {on && line && (
        <>
          <Field label="Color:">
            <ColorButton value={specOf(line.color)} noneLabel="Automatic" onChange={(c) => onChange({ ...line, color: hexOf(ui, c) ?? '#595959' })} />
          </Field>
          <Field label="Width (px):">
            <LazyInput value={String(line.width)} label="Line width" inputMode="decimal" onCommit={(v) => Number(v) > 0 && onChange({ ...line, width: Math.min(20, Number(v)) })} />
          </Field>
          <Field label="Dash type:">
            <select value={line.dash ?? 'solid'} onChange={(e) => onChange({ ...line, dash: e.target.value === 'solid' ? undefined : (e.target.value as LineDash) })}>
              {DASHES.map((d) => <option key={d.dash} value={d.dash}>{d.label}</option>)}
            </select>
          </Field>
        </>
      )}
    </div>
  );
}

type ChartPart = 'area' | 'title' | 'legend' | 'haxis' | 'vaxis' | `series:${number}`;

function ChartFormat({ chart }: { chart: SheetChart }) {
  const ui = useUI();
  const { doc } = ui;
  const [part, setPart] = useState<ChartPart>('area');
  const update = (patch: Partial<SheetChart>) => doc.updateChart(chart.id, patch);
  const data = readChartData(doc, chart);
  const pie = ['pie', 'doughnut', 'treemap', 'sunburst', 'funnel'].includes(chart.type);
  const seriesIndex = part.startsWith('series:') ? Number(part.slice(7)) : -1;
  const series = seriesIndex >= 0 ? chart.series[seriesIndex] : undefined;
  const setSeries = (patch: Partial<SheetChart['series'][number]>) => update({ series: chart.series.map((s, i) => (i === seriesIndex ? { ...s, ...patch } : s)) });
  const lineLike = chart.type.endsWith('line') || chart.type === 'radar' || chart.type === 'scatter' || (chart.type === 'combo' && series?.kind === 'line');
  return (
    <Panel title="Format Chart" testId="format-pane">
      <Field label="Chart Options:" wide>
        <select value={part} data-testid="format-chart-part" onChange={(e) => setPart(e.target.value as ChartPart)}>
          <option value="area">Chart Area</option>
          <option value="title">Chart Title</option>
          <option value="legend">Legend</option>
          {!pie && <option value="haxis">Horizontal (Category) Axis</option>}
          {!pie && <option value="vaxis">Vertical (Value) Axis</option>}
          {chart.series.map((s, i) => <option key={i} value={`series:${i}`}>Series "{data.series[i]?.name ?? s.name ?? i + 1}"</option>)}
        </select>
      </Field>
      {part === 'area' && (
        <>
          <Section title="Fill">
            <div className="sh-row">
              <label className="sh-check"><input type="radio" checked={chart.fill === null} onChange={() => update({ fill: null })} /> <span>No fill</span></label>
              <label className="sh-check"><input type="radio" checked={chart.fill !== null} onChange={() => update({ fill: chart.fill ?? undefined })} /> <span>Solid fill</span></label>
            </div>
            {chart.fill !== null && (
              <Field label="Color:">
                <ColorButton value={specOf(chart.fill ?? '#FFFFFF')} testId="chart-fill" noneLabel="Automatic" onChange={(c) => update({ fill: hexOf(ui, c) ?? undefined })} />
              </Field>
            )}
          </Section>
          <Section title="Border">
            <LineEditor line={chart.border === undefined ? { color: '#D9D9D9', width: 1 } : chart.border} onChange={(l) => update({ border: l })} />
          </Section>
          <Section title="Data Labels and Gridlines">
            <Check label="Data labels" checked={chart.valueLabels} onChange={(v) => update({ valueLabels: v })} />
            {!pie && <Check label="Major horizontal gridlines" checked={chart.gridlines} onChange={(v) => update({ gridlines: v })} />}
            {!pie && <Check label="Major vertical gridlines" checked={!!chart.vGridlines} onChange={(v) => update({ vGridlines: v || undefined })} />}
          </Section>
          <SizeSection id={chart.id} />
          <Section title="Properties" open={false}>
            <Field label="Name:" wide>
              <LazyInput value={chart.name} label="Chart name" onCommit={(v) => v.trim() && update({ name: v.trim() })} />
            </Field>
          </Section>
        </>
      )}
      {part === 'title' && (
        <Section title="Title">
          <Check label="Show chart title" checked={!!chart.title} onChange={(v) => update({ title: v ? chart.title || 'Chart Title' : undefined })} />
          <Field label="Text:" wide>
            <LazyInput value={chart.title ?? ''} label="Chart title" testId="chart-title-input" onCommit={(v) => update({ title: v || undefined })} />
          </Field>
          <Field label="Subtitle:" wide>
            <LazyInput value={chart.subtitle ?? ''} label="Subtitle" onCommit={(v) => update({ subtitle: v || undefined })} />
          </Field>
          <Field label="Position:">
            <select value={chart.titlePosition ?? 'top'} onChange={(e) => update({ titlePosition: e.target.value as Side })}>
              <option value="top">Top</option>
              <option value="bottom">Bottom</option>
              <option value="left">Left</option>
              <option value="right">Right</option>
            </select>
          </Field>
        </Section>
      )}
      {part === 'legend' && (
        <Section title="Legend">
          <Check label="Show legend" checked={chart.showLegend} onChange={(v) => update({ showLegend: v })} testId="chart-legend-show" />
          <Field label="Position:">
            <select value={chart.legendPosition ?? 'bottom'} onChange={(e) => update({ legendPosition: e.target.value as Side, showLegend: true })}>
              <option value="bottom">Bottom</option>
              <option value="top">Top</option>
              <option value="left">Left</option>
              <option value="right">Right</option>
            </select>
          </Field>
          <Field label="Title:" wide>
            <LazyInput value={chart.legendTitle ?? ''} label="Legend title" onCommit={(v) => update({ legendTitle: v || undefined })} />
          </Field>
        </Section>
      )}
      {part === 'haxis' && (
        <Section title="Horizontal Axis">
          <Check label="Show axis" checked={!chart.hideXAxis} onChange={(v) => update({ hideXAxis: !v || undefined })} />
          <Field label="Axis title:" wide>
            <LazyInput value={chart.xAxisTitle ?? ''} label="Horizontal axis title" onCommit={(v) => update({ xAxisTitle: v || undefined })} />
          </Field>
        </Section>
      )}
      {part === 'vaxis' && (
        <Section title="Vertical Axis">
          <Check label="Show axis" checked={!chart.hideYAxis} onChange={(v) => update({ hideYAxis: !v || undefined })} />
          <Field label="Axis title:" wide>
            <LazyInput value={chart.yAxisTitle ?? ''} label="Vertical axis title" onCommit={(v) => update({ yAxisTitle: v || undefined })} />
          </Field>
          <Field label="Minimum:">
            <LazyInput value={chart.yMin === undefined ? '' : String(chart.yMin)} label="Minimum" inputMode="decimal" testId="chart-ymin" onCommit={(v) => update({ yMin: v.trim() === '' || !Number.isFinite(Number(v)) ? undefined : Number(v) })} />
          </Field>
          <Field label="Maximum:">
            <LazyInput value={chart.yMax === undefined ? '' : String(chart.yMax)} label="Maximum" inputMode="decimal" onCommit={(v) => update({ yMax: v.trim() === '' || !Number.isFinite(Number(v)) ? undefined : Number(v) })} />
          </Field>
          <Field label="Secondary axis title:" wide>
            <LazyInput value={chart.secondaryAxisTitle ?? ''} label="Secondary axis title" onCommit={(v) => update({ secondaryAxisTitle: v || undefined })} />
          </Field>
        </Section>
      )}
      {series && (
        <>
          <Section title="Series Options">
            {chart.type === 'combo' && (
              <Field label="Chart type:">
                <select value={series.kind ?? 'column'} onChange={(e) => setSeries({ kind: e.target.value as SeriesKind })}>
                  <option value="column">Clustered Column</option>
                  <option value="stacked-column">Stacked Column</option>
                  <option value="line">Line</option>
                  <option value="area">Area</option>
                  <option value="stacked-area">Stacked Area</option>
                </select>
              </Field>
            )}
            {!pie && <Check label="Plot on secondary axis" checked={!!series.secondaryAxis} onChange={(v) => setSeries({ secondaryAxis: v || undefined })} />}
            {!pie && (
              <Field label="Trendline:">
                <select value={series.trendline ?? 'none'} onChange={(e) => setSeries({ trendline: e.target.value === 'none' ? undefined : (e.target.value as Trendline) })}>
                  {['none', 'linear', 'exponential', 'logarithmic', 'polynomial', 'power', 'moving-average'].map((t) => <option key={t} value={t}>{t === 'none' ? 'None' : t[0].toUpperCase() + t.slice(1).replace('-', ' ')}</option>)}
                </select>
              </Field>
            )}
          </Section>
          <Section title={lineLike ? 'Line' : 'Fill'}>
            <Field label="Color:">
              <ColorButton value={specOf(series.color)} noneLabel="Automatic" testId="series-color" onChange={(c) => setSeries({ color: hexOf(ui, c) ?? undefined })} />
            </Field>
            {lineLike && (
              <>
                <Field label="Width (px):">
                  <LazyInput value={String(series.lineWidth ?? 2.5)} label="Line width" inputMode="decimal" onCommit={(v) => Number(v) > 0 && setSeries({ lineWidth: Math.min(20, Number(v)) })} />
                </Field>
                <Field label="Dash type:">
                  <select value={series.lineStyle ?? 'solid'} onChange={(e) => setSeries({ lineStyle: e.target.value === 'solid' ? undefined : (e.target.value as LineDash) })}>
                    {DASHES.map((d) => <option key={d.dash} value={d.dash}>{d.label}</option>)}
                  </select>
                </Field>
                <Check label="Markers" checked={series.markers ?? !!chart.markers} onChange={(v) => setSeries({ markers: v })} />
                <Check label="Smoothed line" checked={!!series.smooth} onChange={(v) => setSeries({ smooth: v || undefined })} />
              </>
            )}
          </Section>
        </>
      )}
    </Panel>
  );
}

function ShapeFormat({ shape }: { shape: SheetShape }) {
  const ui = useUI();
  const update = (patch: Partial<SheetShape>) => ui.doc.updateShape(shape.id, patch);
  return (
    <Panel title="Format Shape" testId="format-pane">
      <Section title="Fill">
        <div className="sh-row">
          <label className="sh-check"><input type="radio" checked={!shape.fill} onChange={() => update({ fill: null })} /> <span>No fill</span></label>
          <label className="sh-check"><input type="radio" checked={!!shape.fill} onChange={() => update({ fill: shape.fill ?? '#5B9BD5' })} /> <span>Solid fill</span></label>
        </div>
        {shape.fill && (
          <Field label="Color:">
            <ColorButton value={specOf(shape.fill)} noneLabel="No Fill" testId="shape-fill-color" onChange={(c) => update({ fill: hexOf(ui, c) })} />
          </Field>
        )}
      </Section>
      <Section title="Line">
        <LineEditor line={shape.line} onChange={(l) => update({ line: l })} />
      </Section>
      <Section title="Text">
        <Field label="Text:" wide>
          <LazyInput value={shape.text ?? ''} label="Shape text" testId="shape-text-pane" onCommit={(v) => update({ text: v || undefined })} />
        </Field>
        <Field label="Color:">
          <ColorButton value={specOf(shape.fontColor)} noneLabel="Automatic" onChange={(c) => update({ fontColor: hexOf(ui, c) ?? '#000000' })} />
        </Field>
        <Field label="Size (pt):">
          <LazyInput value={String(shape.fontSize ?? 11)} label="Font size" inputMode="decimal" onCommit={(v) => Number(v) >= 1 && update({ fontSize: Math.min(409, Number(v)) })} />
        </Field>
        <div className="sh-row">
          <Check label="Bold" checked={!!shape.bold} onChange={(v) => update({ bold: v || undefined })} />
          <Check label="Italic" checked={!!shape.italic} onChange={(v) => update({ italic: v || undefined })} />
        </div>
      </Section>
      <Section title="Rotation">
        <Field label="Degrees:">
          <LazyInput value={String(shape.rotation ?? 0)} label="Rotation" inputMode="decimal" onCommit={(v) => Number.isFinite(Number(v)) && update({ rotation: ((Number(v) % 360) + 360) % 360 || undefined })} />
        </Field>
      </Section>
      <SizeSection id={shape.id} />
    </Panel>
  );
}
