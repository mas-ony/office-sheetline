/** The galleries of ribbon drop-downs and the controls with fields of their own. */
import { useEffect, useState, type CSSProperties } from 'react';
import { Icon } from '../Icon';
import { applyCellStyle, applyShapeStyle, CELL_STYLE_GROUPS, CELL_STYLES, cellStylePreview, changeShapeTo, cmd, type Command, type Ctx, currentTable, galleryStyles, insertShape, objectRect, type PanelSpec, PASTE_GROUPS, pxToInches, renameTableTo, setObjectSize, SHAPE_GALLERY, SHAPE_STYLES, type ShapeGalleryItem, shapeIconSvg, sparklineStyles, tableStyleLabel, tableStylePreviewSvg, themeTiles, type ViewSpec } from '@sheetline/core/views';

type PanelProps = { ctx: Ctx; close: () => void };

function PastePanel({ ctx, close }: PanelProps) {
  const has = ctx.doc.hasClip();
  return (
    <div className="sh-paste-panel" data-testid="paste-panel">
      {PASTE_GROUPS.map((g) => (
        <div key={g.title}>
          <div className="sh-panel-title">{g.title}</div>
          <div className="sh-icon-grid">
            {g.items.map((it) => (
              <button
                key={it.id}
                type="button"
                className="sh-icon-tile"
                title={it.label}
                aria-label={it.label}
                data-testid={`paste-${it.id}`}
                onClick={() => {
                  close();
                  void cmd.paste(ctx, it.opts);
                }}
              >
                <Icon name="content_paste" className="base" />
                <Icon name={it.icon} className="badge" />
              </button>
            ))}
          </div>
        </div>
      ))}
      <div className="sh-panel-sep" />
      <button type="button" className="sh-panel-link" disabled={!has} data-testid="paste-special" onClick={() => { close(); ctx.openDialog({ kind: 'paste-special' }); }}>
        Paste Special...
      </button>
    </div>
  );
}

function TableStylesPanel({ ctx, close, onPick, showNone }: PanelProps & { onPick: (name: string) => void; showNone?: boolean }) {
  const theme = ctx.doc.wb.theme;
  const tile = (name: string, label: string) => (
    <button
      key={name || 'none'}
      type="button"
      className="sh-style-tile"
      title={label}
      aria-label={label}
      data-testid={`table-style-${name || 'none'}`}
      onClick={() => {
        close();
        onPick(name);
      }}
      dangerouslySetInnerHTML={{ __html: tableStylePreviewSvg(name, theme, 58) }}
    />
  );
  return (
    <div className="sh-gallery-panel" data-testid="table-styles">
      {(['Light', 'Medium', 'Dark'] as const).map((cat) => (
        <div key={cat}>
          <div className="sh-panel-title">{cat}</div>
          <div className="sh-style-grid">
            {cat === 'Light' && showNone && tile('', 'None')}
            {galleryStyles(cat).map((s) => tile(s.name, tableStyleLabel(s.name)))}
          </div>
        </div>
      ))}
    </div>
  );
}

function CellStylesPanel({ ctx, close }: PanelProps) {
  return (
    <div className="sh-gallery-panel cell-styles" data-testid="cell-styles">
      {CELL_STYLE_GROUPS.map((g) => (
        <div key={g}>
          <div className="sh-panel-title">{g}</div>
          <div className="sh-cellstyle-grid">
            {CELL_STYLES.filter((s) => s.group === g).map((s) => (
              <button
                key={s.name}
                type="button"
                className="sh-cellstyle"
                style={cellStylePreview(ctx, s) as CSSProperties}
                title={s.name}
                data-testid={`cell-style-${s.name.replace(/[^A-Za-z0-9]+/g, '-').toLowerCase()}`}
                onClick={() => {
                  close();
                  applyCellStyle(ctx, s);
                }}
              >
                {s.name}
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

/** Insert > Shapes and Shape Format > Change Shape. */
function ShapesPanel({ close, onPick, testId, itemTestIds }: { close: () => void; onPick: (it: ShapeGalleryItem) => void; testId: string; itemTestIds?: boolean }) {
  return (
    <div className="sh-shapes-panel" data-testid={testId}>
      {SHAPE_GALLERY.map((group) => (
        <div key={group.title}>
          <div className="sh-panel-title">{group.title}</div>
          <div className="sh-shape-grid">
            {group.items.map((it, i) => (
              <button
                key={`${it.preset}-${i}`}
                type="button"
                title={it.label}
                aria-label={it.label}
                data-testid={itemTestIds ? `shape-${it.preset}${it.arrowEnd ? (it.arrowStart ? '-both' : '-end') : ''}` : undefined}
                onClick={() => {
                  close();
                  onPick(it);
                }}
                // Generated from the shape geometry (no user content).
                dangerouslySetInnerHTML={{ __html: shapeIconSvg(it) }}
              />
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function ShapeStylesPanel({ ctx, close }: PanelProps) {
  return (
    <div className="sh-shape-styles" data-testid="shape-styles-panel">
      {SHAPE_STYLES.map((st) => (
        <button
          key={st.label}
          type="button"
          title={st.label}
          className="sh-shape-style"
          style={{ background: st.fill ?? 'transparent', borderColor: st.line, color: st.font }}
          onClick={() => {
            close();
            applyShapeStyle(ctx, st);
          }}
        >
          Abc
        </button>
      ))}
    </div>
  );
}

function ThemesPanel({ ctx, close }: PanelProps) {
  return (
    <div className="sh-themes-panel" data-testid="themes-panel">
      <div className="sh-panel-title">Built-In</div>
      <div className="sh-theme-grid">
        {themeTiles(ctx).map((t) => (
          <button
            key={t.name}
            type="button"
            className={`sh-theme-tile${t.active ? ' active' : ''}`}
            title={t.name}
            data-testid={t.testId}
            onClick={() => {
              close();
              t.apply();
            }}
          >
            <span className="sh-theme-preview" style={{ background: t.background, color: t.color }}>
              <span className="sh-theme-aa" style={{ fontFamily: t.font }}>Aa</span>
              <span className="sh-theme-bars">
                {t.accents.map((x, i) => (
                  <i key={i} style={{ background: x }} />
                ))}
              </span>
            </span>
            <span className="sh-theme-name">{t.name}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

function SparklineStylesPanel({ ctx, close }: PanelProps) {
  return (
    <div className="sh-spark-styles" data-testid="sparkline-styles">
      {sparklineStyles(ctx).map((st) => (
        <button
          key={st.key}
          type="button"
          title={st.title}
          onClick={() => {
            close();
            st.apply();
          }}
        >
          <svg width="44" height="22" viewBox="0 0 44 22">
            <polyline points="2,16 9,8 16,12 23,4 30,14 37,9 42,11" fill="none" stroke={st.stroke} strokeWidth="1.5" />
          </svg>
        </button>
      ))}
    </div>
  );
}

/** The content of a ribbon drop-down panel. */
export function RibbonPanel({ spec, ctx, close }: PanelProps & { spec: PanelSpec }) {
  switch (spec.kind) {
    case 'paste':
      return <PastePanel ctx={ctx} close={close} />;
    case 'table-styles':
      return <TableStylesPanel ctx={ctx} close={close} showNone={spec.showNone} onPick={(name) => spec.pick(ctx, name)} />;
    case 'cell-styles':
      return <CellStylesPanel ctx={ctx} close={close} />;
    case 'shapes':
      return <ShapesPanel close={close} testId="shapes-panel" itemTestIds onPick={(it) => insertShape(ctx, it)} />;
    case 'change-shape':
      return <ShapesPanel close={close} testId="change-shape-panel" onPick={(it) => changeShapeTo(ctx, it)} />;
    case 'shape-styles':
      return <ShapeStylesPanel ctx={ctx} close={close} />;
    case 'themes':
      return <ThemesPanel ctx={ctx} close={close} />;
    case 'sparkline-styles':
      return <SparklineStylesPanel ctx={ctx} close={close} />;
  }
}

function TableName({ ctx }: { ctx: Ctx }) {
  const t = currentTable(ctx);
  const [draft, setDraft] = useState(t?.name ?? '');
  useEffect(() => setDraft(t?.name ?? ''), [t?.name]);
  if (!t) return null;
  const commit = () => {
    if (draft !== t.name && !renameTableTo(ctx, t.id, draft)) setDraft(t.name);
  };
  return (
    <label className="sh-field-inline">
      <span>Table Name:</span>
      <input
        aria-label="Table Name"
        data-testid="table-name"
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
          if (e.key === 'Escape') setDraft(t.name);
        }}
      />
    </label>
  );
}

function ObjectSize({ ctx }: { ctx: Ctx }) {
  const id = ctx.selectedObject;
  const r = id ? objectRect(ctx, id) : null;
  const [w, setW] = useState('');
  const [h, setH] = useState('');
  useEffect(() => {
    setW(r ? pxToInches(r.w) : '');
    setH(r ? pxToInches(r.h) : '');
  }, [r?.w, r?.h, id]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!id || !r) return null;
  const apply = () => {
    const nw = parseFloat(w);
    const nh = parseFloat(h);
    if (nw > 0 && nh > 0) setObjectSize(ctx, id, nw * 96, nh * 96);
  };
  const field = (label: string, v: string, set: (s: string) => void, testId: string) => (
    <label className="sh-field-inline">
      <span>{label}</span>
      <input inputMode="decimal" aria-label={label} data-testid={testId} value={v} onChange={(e) => set(e.target.value)} onBlur={apply} onKeyDown={(e) => e.key === 'Enter' && (e.target as HTMLInputElement).blur()} />
      <span>"</span>
    </label>
  );
  return (
    <div className="sh-size-fields">
      {field('Height:', h, setH, 'object-height')}
      {field('Width:', w, setW, 'object-width')}
    </div>
  );
}

/** Gridlines / Headings with View and Print check boxes, as in Excel. */
function SheetOptions({ ctx, columns }: { ctx: Ctx; columns: { head: string; items: Command[] }[] }) {
  return (
    <div className="sh-sheet-options">
      {columns.map((col) => (
        <div key={col.head} className="sh-sheet-options-col">
          <span className="sh-sheet-options-head">{col.head}</span>
          {col.items.map((c) => (
            <label key={c.id} className="sh-check">
              <input type="checkbox" data-testid={`rb-${c.id}`} checked={!!c.checked?.(ctx)} onChange={() => c.run?.(ctx)} />
              <span>{c.label}</span>
            </label>
          ))}
        </div>
      ))}
    </div>
  );
}

/** A ribbon control with fields of its own. */
export function RibbonView({ spec, ctx }: { spec: ViewSpec; ctx: Ctx }) {
  switch (spec.kind) {
    case 'table-name':
      return <TableName ctx={ctx} />;
    case 'object-size':
      return <ObjectSize ctx={ctx} />;
    case 'sheet-options':
      return <SheetOptions ctx={ctx} columns={spec.columns} />;
    case 'separator':
      return <span className="sh-tsep" />;
  }
}
