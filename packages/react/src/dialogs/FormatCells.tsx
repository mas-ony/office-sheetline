/**
 * Format Cells: Number (categories with options and a sample), Alignment, Font, Border (presets,
 * edges, line style and colour), Fill (colour, pattern) and Protection. Only what was changed is
 * applied to the selection.
 */
import { useMemo, useState } from 'react';
import { useUI } from '../context';
import { Check, ColorButton, Dialog, Field, ListBox, NumberField, Tabs } from './Dialog';
import { BORDER_LINE_PREVIEW as LINE_PREVIEW, BORDER_LINE_STYLES, type BorderEdge as Edge, type BorderEdges as Edges, type BorderLine, type BorderLineStyle, type CellStyle, type ColorSpec, commit, CURRENCY_SYMBOLS as SYMBOLS, CUSTOM_FORMATS, DATE_FORMAT_TYPES as DATE_TYPES, FILL_PATTERNS as PATTERNS, FONT_NAMES, FONT_SIZES, type FormatTab, formatValue, FRACTION_FORMAT_TYPES as FRACTION_TYPES, H_ALIGN_OPTIONS as H_ALIGNS, type HAlign, NUMBER_CATEGORIES as CATEGORIES, NUMBER_CATEGORY_DESCRIPTIONS as DESCRIPTIONS, numberCode, type NumberFormatState as NumberState, numberStateOf, resolveColor, SPECIAL_FORMAT_TYPES as SPECIAL_TYPES, type StylePatch, textRotationDegrees as degreesOf, TIME_FORMAT_TYPES as TIME_TYPES, UNDERLINE_OPTIONS as UNDERLINES, type UnderlineType, V_ALIGN_OPTIONS as V_ALIGNS, type VAlign } from '@sheetline/core/views';

// -----------------------------------------------------------------------------------------------
// Borders

function lineSvg(style: BorderLineStyle | 'none', color = '#000000', w = 70) {
  if (style === 'none') return <span className="sh-line-none">None</span>;
  const spec = LINE_PREVIEW[style];
  if (spec === 'd') {
    return (
      <svg width={w} height={10}>
        <line x1="2" y1="3.5" x2={w - 2} y2="3.5" stroke={color} strokeWidth="1" />
        <line x1="2" y1="6.5" x2={w - 2} y2="6.5" stroke={color} strokeWidth="1" />
      </svg>
    );
  }
  const [width, ...dash] = spec.split(' ').map(Number);
  return (
    <svg width={w} height={10}>
      <line x1="2" y1="5" x2={w - 2} y2="5" stroke={color} strokeWidth={width} strokeDasharray={dash.length ? dash.join(' ') : undefined} />
    </svg>
  );
}

export function FormatCellsDialog({ onClose, tab: initialTab = 'number' }: { onClose: () => void; tab?: FormatTab }) {
  const ui = useUI();
  const { doc } = ui;
  const { r, c } = doc.selection.active;
  const style: CellStyle = doc.styleAt(r, c);
  const theme = doc.wb.theme;
  const [tab, setTab] = useState<FormatTab>(initialTab);
  const [dirty, setDirty] = useState<Set<string>>(new Set());
  const touch = (k: string) => setDirty((d) => new Set(d).add(k));

  // Number
  const [num, setNum] = useState<NumberState>(() => numberStateOf(style.numFmt));
  const setN = (p: Partial<NumberState>) => {
    setNum((x) => ({ ...x, ...p }));
    touch('numFmt');
  };
  const code = numberCode(num);
  const cellValue = doc.engine.cellValue(doc.sheet, r, c);
  const sampleValue = typeof cellValue === 'number' ? cellValue : num.category === 'Date' || num.category === 'Time' ? 40982.5632 : -1234.1;
  const sample = useMemo(() => {
    try {
      const v = num.category === 'Text' && typeof cellValue === 'string' ? cellValue : sampleValue;
      return formatValue(v, code, doc.wb.date1904);
    } catch {
      return { text: String(sampleValue), color: undefined };
    }
  }, [code, cellValue, sampleValue, num.category, doc]);

  // Alignment
  const align = style.align ?? {};
  const [hAlign, setHAlign] = useState<HAlign>(align.horizontal ?? 'general');
  const [vAlign, setVAlign] = useState<VAlign>(align.vertical ?? 'bottom');
  const [indent, setIndent] = useState<number | ''>(align.indent ?? 0);
  const [wrap, setWrap] = useState(!!align.wrap);
  const [shrink, setShrink] = useState(!!align.shrink);
  const merged = ui.state.format.merged;
  const [merge, setMerge] = useState(merged);
  const [stacked, setStacked] = useState(align.rotation === 255);
  const [degrees, setDegrees] = useState<number | ''>(degreesOf(align.rotation));
  const [readingOrder, setReadingOrder] = useState(align.readingOrder ?? 0);

  // Font
  const font = style.font;
  const fontName = font.scheme === 'minor' ? theme.minorFont : font.scheme === 'major' ? theme.majorFont : font.name;
  const [fName, setFName] = useState(fontName);
  const [fStyle, setFStyle] = useState<'Regular' | 'Italic' | 'Bold' | 'Bold Italic'>(font.bold && font.italic ? 'Bold Italic' : font.bold ? 'Bold' : font.italic ? 'Italic' : 'Regular');
  const [fSize, setFSize] = useState(String(font.size));
  const [underline, setUnderline] = useState<UnderlineType | 'none'>(font.underline ?? 'none');
  const [fColor, setFColor] = useState<ColorSpec | null>(font.color ?? null);
  const [strike, setStrike] = useState(!!font.strike);
  const [vert, setVert] = useState<'superscript' | 'subscript' | ''>(font.vertAlign ?? '');

  // Border
  const [lineStyle, setLineStyle] = useState<BorderLineStyle>('thin');
  const [lineColor, setLineColor] = useState<ColorSpec | null>(null);
  const multiRow = doc.range.r1 > doc.range.r0;
  const multiCol = doc.range.c1 > doc.range.c0;
  const own = style.border ?? {};
  const [edges, setEdges] = useState<Edges>({});
  const edgeOn = (e: Edge): boolean => {
    if (edges[e] !== undefined) return !!edges[e];
    if (e === 'top') return !!own.top;
    if (e === 'bottom') return !!own.bottom;
    if (e === 'left') return !!own.left;
    if (e === 'right') return !!own.right;
    if (e === 'diagUp') return !!own.diagonalUp;
    if (e === 'diagDown') return !!own.diagonalDown;
    return false;
  };
  const line = (): BorderLine => ({ style: lineStyle, color: lineColor ?? undefined });
  const toggleEdge = (e: Edge) => setEdges((x) => ({ ...x, [e]: edgeOn(e) && (x[e] === undefined || x[e]?.style === lineStyle) ? null : line() }));
  const preset = (p: 'none' | 'outline' | 'inside') => {
    if (p === 'none') setEdges({ top: null, bottom: null, left: null, right: null, insideH: null, insideV: null, diagUp: null, diagDown: null });
    else if (p === 'outline') setEdges((x) => ({ ...x, top: line(), bottom: line(), left: line(), right: line() }));
    else setEdges((x) => ({ ...x, insideH: multiRow ? line() : x.insideH, insideV: multiCol ? line() : x.insideV }));
  };

  // Fill
  const fill = style.fill && !('gradientXml' in style.fill) ? style.fill : undefined;
  const [bg, setBg] = useState<ColorSpec | null>(fill && fill.pattern !== 'none' ? (fill.pattern === 'solid' ? fill.fg ?? null : fill.bg ?? null) : null);
  const [patternColor, setPatternColor] = useState<ColorSpec | null>(fill && fill.pattern !== 'solid' && fill.pattern !== 'none' ? fill.fg ?? null : null);
  const [pattern, setPattern] = useState(fill && fill.pattern !== 'solid' ? fill.pattern : 'none');

  // Protection
  const [locked, setLocked] = useState(style.locked !== false);
  const [hidden, setHidden] = useState(!!style.hidden);

  const apply = () => {
    commit(ui);
    const patch: StylePatch = {};
    if (dirty.has('numFmt')) patch.numFmt = code;
    if (dirty.has('align')) {
      const rot = stacked ? 255 : typeof degrees === 'number' && degrees !== 0 ? (degrees > 0 ? Math.min(90, degrees) : 90 + Math.min(90, -degrees)) : null;
      patch.align = {
        horizontal: hAlign === 'general' ? null : hAlign,
        vertical: vAlign === 'bottom' ? null : vAlign,
        indent: typeof indent === 'number' && indent > 0 && ['left', 'right', 'distributed'].includes(hAlign) ? indent : null,
        wrap: wrap || null,
        shrink: shrink || null,
        rotation: rot,
        readingOrder: readingOrder || null,
      };
    }
    if (dirty.has('font')) {
      const size = Number(fSize);
      patch.font = {
        name: fName,
        scheme: fName === fontName && font.scheme ? font.scheme : null,
        size: Number.isFinite(size) && size >= 1 && size <= 409 ? size : font.size,
        bold: fStyle.includes('Bold') || null,
        italic: fStyle.includes('Italic') || null,
        underline: underline === 'none' ? null : underline,
        color: fColor ?? { theme: 1 },
        strike: strike || null,
        vertAlign: vert || null,
      };
    }
    if (dirty.has('fill')) {
      if (!bg && pattern === 'none') patch.fill = null;
      else if (pattern === 'none' || pattern === 'solid') patch.fill = bg ? { pattern: 'solid', fg: bg, bg: { indexed: 64 } } : null;
      else patch.fill = { pattern, fg: patternColor ?? { rgb: 'FF000000' }, bg: bg ?? { rgb: 'FFFFFFFF' } };
    }
    if (dirty.has('protection')) {
      patch.locked = locked;
      patch.hidden = hidden;
    }
    const borderChanged = Object.keys(edges).length > 0;
    doc.transact('Format Cells', () => {
      if (Object.keys(patch).length) doc.format(patch);
      if (dirty.has('align') && merge !== merged) {
        if (merge) doc.merge(undefined, undefined, 'all');
        else doc.unmerge();
      }
      if (borderChanged) {
        const set = (kind: 'top' | 'bottom' | 'left' | 'right' | 'horizontal' | 'vertical', v: BorderLine | null | undefined) => {
          if (v !== undefined) doc.setBorder(kind, v);
        };
        set('top', edges.top);
        set('bottom', edges.bottom);
        set('left', edges.left);
        set('right', edges.right);
        set('horizontal', edges.insideH);
        set('vertical', edges.insideV);
        if (edges.diagUp !== undefined || edges.diagDown !== undefined) {
          const up = edges.diagUp !== undefined ? !!edges.diagUp : !!own.diagonalUp;
          const down = edges.diagDown !== undefined ? !!edges.diagDown : !!own.diagonalDown;
          const diag = edges.diagUp || edges.diagDown || own.diagonal;
          doc.format({ border: { diagonal: up || down ? diag ?? line() : null, diagonalUp: up || null, diagonalDown: down || null } }, undefined, undefined, 'Borders');
        }
      }
    });
  };

  const typeList = num.category === 'Date' ? DATE_TYPES : num.category === 'Time' ? TIME_TYPES : num.category === 'Fraction' ? FRACTION_TYPES : num.category === 'Special' ? SPECIAL_TYPES : null;
  const css = (c: ColorSpec | null | undefined, fb: string) => `#${resolveColor(c ?? undefined, theme, fb)}`;
  const previewEdge = (e: Edge) => {
    const v = edges[e];
    const ln = v === undefined ? (e === 'top' ? own.top : e === 'bottom' ? own.bottom : e === 'left' ? own.left : e === 'right' ? own.right : e === 'diagUp' || e === 'diagDown' ? own.diagonal : undefined) : v;
    return ln && edgeOn(e) ? ln : null;
  };
  const edgeStroke = (ln: BorderLine | null) => (ln ? { stroke: css(ln.color, '000000'), strokeWidth: ln.style === 'thick' ? 3 : ln.style.startsWith('medium') ? 2 : 1, strokeDasharray: /dash|dot|hair/i.test(ln.style) ? '4 2' : undefined } : { stroke: 'transparent' });

  return (
    <Dialog title="Format Cells" onClose={onClose} onOk={apply} width={620} testId="format-cells">
      <Tabs
        value={tab}
        onChange={setTab}
        testId="format-cells-tabs"
        tabs={[
          { id: 'number', label: 'Number' },
          { id: 'alignment', label: 'Alignment' },
          { id: 'font', label: 'Font' },
          { id: 'border', label: 'Border' },
          { id: 'fill', label: 'Fill' },
          { id: 'protection', label: 'Protection' },
        ]}
      />
      <div className="sh-tab-body" style={{ minHeight: 330 }}>
        {tab === 'number' && (
          <div className="sh-two-col">
            <div>
              <div className="sh-label">Category:</div>
              <ListBox
                items={CATEGORIES}
                value={num.category}
                keyOf={(x) => x}
                render={(x) => x}
                height={280}
                testId="number-category"
                label="Category"
                onChange={(cat) => {
                  const next: Partial<NumberState> = { category: cat };
                  if (cat === 'Date') next.code = DATE_TYPES[0].code;
                  if (cat === 'Time') next.code = TIME_TYPES[0].code;
                  if (cat === 'Fraction') next.code = FRACTION_TYPES[0].code;
                  if (cat === 'Special') next.code = SPECIAL_TYPES[0].code;
                  if (cat === 'Custom') next.code = code;
                  if (cat === 'Currency' || cat === 'Accounting') next.symbol = num.symbol || '"$"';
                  setN(next);
                }}
              />
            </div>
            <div className="sh-stack">
              <fieldset className="sh-group">
                <legend>Sample</legend>
                <div className="sh-sample" data-testid="number-sample" style={{ color: sample.color }}>{sample.text || ' '}</div>
              </fieldset>
              {['Number', 'Currency', 'Accounting', 'Percentage', 'Scientific'].includes(num.category) && (
                <Field label="Decimal places:">
                  <NumberField value={num.decimals} min={0} max={30} testId="number-decimals" label="Decimal places" onChange={(v) => setN({ decimals: typeof v === 'number' ? v : 0 })} />
                </Field>
              )}
              {num.category === 'Number' && <Check label="Use 1000 Separator (,)" checked={num.separator} testId="number-separator" onChange={(v) => setN({ separator: v })} />}
              {(num.category === 'Currency' || num.category === 'Accounting') && (
                <Field label="Symbol:">
                  <select value={num.symbol} data-testid="number-symbol" onChange={(e) => setN({ symbol: e.target.value })}>
                    {SYMBOLS.map((s) => (
                      <option key={s.label} value={s.code}>{s.label}</option>
                    ))}
                  </select>
                </Field>
              )}
              {(num.category === 'Number' || num.category === 'Currency') && (
                <>
                  <div className="sh-label">Negative numbers:</div>
                  <ListBox
                    items={[0, 1, 2, 3]}
                    value={num.negative}
                    keyOf={(x) => String(x)}
                    height={88}
                    label="Negative numbers"
                    render={(i) => {
                      const v = formatValue(-1234.1, numberCode({ ...num, negative: i }), doc.wb.date1904);
                      return <span style={{ color: v.color }}>{v.text}</span>;
                    }}
                    onChange={(i) => setN({ negative: i })}
                  />
                </>
              )}
              {typeList && (
                <>
                  <div className="sh-label">Type:</div>
                  <ListBox items={typeList} value={typeList.find((t) => t.code === num.code) ?? null} keyOf={(t) => t.code} render={(t) => t.label} height={170} testId="number-type" label="Type" onChange={(t) => setN({ code: t.code })} />
                </>
              )}
              {num.category === 'Custom' && (
                <>
                  <Field label="Type:" wide>
                    <input value={num.code} spellCheck={false} data-testid="number-custom" onChange={(e) => setN({ code: e.target.value })} />
                  </Field>
                  <ListBox items={[...new Set([style.numFmt, ...CUSTOM_FORMATS])]} value={num.code} keyOf={(x) => x} render={(x) => x} height={150} label="Format codes" onChange={(x) => setN({ code: x })} />
                </>
              )}
              <p className="sh-hint">{DESCRIPTIONS[num.category]}</p>
            </div>
          </div>
        )}
        {tab === 'alignment' && (
          <div className="sh-two-col wide-left">
            <div className="sh-stack">
              <fieldset className="sh-group">
                <legend>Text alignment</legend>
                <Field label="Horizontal:">
                  <select value={hAlign} data-testid="align-horizontal" onChange={(e) => { setHAlign(e.target.value as HAlign); touch('align'); }}>
                    {H_ALIGNS.map((a) => <option key={a.value} value={a.value}>{a.label}</option>)}
                  </select>
                </Field>
                <Field label="Indent:">
                  <NumberField value={indent} min={0} max={250} label="Indent" testId="align-indent" onChange={(v) => { setIndent(v); touch('align'); }} />
                </Field>
                <Field label="Vertical:">
                  <select value={vAlign} data-testid="align-vertical" onChange={(e) => { setVAlign(e.target.value as VAlign); touch('align'); }}>
                    {V_ALIGNS.map((a) => <option key={a.value} value={a.value}>{a.label}</option>)}
                  </select>
                </Field>
              </fieldset>
              <fieldset className="sh-group">
                <legend>Text control</legend>
                <Check label="Wrap text" checked={wrap} testId="align-wrap" onChange={(v) => { setWrap(v); if (v) setShrink(false); touch('align'); }} />
                <Check label="Shrink to fit" checked={shrink} disabled={wrap} testId="align-shrink" onChange={(v) => { setShrink(v); touch('align'); }} />
                <Check label="Merge cells" checked={merge} testId="align-merge" onChange={(v) => { setMerge(v); touch('align'); }} />
              </fieldset>
              <fieldset className="sh-group">
                <legend>Right-to-left</legend>
                <Field label="Text direction:">
                  <select value={readingOrder} onChange={(e) => { setReadingOrder(Number(e.target.value)); touch('align'); }}>
                    <option value={0}>Context</option>
                    <option value={1}>Left-to-Right</option>
                    <option value={2}>Right-to-Left</option>
                  </select>
                </Field>
              </fieldset>
            </div>
            <fieldset className="sh-group sh-orientation">
              <legend>Orientation</legend>
              <div className="sh-orient-row">
                <button type="button" className={`sh-orient-vertical${stacked ? ' active' : ''}`} aria-pressed={stacked} title="Vertical text" onClick={() => { setStacked(!stacked); touch('align'); }}>
                  {'Text'.split('').map((ch, i) => <span key={i}>{ch}</span>)}
                </button>
                <svg className="sh-orient-dial" width="110" height="110" viewBox="-55 -55 110 110" onClick={(e) => {
                  const rect = (e.currentTarget as SVGElement).getBoundingClientRect();
                  const x = e.clientX - rect.left - 55;
                  const y = rect.top + 55 - e.clientY;
                  if (x < 0) return;
                  setDegrees(Math.round((Math.atan2(y, x) * 180) / Math.PI));
                  setStacked(false);
                  touch('align');
                }}>
                  <path d="M0 -50 A50 50 0 0 1 0 50" fill="none" stroke="#bbb" />
                  {[-90, -45, 0, 45, 90].map((d) => <circle key={d} cx={Math.cos((d * Math.PI) / 180) * 50} cy={-Math.sin((d * Math.PI) / 180) * 50} r="2.5" fill="#888" />)}
                  {!stacked && <line x1="0" y1="0" x2={Math.cos(((Number(degrees) || 0) * Math.PI) / 180) * 46} y2={-Math.sin(((Number(degrees) || 0) * Math.PI) / 180) * 46} stroke="#217346" strokeWidth="2" />}
                  <text x="6" y="4" fontSize="11" fill="#333" transform={`rotate(${-(Number(degrees) || 0)} 0 0)`}>Text ——</text>
                </svg>
              </div>
              <Field label="Degrees:">
                <NumberField value={degrees} min={-90} max={90} label="Degrees" testId="align-degrees" disabled={stacked} onChange={(v) => { setDegrees(v); touch('align'); }} />
              </Field>
            </fieldset>
          </div>
        )}
        {tab === 'font' && (
          <div className="sh-stack">
            <div className="sh-font-grid">
              <div>
                <Field label="Font:" wide>
                  <input value={fName} data-testid="font-name-input" onChange={(e) => { setFName(e.target.value); touch('font'); }} />
                </Field>
                <ListBox items={[...new Set([theme.majorFont, theme.minorFont, ...FONT_NAMES])]} value={fName} keyOf={(x) => x} height={140} label="Fonts" render={(x) => <span style={{ fontFamily: `"${x}"` }}>{x === theme.majorFont ? `${x} (Headings)` : x === theme.minorFont ? `${x} (Body)` : x}</span>} onChange={(x) => { setFName(x); touch('font'); }} />
              </div>
              <div>
                <Field label="Font style:" wide>
                  <input value={fStyle} readOnly />
                </Field>
                <ListBox items={['Regular', 'Italic', 'Bold', 'Bold Italic'] as const} value={fStyle} keyOf={(x) => x} height={140} label="Font styles" testId="font-style" render={(x) => <span style={{ fontWeight: x.includes('Bold') ? 700 : 400, fontStyle: x.includes('Italic') ? 'italic' : 'normal' }}>{x}</span>} onChange={(x) => { setFStyle(x); touch('font'); }} />
              </div>
              <div>
                <Field label="Size:" wide>
                  <input value={fSize} inputMode="decimal" data-testid="font-size-input" onChange={(e) => { setFSize(e.target.value); touch('font'); }} />
                </Field>
                <ListBox items={FONT_SIZES.map(String)} value={fSize} keyOf={(x) => x} height={140} label="Sizes" render={(x) => x} onChange={(x) => { setFSize(x); touch('font'); }} />
              </div>
            </div>
            <div className="sh-row">
              <Field label="Underline:">
                <select value={underline} data-testid="font-underline" onChange={(e) => { setUnderline(e.target.value as UnderlineType | 'none'); touch('font'); }}>
                  {UNDERLINES.map((u) => <option key={u.value} value={u.value}>{u.label}</option>)}
                </select>
              </Field>
              <Field label="Color:">
                <ColorButton value={fColor} testId="font-color" noneSwatch="#000000" onChange={(v) => { setFColor(v); touch('font'); }} />
              </Field>
            </div>
            <div className="sh-row">
              <fieldset className="sh-group">
                <legend>Effects</legend>
                <Check label="Strikethrough" checked={strike} testId="font-strike" onChange={(v) => { setStrike(v); touch('font'); }} />
                <Check label="Superscript" checked={vert === 'superscript'} onChange={(v) => { setVert(v ? 'superscript' : ''); touch('font'); }} />
                <Check label="Subscript" checked={vert === 'subscript'} onChange={(v) => { setVert(v ? 'subscript' : ''); touch('font'); }} />
              </fieldset>
              <fieldset className="sh-group grow">
                <legend>Preview</legend>
                <div
                  className="sh-font-preview"
                  style={{
                    fontFamily: `"${fName}"`,
                    fontWeight: fStyle.includes('Bold') ? 700 : 400,
                    fontStyle: fStyle.includes('Italic') ? 'italic' : 'normal',
                    fontSize: `${Math.min(36, Number(fSize) || 11) * 1.333}px`,
                    color: css(fColor, '000000'),
                    textDecoration: `${underline !== 'none' ? 'underline' : ''} ${strike ? 'line-through' : ''}`.trim() || undefined,
                    textDecorationStyle: underline.toLowerCase().includes('double') ? 'double' : undefined,
                    verticalAlign: vert === 'superscript' ? 'super' : vert === 'subscript' ? 'sub' : undefined,
                  }}
                >
                  AaBbCcYyZz
                </div>
              </fieldset>
            </div>
          </div>
        )}
        {tab === 'border' && (
          <div className="sh-two-col">
            <div className="sh-stack">
              <fieldset className="sh-group">
                <legend>Line</legend>
                <div className="sh-label">Style:</div>
                <div className="sh-line-styles" role="listbox" aria-label="Line style" data-testid="border-styles">
                  {BORDER_LINE_STYLES.map((ls) => (
                    <button key={ls} type="button" role="option" aria-selected={lineStyle === ls} className={lineStyle === ls ? 'selected' : ''} title={ls} onClick={() => setLineStyle(ls)}>
                      {lineSvg(ls, css(lineColor, '000000'))}
                    </button>
                  ))}
                </div>
                <div className="sh-label">Color:</div>
                <ColorButton value={lineColor} testId="border-color" noneSwatch="#000000" onChange={setLineColor} />
              </fieldset>
            </div>
            <div className="sh-stack">
              <fieldset className="sh-group">
                <legend>Presets</legend>
                <div className="sh-border-presets">
                  <button type="button" data-testid="border-preset-none" onClick={() => preset('none')}>
                    <svg width="34" height="34"><rect x="4" y="4" width="26" height="26" fill="none" stroke="#c8c8c8" /></svg>
                    <span>None</span>
                  </button>
                  <button type="button" data-testid="border-preset-outline" onClick={() => preset('outline')}>
                    <svg width="34" height="34"><rect x="4" y="4" width="26" height="26" fill="none" stroke="#333" strokeWidth="2" /></svg>
                    <span>Outline</span>
                  </button>
                  <button type="button" data-testid="border-preset-inside" disabled={!multiRow && !multiCol} onClick={() => preset('inside')}>
                    <svg width="34" height="34"><rect x="4" y="4" width="26" height="26" fill="none" stroke="#c8c8c8" /><path d="M17 4 V30 M4 17 H30" stroke="#333" strokeWidth="2" /></svg>
                    <span>Inside</span>
                  </button>
                </div>
              </fieldset>
              <fieldset className="sh-group">
                <legend>Border</legend>
                <div className="sh-border-editor">
                  <div className="sh-border-buttons left">
                    {(['top', 'insideH', 'bottom'] as Edge[]).map((e) => (
                      <button key={e} type="button" aria-pressed={edgeOn(e)} className={edgeOn(e) ? 'on' : ''} disabled={e === 'insideH' && !multiRow} title={e} data-testid={`border-${e}`} onClick={() => toggleEdge(e)}>
                        {e === 'top' ? '▔' : e === 'bottom' ? '▁' : '═'}
                      </button>
                    ))}
                  </div>
                  <svg className="sh-border-preview" width="150" height="110" viewBox="0 0 150 110">
                    <rect x="10" y="10" width="130" height="90" fill="#fff" />
                    <text x="40" y="38" fontSize="12" fill="#999">Text</text>
                    {multiCol && <text x="100" y="38" fontSize="12" fill="#999">Text</text>}
                    {multiRow && <text x="40" y="82" fontSize="12" fill="#999">Text</text>}
                    {multiRow && multiCol && <text x="100" y="82" fontSize="12" fill="#999">Text</text>}
                    <line x1="10" y1="10" x2="140" y2="10" {...edgeStroke(previewEdge('top'))} />
                    <line x1="10" y1="100" x2="140" y2="100" {...edgeStroke(previewEdge('bottom'))} />
                    <line x1="10" y1="10" x2="10" y2="100" {...edgeStroke(previewEdge('left'))} />
                    <line x1="140" y1="10" x2="140" y2="100" {...edgeStroke(previewEdge('right'))} />
                    {multiRow && <line x1="10" y1="55" x2="140" y2="55" {...edgeStroke(previewEdge('insideH'))} />}
                    {multiCol && <line x1="75" y1="10" x2="75" y2="100" {...edgeStroke(previewEdge('insideV'))} />}
                    <line x1="10" y1="100" x2="140" y2="10" {...edgeStroke(previewEdge('diagUp'))} />
                    <line x1="10" y1="10" x2="140" y2="100" {...edgeStroke(previewEdge('diagDown'))} />
                    <rect x="10" y="0" width="130" height="14" fill="transparent" onClick={() => toggleEdge('top')} />
                    <rect x="10" y="96" width="130" height="14" fill="transparent" onClick={() => toggleEdge('bottom')} />
                    <rect x="0" y="10" width="14" height="90" fill="transparent" onClick={() => toggleEdge('left')} />
                    <rect x="136" y="10" width="14" height="90" fill="transparent" onClick={() => toggleEdge('right')} />
                  </svg>
                  <div className="sh-border-buttons bottom">
                    {(['diagUp', 'left', 'insideV', 'right', 'diagDown'] as Edge[]).map((e) => (
                      <button key={e} type="button" aria-pressed={edgeOn(e)} className={edgeOn(e) ? 'on' : ''} disabled={e === 'insideV' && !multiCol} title={e} data-testid={`border-${e}`} onClick={() => toggleEdge(e)}>
                        {e === 'diagUp' ? '╱' : e === 'diagDown' ? '╲' : e === 'insideV' ? '║' : e === 'left' ? '▏' : '▕'}
                      </button>
                    ))}
                  </div>
                </div>
              </fieldset>
              <p className="sh-hint">The selected border style can be applied by clicking the presets, preview diagram or the buttons above.</p>
            </div>
          </div>
        )}
        {tab === 'fill' && (
          <div className="sh-two-col">
            <div className="sh-stack">
              <div className="sh-label">Background Color:</div>
              <ColorButton value={bg} noneLabel="No Color" testId="fill-color" onChange={(v) => { setBg(v); if (!v && pattern === 'solid') setPattern('none'); touch('fill'); }} />
            </div>
            <div className="sh-stack">
              <Field label="Pattern Color:">
                <ColorButton value={patternColor} noneLabel="Automatic" noneSwatch="#000000" onChange={(v) => { setPatternColor(v); touch('fill'); }} />
              </Field>
              <Field label="Pattern Style:">
                <select value={pattern} data-testid="fill-pattern" onChange={(e) => { setPattern(e.target.value); touch('fill'); }}>
                  {PATTERNS.map((p) => <option key={p} value={p}>{p === 'none' ? 'None' : p.replace(/([A-Z])/g, ' $1').replace(/^./, (m) => m.toUpperCase())}</option>)}
                </select>
              </Field>
              <fieldset className="sh-group">
                <legend>Sample</legend>
                <div className="sh-fill-sample" data-testid="fill-sample" style={{ background: pattern === 'none' || pattern === 'solid' ? (bg ? css(bg, 'FFFFFF') : '#fff') : `repeating-linear-gradient(45deg, ${css(patternColor, '000000')} 0 2px, ${bg ? css(bg, 'FFFFFF') : '#fff'} 2px 5px)` }} />
              </fieldset>
            </div>
          </div>
        )}
        {tab === 'protection' && (
          <div className="sh-stack">
            <Check label="Locked" checked={locked} testId="protection-locked" onChange={(v) => { setLocked(v); touch('protection'); }} />
            <Check label="Hidden" checked={hidden} testId="protection-hidden" onChange={(v) => { setHidden(v); touch('protection'); }} />
            <p className="sh-hint">Locking cells or hiding formulas has no effect until you protect the worksheet (Review tab, Protect group, Protect Sheet button).</p>
          </div>
        )}
      </div>
    </Dialog>
  );
}
