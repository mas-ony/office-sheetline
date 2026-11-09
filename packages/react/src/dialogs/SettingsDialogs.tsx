/**
 * Settings and tools: Spread Settings, Sheet Settings, Customize Theme Colors, New Cell Style, the
 * Command Palette and the keyboard shortcut list.
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import { usePersistent, useUI } from '../context';
import { Check, ColorButton, Dialog, Field, Tabs } from './Dialog';
import { ColorField } from '../ColorField';
import { type ColorSpec, KEYBOARD_SHORTCUTS as SHORTCUTS, type PaletteEntry, paletteEntries, resolveColor, type StylePatch, THEME_SLOT_LABELS, type ThemeSlot } from '@sheetline/core/views';
import { Icon } from '../Icon';

// -----------------------------------------------------------------------------------------------
// Spread Settings

export function SpreadSettingsDialog({ tab: initial = 'general', onClose }: { tab?: string; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const [tab, setTab] = useState(initial);
  const [s, setS] = useState(ui.settings);
  const [calc, setCalc] = useState(doc.wb.calcMode ?? 'auto');
  const [tabStrip, setTabStrip] = useState(ui.view.tabStrip);
  const [newTab, setNewTab] = useState(ui.view.newTab);
  const [formulaBar, setFormulaBar] = useState(ui.view.formulaBar);
  return (
    <Dialog
      title="Spread Settings"
      onClose={onClose}
      width={440}
      testId="spread-settings"
      onOk={() => {
        ui.setSettings(s);
        ui.setView({ tabStrip, newTab, formulaBar });
        if (calc !== (doc.wb.calcMode ?? 'auto')) doc.setCalcMode(calc);
      }}
    >
      <Tabs value={tab} onChange={setTab} tabs={[{ id: 'general', label: 'General' }, { id: 'scrollbars', label: 'Scroll Bars' }, { id: 'calculation', label: 'Calculation' }, { id: 'tabstrip', label: 'Tab Strip' }]} testId="spread-settings-tabs" />
      <div className="sh-tab-body">
        {tab === 'general' && (
          <div className="sh-stack">
            <Check label="Allow Drag and Fill" checked={s.allowFill} onChange={(v) => setS({ ...s, allowFill: v })} testId="setting-fill" />
            <Check label="Allow Zoom" checked={s.allowZoom} onChange={(v) => setS({ ...s, allowZoom: v })} testId="setting-zoom" />
            <Check label="Formula Format Hints" checked={s.formulaHints} onChange={(v) => setS({ ...s, formulaHints: v })} />
            <Check label="Show Formula Bar" checked={formulaBar} onChange={setFormulaBar} />
          </div>
        )}
        {tab === 'scrollbars' && (
          <div className="sh-stack">
            <Check label="Show Horizontal Scroll Bar" checked={s.showHScroll} onChange={(v) => setS({ ...s, showHScroll: v })} testId="setting-hscroll" />
            <Check label="Show Vertical Scroll Bar" checked={s.showVScroll} onChange={(v) => setS({ ...s, showVScroll: v })} />
          </div>
        )}
        {tab === 'calculation' && (
          <div className="sh-stack">
            <Field label="Calculation:">
              <select value={calc} data-testid="setting-calc" onChange={(e) => setCalc(e.target.value as typeof calc)}>
                <option value="auto">Automatic</option>
                <option value="autoNoTable">Automatic except for data tables</option>
                <option value="manual">Manual</option>
              </select>
            </Field>
            <p className="sh-hint">With Manual, formulas recalculate when you press F9 (Calculate Now) or Shift+F9 (Calculate Sheet).</p>
          </div>
        )}
        {tab === 'tabstrip' && (
          <div className="sh-stack">
            <Check label="Show Tab Strip" checked={tabStrip} onChange={setTabStrip} testId="setting-tabstrip" />
            <Check label="Show New Tab" checked={newTab} onChange={setNewTab} />
          </div>
        )}
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Sheet Settings

export function SheetSettingsDialog({ tab: initial = 'general', onClose }: { tab?: string; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const sheet = doc.sheet;
  const [tab, setTab] = useState(initial);
  const [frozenRows, setFrozenRows] = useState(String(sheet.freeze.rows));
  const [frozenCols, setFrozenCols] = useState(String(sheet.freeze.cols));
  const [zeros, setZeros] = useState(sheet.showZeros);
  const [gridH, setGridH] = useState(sheet.showGridLines && sheet.gridH);
  const [gridV, setGridV] = useState(sheet.showGridLines && sheet.gridV);
  const [gridColor, setGridColor] = useState(sheet.gridColor ?? '');
  const [rowHeaders, setRowHeaders] = useState(sheet.showRowHeaders);
  const [colHeaders, setColHeaders] = useState(sheet.showColHeaders);
  const [tabColor, setTabColor] = useState<ColorSpec | null>(sheet.tabColor ?? null);
  return (
    <Dialog
      title="Sheet Settings"
      onClose={onClose}
      width={420}
      testId="sheet-settings"
      onOk={() => {
        const fr = Math.max(0, Number(frozenRows) || 0);
        const fc = Math.max(0, Number(frozenCols) || 0);
        if (fr !== sheet.freeze.rows || fc !== sheet.freeze.cols) doc.setFreeze(fr, fc);
        doc.setSheetView({ showZeros: zeros, gridH, gridV, showGridLines: gridH || gridV, gridColor: gridColor || undefined, showRowHeaders: rowHeaders, showColHeaders: colHeaders });
        if (JSON.stringify(tabColor) !== JSON.stringify(sheet.tabColor ?? null)) doc.setTabColorSpec(doc.activeIndex, tabColor);
      }}
    >
      <Tabs value={tab} onChange={setTab} tabs={[{ id: 'general', label: 'General' }, { id: 'gridlines', label: 'Gridlines' }, { id: 'headers', label: 'Headers' }, { id: 'sheettab', label: 'Sheet Tab' }]} testId="sheet-settings-tabs" />
      <div className="sh-tab-body">
        {tab === 'general' && (
          <div className="sh-stack">
            <Field label="Frozen rows:">
              <input value={frozenRows} inputMode="numeric" style={{ width: 70 }} data-testid="setting-frozen-rows" onChange={(e) => setFrozenRows(e.target.value)} />
            </Field>
            <Field label="Frozen columns:">
              <input value={frozenCols} inputMode="numeric" style={{ width: 70 }} onChange={(e) => setFrozenCols(e.target.value)} />
            </Field>
            <Check label="Show zero values" checked={zeros} onChange={setZeros} testId="setting-zeros" />
          </div>
        )}
        {tab === 'gridlines' && (
          <div className="sh-stack">
            <Check label="Show horizontal gridlines" checked={gridH} onChange={setGridH} />
            <Check label="Show vertical gridlines" checked={gridV} onChange={setGridV} />
            <Field label="Gridline color:">
              <ColorField value={gridColor} onChange={setGridColor} placeholder="Automatic" />
            </Field>
          </div>
        )}
        {tab === 'headers' && (
          <div className="sh-stack">
            <Check label="Show row header" checked={rowHeaders} onChange={setRowHeaders} testId="setting-row-header" />
            <Check label="Show column header" checked={colHeaders} onChange={setColHeaders} />
          </div>
        )}
        {tab === 'sheettab' && (
          <div className="sh-stack">
            <Field label="Tab color:">
              <ColorButton value={tabColor} noneLabel="No Color" onChange={setTabColor} testId="setting-tab-color" />
            </Field>
          </div>
        )}
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Customize theme colors

export function CustomizeColorsDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const theme = ui.doc.wb.theme;
  const [colors, setColors] = useState<Record<ThemeSlot, string>>({ ...theme.colors });
  const [name, setName] = useState('Custom 1');
  return (
    <Dialog
      title="Create New Theme Colors"
      onClose={onClose}
      width={480}
      testId="customize-colors"
      onOk={() => ui.doc.applyTheme({ name, colors })}
      footerLeft={<button type="button" className="sh-btn" onClick={() => setColors({ ...theme.colors })}>Reset</button>}
    >
      <div className="sh-two-col wide-left">
        <div className="sh-stack tight">
          <div className="sh-label">Theme colors</div>
          {THEME_SLOT_LABELS.map((s) => (
            <div key={s.slot} className="sh-row">
              <span className="grow">{s.label}</span>
              <ColorField value={`#${colors[s.slot]}`} onChange={(hex) => hex && setColors({ ...colors, [s.slot]: hex.replace('#', '').toUpperCase() })} />
            </div>
          ))}
        </div>
        <div className="sh-stack">
          <div className="sh-label">Sample</div>
          <div className="sh-theme-sample" style={{ background: `#${colors.lt1}`, color: `#${colors.dk1}` }}>
            <div style={{ background: `#${colors.dk2}`, color: `#${colors.lt1}` }}>Text</div>
            <div className="bars">
              {(['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as ThemeSlot[]).map((k) => <i key={k} style={{ background: `#${colors[k]}` }} />)}
            </div>
            <u style={{ color: `#${colors.hlink}` }}>Hyperlink</u> <u style={{ color: `#${colors.folHlink}` }}>Followed</u>
          </div>
        </div>
      </div>
      <Field label="Name:" wide>
        <input value={name} onChange={(e) => setName(e.target.value)} />
      </Field>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// New Cell Style

export interface CustomCellStyle {
  name: string;
  patch: StylePatch;
}

export function NewCellStyleDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [styles, setStyles] = usePersistent<CustomCellStyle[]>('custom-cell-styles', []);
  const [name, setName] = useState(`Style ${styles.length + 1}`);
  const { r, c } = ui.doc.selection.active;
  const style = ui.doc.styleAt(r, c);
  const [parts, setParts] = useState({ number: true, alignment: true, font: true, border: true, fill: true, protection: true });
  const theme = ui.doc.wb.theme;
  const describe = {
    number: style.numFmt,
    alignment: `${style.align?.horizontal ?? 'General'}, ${style.align?.vertical ?? 'Bottom'} Aligned`,
    font: `${style.font.name} ${style.font.size}${style.font.bold ? ', Bold' : ''}${style.font.italic ? ', Italic' : ''}, ${style.font.color ? `#${resolveColor(style.font.color, theme, '000000')}` : 'Text 1'}`,
    border: style.border ? 'Borders' : 'No Borders',
    fill: style.fill ? 'Shaded' : 'No Shading',
    protection: style.locked === false ? 'Unlocked' : 'Locked',
  };
  return (
    <Dialog
      title="Style"
      onClose={onClose}
      width={400}
      testId="new-cell-style"
      onOk={() => {
        if (!name.trim()) return false;
        const patch: StylePatch = {};
        if (parts.number) patch.numFmt = style.numFmt;
        if (parts.alignment) patch.align = { ...(style.align ?? {}) };
        if (parts.font) patch.font = { ...style.font };
        if (parts.border) patch.border = style.border ? { ...style.border } : null;
        if (parts.fill) patch.fill = style.fill ?? null;
        if (parts.protection) patch.locked = style.locked !== false;
        setStyles([...styles.filter((s) => s.name !== name.trim()), { name: name.trim(), patch }]);
        return true;
      }}
    >
      <Field label="Style name:" wide>
        <input value={name} autoFocus data-testid="cell-style-name" onChange={(e) => setName(e.target.value)} />
      </Field>
      <div className="sh-label">Style Includes (By Example)</div>
      {(Object.keys(parts) as (keyof typeof parts)[]).map((k) => (
        <div key={k} className="sh-row">
          <Check label={k[0].toUpperCase() + k.slice(1)} checked={parts[k]} onChange={(v) => setParts({ ...parts, [k]: v })} />
          <span className="sh-hint">{describe[k]}</span>
        </div>
      ))}
      <p className="sh-hint">Custom styles are kept in this browser and appear under Custom in the Cell Styles gallery.</p>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Command Palette

export function CommandPalette({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [query, setQuery] = useState('');
  const [recent, setRecent] = usePersistent<string[]>('recent-commands', []);
  const [index, setIndex] = useState(0);
  const list = useMemo(() => paletteEntries(ui), [ui]);
  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) {
      const rec = recent.map((id) => list.find((c) => c.id === id)).filter((c): c is PaletteEntry => !!c);
      // Suggestions: plain commands, not every menu item.
      return { recent: rec, rest: list.filter((c) => !recent.includes(c.id) && !c.id.includes('/')).slice(0, 40) };
    }
    const score = (c: PaletteEntry) => {
      const last = c.label.split(' › ').pop()!.toLowerCase();
      if (c.label.toLowerCase().startsWith(q) || last.startsWith(q)) return 0;
      if (c.hay.includes(q)) return 1;
      return q.split(/\s+/).every((w) => c.hay.includes(w)) ? 2 : 9;
    };
    return { recent: [], rest: list.map((c) => ({ c, s: score(c) })).filter((x) => x.s < 9).sort((a, b) => a.s - b.s).map((x) => x.c).slice(0, 60) };
  }, [query, list, recent]);
  const flat = [...results.recent, ...results.rest];
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => input.current?.focus(), []);
  const run = (c: PaletteEntry) => {
    onClose();
    setRecent([c.id, ...recent.filter((x) => x !== c.id)].slice(0, 8));
    window.setTimeout(() => c.run(), 0);
  };
  const row = (c: PaletteEntry, i: number) => (
    <div key={c.id} role="option" aria-selected={i === index} className={`sh-palette-item${i === index ? ' selected' : ''}`} data-testid={`palette-${c.id}`} onMouseEnter={() => setIndex(i)} onClick={() => run(c)}>
      {c.icon ? <Icon name={c.icon} /> : <span className="sh-icon-space" />}
      <span className="grow">{c.label}</span>
      {c.shortcut && <kbd>{c.shortcut}</kbd>}
    </div>
  );
  return (
    <Dialog title="Command Palette" onClose={onClose} closeOnly width={520} testId="command-palette">
      <input
        ref={input}
        className="sh-palette-search"
        placeholder="Search for actions, text, and more"
        value={query}
        data-testid="palette-search"
        onChange={(e) => {
          setQuery(e.target.value);
          setIndex(0);
        }}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown') {
            e.preventDefault();
            setIndex(Math.min(flat.length - 1, index + 1));
          } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            setIndex(Math.max(0, index - 1));
          } else if (e.key === 'Enter' && flat[index]) {
            e.preventDefault();
            run(flat[index]);
          }
        }}
      />
      <div className="sh-palette-list" role="listbox" aria-label="Commands">
        {results.recent.length > 0 && <div className="sh-panel-title">Recently Used Actions</div>}
        {results.recent.map((c, i) => row(c, i))}
        {!query && <div className="sh-panel-title">Suggested Actions</div>}
        {results.rest.map((c, i) => row(c, i + results.recent.length))}
        {!flat.length && <div className="sh-hint">No matching actions.</div>}
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// Keyboard shortcuts

export function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  return (
    <Dialog title="Keyboard Shortcuts" onClose={onClose} closeOnly width={480} testId="shortcuts">
      <div className="sh-shortcuts">
        {SHORTCUTS.map(([k, v]) => (
          <div key={k} className="sh-kv">
            <kbd>{k}</kbd>
            <span>{v}</span>
          </div>
        ))}
      </div>
    </Dialog>
  );
}

