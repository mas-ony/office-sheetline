/**
 * Conditional formatting dialogs: the quick rules (Greater Than..., Top 10 Items...), New / Edit
 * Formatting Rule (six rule types) and the Conditional Formatting Rules Manager.
 */
import { useMemo, useState } from 'react';
import { useUI } from '../context';
import { Check, ColorButton, Dialog, Field, NumberField, RangeField, parseRangeText, rangeText } from './Dialog';
import { CELL_IS_OPERATORS as CELL_OPS, type CellIsOperator, type CfContainType as ContainType, type CfFormat, cfHex as hex, cfIconUrl as iconUrl, cfInitialRule, cfOperand as operand, CF_PRESET_TEXT as PRESET_TEXT, type CfPresetKind, type CfRule, CF_RULE_KINDS as RULE_KINDS, type CfRuleKind as RuleKind, cfRuleKindOf as kindOf, cfSpec as spec, type Cfvo, CFVO_TYPE_OPTIONS as CFVO_TYPES, type CfvoType, commit, DATA_BAR_COLORS, defaultThresholds, describeRule, HIGHLIGHT_PRESETS, ICON_SET_GROUPS, iconCount, type IconSetName, type Range, selectionAverage, TEXT_OPERATOR_OPTIONS as TEXT_OPS, type TextOperator, TIME_PERIODS, type TimePeriod, type XmlElement } from '@sheetline/core/views';

// -----------------------------------------------------------------------------------------------
// Shared pieces

/** How a format looks (preview cell of the dialogs and the manager). */
export function FormatPreview({ format, text = 'AaBbCcYyZz' }: { format: CfFormat | null; text?: string }) {
  if (!format) return <span className="sh-cf-preview none">No Format Set</span>;
  return (
    <span
      className="sh-cf-preview"
      style={{
        color: format.color,
        background: format.fill,
        fontWeight: format.bold ? 700 : undefined,
        fontStyle: format.italic ? 'italic' : undefined,
        textDecoration: [format.underline ? 'underline' : '', format.strike ? 'line-through' : ''].join(' ').trim() || undefined,
        outline: format.border ? `1px solid ${format.border}` : undefined,
      }}
    >
      {text}
    </span>
  );
}

function RulePreview({ rule }: { rule: CfRule | null }) {
  if (!rule) return <span className="sh-cf-preview none">Unknown rule</span>;
  if (rule.type === 'colorScale') return <span className="sh-cf-preview scale" style={{ background: `linear-gradient(90deg, ${rule.points.map((p) => p.color).join(', ')})` }} />;
  if (rule.type === 'dataBar') return <span className="sh-cf-preview bar" style={{ background: rule.gradient ? `linear-gradient(90deg, ${rule.color} 0, #fff 70%)` : rule.color, width: 60 }} />;
  if (rule.type === 'iconSet') {
    const n = iconCount(rule.iconSet);
    return (
      <span className="sh-cf-preview icons">
        {Array.from({ length: n }, (_, i) => (
          <img key={i} alt="" src={iconUrl(rule.iconSet, rule.reverse ? i : n - 1 - i)} width={16} height={16} />
        ))}
      </span>
    );
  }
  return <FormatPreview format={rule.format} />;
}

/** The format part of a rule: a preset list plus "Custom Format..." (font, border, fill, number format). */
function FormatChooser({ value, onChange, testId }: { value: CfFormat; onChange: (f: CfFormat) => void; testId?: string }) {
  const [custom, setCustom] = useState(false);
  const preset = HIGHLIGHT_PRESETS.findIndex((p) => JSON.stringify(p.format) === JSON.stringify(value));
  return (
    <>
      <select
        data-testid={testId}
        value={custom || preset < 0 ? 'custom' : String(preset)}
        onChange={(e) => {
          if (e.target.value === 'custom') setCustom(true);
          else {
            setCustom(false);
            onChange({ ...HIGHLIGHT_PRESETS[Number(e.target.value)].format });
          }
        }}
      >
        {HIGHLIGHT_PRESETS.map((p, i) => (
          <option key={p.label} value={i}>{p.label}</option>
        ))}
        <option value="custom">Custom Format...</option>
      </select>
      {(custom || preset < 0) && <FormatEditor value={value} onChange={onChange} />}
    </>
  );
}

/** Compact Format Cells for rules: font style and colour, underline, strikethrough, border, fill, number format. */
function FormatEditor({ value, onChange }: { value: CfFormat; onChange: (f: CfFormat) => void }) {
  const ui = useUI();
  const theme = ui.doc.wb.theme;
  const set = (p: Partial<CfFormat>) => onChange({ ...value, ...p });
  return (
    <fieldset className="sh-group sh-cf-format">
      <legend>Format</legend>
      <div className="sh-row wrap">
        <Check label="Bold" checked={!!value.bold} onChange={(v) => set({ bold: v || undefined })} />
        <Check label="Italic" checked={!!value.italic} onChange={(v) => set({ italic: v || undefined })} />
        <Check label="Underline" checked={!!value.underline} onChange={(v) => set({ underline: v || undefined })} />
        <Check label="Strikethrough" checked={!!value.strike} onChange={(v) => set({ strike: v || undefined })} />
      </div>
      <div className="sh-row wrap">
        <Field label="Font color:">
          <ColorButton value={spec(value.color)} noneLabel="Automatic" noneSwatch="#000000" onChange={(c) => set({ color: hex(c, theme) })} />
        </Field>
        <Field label="Fill:">
          <ColorButton value={spec(value.fill)} noneLabel="No Color" onChange={(c) => set({ fill: hex(c, theme) })} />
        </Field>
        <Field label="Border:">
          <ColorButton value={spec(value.border)} noneLabel="None" onChange={(c) => set({ border: hex(c, theme) })} />
        </Field>
      </div>
      <Field label="Number format:">
        <input value={value.numFmt ?? ''} placeholder="(unchanged)" spellCheck={false} onChange={(e) => set({ numFmt: e.target.value || undefined })} />
      </Field>
      <div className="sh-row">
        <span className="sh-label">Preview:</span>
        <FormatPreview format={value} />
      </div>
    </fieldset>
  );
}

// -----------------------------------------------------------------------------------------------
// Quick rules

export function CfPresetDialog({ preset, onClose }: { preset: CfPresetKind; onClose: () => void }) {
  const ui = useUI();
  const avg = useMemo(() => selectionAverage(ui.doc), []); // eslint-disable-line react-hooks/exhaustive-deps
  const [v1, setV1] = useState(String(avg));
  const [v2, setV2] = useState(String(avg));
  const [text, setText] = useState('');
  const [period, setPeriod] = useState<TimePeriod>('yesterday');
  const [dup, setDup] = useState<'duplicateValues' | 'uniqueValues'>('duplicateValues');
  const [rank, setRank] = useState<number | ''>(10);
  const [format, setFormat] = useState<CfFormat>({ ...HIGHLIGHT_PRESETS[0].format });
  const meta = PRESET_TEXT[preset];

  const ok = () => {
    commit(ui);
    let rule: CfRule;
    switch (preset) {
      case 'greater':
      case 'less':
      case 'equal':
        rule = { type: 'cellIs', operator: preset === 'greater' ? 'greaterThan' : preset === 'less' ? 'lessThan' : 'equal', formula1: operand(v1), format };
        break;
      case 'between':
        rule = { type: 'cellIs', operator: 'between', formula1: operand(v1), formula2: operand(v2), format };
        break;
      case 'text':
        if (!text) {
          ui.notify('Enter the text to look for.');
          return false;
        }
        rule = { type: 'text', operator: 'containsText', text, format };
        break;
      case 'date':
        rule = { type: 'timePeriod', period, format };
        break;
      case 'duplicate':
        rule = { type: dup, format };
        break;
      case 'above':
      case 'below':
        rule = { type: 'aboveAverage', above: preset === 'above', equal: false, format };
        break;
      default: {
        const n = typeof rank === 'number' ? rank : 10;
        if (n < 1 || n > (preset.endsWith('pct') ? 100 : 1000)) {
          ui.notify(preset.endsWith('pct') ? 'Enter a whole number between 1 and 100.' : 'Enter a whole number between 1 and 1000.');
          return false;
        }
        rule = { type: 'top10', bottom: preset.startsWith('bottom'), percent: preset.endsWith('pct'), rank: n, format };
      }
    }
    ui.doc.addCondFormat(rule);
    return true;
  };

  return (
    <Dialog title={meta.title} onClose={onClose} onOk={ok} width={460} blocking={false} testId="cf-preset">
      <p className="sh-lead">{meta.lead}</p>
      <div className="sh-row wrap">
        {(preset === 'greater' || preset === 'less' || preset === 'equal' || preset === 'between') && <RangeField value={v1} onChange={setV1} label="Value" testId="cf-value1" />}
        {preset === 'between' && (
          <>
            <span>and</span>
            <RangeField value={v2} onChange={setV2} label="Second value" testId="cf-value2" />
          </>
        )}
        {preset === 'text' && <RangeField value={text} onChange={setText} label="Text" testId="cf-text" />}
        {preset === 'date' && (
          <select value={period} data-testid="cf-period" onChange={(e) => setPeriod(e.target.value as TimePeriod)}>
            {TIME_PERIODS.map((p) => <option key={p.period} value={p.period}>{p.label}</option>)}
          </select>
        )}
        {preset === 'duplicate' && (
          <select value={dup} data-testid="cf-dup" onChange={(e) => setDup(e.target.value as typeof dup)}>
            <option value="duplicateValues">Duplicate</option>
            <option value="uniqueValues">Unique</option>
          </select>
        )}
        {preset.startsWith('top') || preset.startsWith('bottom') ? (
          <>
            <NumberField value={rank} min={1} max={preset.endsWith('pct') ? 100 : 1000} label="Rank" testId="cf-rank" onChange={setRank} />
            <span>{preset.endsWith('pct') ? '%' : ''}</span>
          </>
        ) : null}
        {(preset === 'above' || preset === 'below') && <span>for the selected range</span>}
        <span>with</span>
        <FormatChooser value={format} onChange={setFormat} testId="cf-format" />
      </div>
    </Dialog>
  );
}

// -----------------------------------------------------------------------------------------------
// New / Edit Formatting Rule

function CfvoEditor({ value, onChange, which }: { value: Cfvo; onChange: (v: Cfvo) => void; which: 'min' | 'mid' | 'max' }) {
  const types = CFVO_TYPES.filter((t) => (which === 'min' ? t.type !== 'max' : which === 'max' ? t.type !== 'min' : t.type !== 'min' && t.type !== 'max'));
  return (
    <div className="sh-stack tight">
      <select value={value.type} onChange={(e) => onChange({ type: e.target.value as CfvoType, value: e.target.value === 'min' || e.target.value === 'max' ? undefined : value.value ?? (e.target.value === 'percentile' || e.target.value === 'percent' ? '50' : '0') })}>
        {types.map((t) => <option key={t.type} value={t.type}>{t.label}</option>)}
      </select>
      <input value={value.type === 'min' || value.type === 'max' ? (value.type === 'min' ? '(Lowest value)' : '(Highest value)') : value.value ?? ''} disabled={value.type === 'min' || value.type === 'max'} onChange={(e) => onChange({ ...value, value: e.target.value })} />
    </div>
  );
}

export function RuleEditorDialog({ initial, hint, title, onClose, onSave }: { initial: CfRule | null; hint?: string; title: string; onClose: () => void; onSave: (rule: CfRule) => void }) {
  const fmt0: CfFormat = initial && 'format' in initial ? { ...initial.format } : {};
  const [kind, setKind] = useState<RuleKind>(kindOf(initial, hint));
  const [format, setFormat] = useState<CfFormat>(fmt0);
  // Values
  const [style, setStyle] = useState<'2scale' | '3scale' | 'dataBar' | 'iconSet'>(initial?.type === 'dataBar' || hint === 'dataBar' ? 'dataBar' : initial?.type === 'iconSet' || hint === 'iconSet' ? 'iconSet' : initial?.type === 'colorScale' && initial.points.length === 3 ? '3scale' : '2scale');
  const scale0 = initial?.type === 'colorScale' ? initial.points : null;
  const [minV, setMinV] = useState<Cfvo>(scale0?.[0]?.cfvo ?? (initial?.type === 'dataBar' ? initial.min : { type: 'min' }));
  const [midV, setMidV] = useState<Cfvo>(scale0 && scale0.length === 3 ? scale0[1].cfvo : { type: 'percentile', value: '50' });
  const [maxV, setMaxV] = useState<Cfvo>(scale0?.[scale0.length - 1]?.cfvo ?? (initial?.type === 'dataBar' ? initial.max : { type: 'max' }));
  const [minC, setMinC] = useState(scale0?.[0]?.color ?? '#F8696B');
  const [midC, setMidC] = useState(scale0 && scale0.length === 3 ? scale0[1].color : '#FFEB84');
  const [maxC, setMaxC] = useState(scale0?.[scale0.length - 1]?.color ?? '#63BE7B');
  const [barColor, setBarColor] = useState(initial?.type === 'dataBar' ? initial.color : DATA_BAR_COLORS[0].color);
  const [gradient, setGradient] = useState(initial?.type === 'dataBar' ? initial.gradient : true);
  const [barOnly, setBarOnly] = useState(initial?.type === 'dataBar' ? !initial.showValue : false);
  const [iconSet, setIconSet] = useState<IconSetName>(initial?.type === 'iconSet' ? initial.iconSet : '3TrafficLights1');
  const [reverse, setReverse] = useState(initial?.type === 'iconSet' ? initial.reverse : false);
  const [iconOnly, setIconOnly] = useState(initial?.type === 'iconSet' ? !initial.showValue : false);
  const [thresholds, setThresholds] = useState(initial?.type === 'iconSet' ? initial.thresholds : defaultThresholds('3TrafficLights1'));
  // Contain
  const [containType, setContainType] = useState<ContainType>(initial && ['cellIs', 'text', 'timePeriod', 'blanks', 'noBlanks', 'errors', 'noErrors'].includes(initial.type) ? (initial.type as ContainType) : 'cellIs');
  const [op, setOp] = useState<CellIsOperator>(initial?.type === 'cellIs' ? initial.operator : 'between');
  const [f1, setF1] = useState(initial?.type === 'cellIs' ? `=${initial.formula1}` : '');
  const [f2, setF2] = useState(initial?.type === 'cellIs' && initial.formula2 ? `=${initial.formula2}` : '');
  const [textOp, setTextOp] = useState<TextOperator>(initial?.type === 'text' ? initial.operator : 'containsText');
  const [text, setText] = useState(initial?.type === 'text' ? initial.text : '');
  const [period, setPeriod] = useState<TimePeriod>(initial?.type === 'timePeriod' ? initial.period : 'yesterday');
  // Rank / average / unique / formula
  const [bottom, setBottom] = useState(initial?.type === 'top10' ? initial.bottom : false);
  const [percent, setPercent] = useState(initial?.type === 'top10' ? initial.percent : false);
  const [rank, setRank] = useState<number | ''>(initial?.type === 'top10' ? initial.rank : 10);
  const [avgKind, setAvgKind] = useState(initial?.type === 'aboveAverage' ? `${initial.above ? 'above' : 'below'}${initial.equal ? '-eq' : ''}${initial.stdDev ? `-${initial.stdDev}` : ''}` : 'above');
  const [unique, setUnique] = useState<'duplicateValues' | 'uniqueValues'>(initial?.type === 'uniqueValues' ? 'uniqueValues' : 'duplicateValues');
  const [formula, setFormula] = useState(initial?.type === 'expression' ? `=${initial.formula}` : '');
  const ui = useUI();
  const theme = ui.doc.wb.theme;
  const strip = (s: string) => operand(s.trim());

  const build = (): CfRule | string => {
    switch (kind) {
      case 'values':
        if (style === 'dataBar') return { type: 'dataBar', min: minV, max: maxV, color: barColor, gradient, showValue: !barOnly };
        if (style === 'iconSet') return { type: 'iconSet', iconSet, reverse, showValue: !iconOnly, thresholds };
        return {
          type: 'colorScale',
          points: style === '3scale' ? [{ cfvo: minV, color: minC }, { cfvo: midV, color: midC }, { cfvo: maxV, color: maxC }] : [{ cfvo: minV, color: minC }, { cfvo: maxV, color: maxC }],
        };
      case 'contain':
        if (containType === 'cellIs') {
          if (!f1.trim()) return 'Enter a value.';
          const two = op === 'between' || op === 'notBetween';
          if (two && !f2.trim()) return 'Enter the second value.';
          return { type: 'cellIs', operator: op, formula1: strip(f1), formula2: two ? strip(f2) : undefined, format };
        }
        if (containType === 'text') return text ? { type: 'text', operator: textOp, text, format } : 'Enter the text.';
        if (containType === 'timePeriod') return { type: 'timePeriod', period, format };
        return { type: containType, format };
      case 'rank': {
        const n = typeof rank === 'number' ? rank : 10;
        return { type: 'top10', bottom, percent, rank: Math.max(1, Math.min(percent ? 100 : 1000, n)), format };
      }
      case 'average': {
        const [dir, a, b] = avgKind.split('-');
        const equal = a === 'eq';
        const sd = Number(equal ? b : a) || undefined;
        return { type: 'aboveAverage', above: dir === 'above', equal, stdDev: sd, format };
      }
      case 'unique':
        return { type: unique, format };
      case 'formula':
        return formula.trim() ? { type: 'expression', formula: formula.trim().replace(/^=/, ''), format } : 'Enter a formula.';
    }
  };

  const css = (h: string) => h;
  const needsFormat = kind !== 'values';

  return (
    <Dialog
      title={title}
      onClose={onClose}
      width={560}
      testId="cf-rule"
      onOk={() => {
        const r = build();
        if (typeof r === 'string') {
          ui.notify(r);
          return false;
        }
        onSave(r);
        return true;
      }}
    >
      <div className="sh-label">Select a Rule Type:</div>
      <div className="sh-listbox" role="listbox" aria-label="Rule type" style={{ height: 'auto' }} data-testid="cf-rule-types">
        {RULE_KINDS.map((k) => (
          <div key={k.kind} role="option" aria-selected={kind === k.kind} className={`sh-listbox-item${kind === k.kind ? ' selected' : ''}`} onClick={() => setKind(k.kind)}>
            ► {k.label}
          </div>
        ))}
      </div>
      <div className="sh-label">Edit the Rule Description:</div>
      <fieldset className="sh-group">
        {kind === 'values' && (
          <div className="sh-stack">
            <Field label="Format Style:">
              <select value={style} data-testid="cf-style" onChange={(e) => setStyle(e.target.value as typeof style)}>
                <option value="2scale">2-Color Scale</option>
                <option value="3scale">3-Color Scale</option>
                <option value="dataBar">Data Bar</option>
                <option value="iconSet">Icon Sets</option>
              </select>
            </Field>
            {(style === '2scale' || style === '3scale' || style === 'dataBar') && (
              <div className="sh-cfvo-grid">
                <div>
                  <b>Minimum</b>
                  <CfvoEditor value={minV} which="min" onChange={setMinV} />
                  {style !== 'dataBar' && <ColorButton value={spec(minC)} noneLabel="—" onChange={(c) => c && setMinC(css(hex(c, theme)!))} />}
                </div>
                {style === '3scale' && (
                  <div>
                    <b>Midpoint</b>
                    <CfvoEditor value={midV} which="mid" onChange={setMidV} />
                    <ColorButton value={spec(midC)} noneLabel="—" onChange={(c) => c && setMidC(hex(c, theme)!)} />
                  </div>
                )}
                <div>
                  <b>Maximum</b>
                  <CfvoEditor value={maxV} which="max" onChange={setMaxV} />
                  {style !== 'dataBar' && <ColorButton value={spec(maxC)} noneLabel="—" onChange={(c) => c && setMaxC(hex(c, theme)!)} />}
                </div>
              </div>
            )}
            {style === 'dataBar' && (
              <div className="sh-row wrap">
                <Field label="Bar color:">
                  <ColorButton value={spec(barColor)} noneLabel="—" onChange={(c) => c && setBarColor(hex(c, theme)!)} />
                </Field>
                <Field label="Fill:">
                  <select value={gradient ? 'gradient' : 'solid'} onChange={(e) => setGradient(e.target.value === 'gradient')}>
                    <option value="gradient">Gradient Fill</option>
                    <option value="solid">Solid Fill</option>
                  </select>
                </Field>
                <Check label="Show Bar Only" checked={barOnly} onChange={setBarOnly} />
              </div>
            )}
            {style === 'iconSet' && (
              <div className="sh-stack">
                <div className="sh-row wrap">
                  <Field label="Icon Style:">
                    <select value={iconSet} data-testid="cf-iconset" onChange={(e) => { const s = e.target.value as IconSetName; setIconSet(s); setThresholds(defaultThresholds(s)); }}>
                      {ICON_SET_GROUPS.flatMap((g) => g.sets).map((s) => <option key={s.name} value={s.name}>{s.label}</option>)}
                    </select>
                  </Field>
                  <Check label="Reverse Icon Order" checked={reverse} onChange={setReverse} />
                  <Check label="Show Icon Only" checked={iconOnly} onChange={setIconOnly} />
                </div>
                <RulePreview rule={{ type: 'iconSet', iconSet, reverse, showValue: !iconOnly, thresholds }} />
                {thresholds.slice(1).reverse().map((t, k) => {
                  const i = thresholds.length - 1 - k;
                  return (
                    <div key={i} className="sh-row">
                      <span>Icon {k + 1}: when value is</span>
                      <select value={t.gte ? '>=' : '>'} onChange={(e) => setThresholds(thresholds.map((x, j) => (j === i ? { ...x, gte: e.target.value === '>=' } : x)))}>
                        <option value=">=">&gt;=</option>
                        <option value=">">&gt;</option>
                      </select>
                      <input value={t.cfvo.value ?? ''} style={{ width: 70 }} onChange={(e) => setThresholds(thresholds.map((x, j) => (j === i ? { ...x, cfvo: { ...x.cfvo, value: e.target.value } } : x)))} />
                      <select value={t.cfvo.type} onChange={(e) => setThresholds(thresholds.map((x, j) => (j === i ? { ...x, cfvo: { ...x.cfvo, type: e.target.value as CfvoType } } : x)))}>
                        {CFVO_TYPES.filter((c) => c.type !== 'min' && c.type !== 'max').map((c) => <option key={c.type} value={c.type}>{c.label}</option>)}
                      </select>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
        {kind === 'contain' && (
          <div className="sh-stack">
            <span>Format only cells with:</span>
            <div className="sh-row wrap">
              <select value={containType} data-testid="cf-contain-type" onChange={(e) => setContainType(e.target.value as ContainType)}>
                <option value="cellIs">Cell Value</option>
                <option value="text">Specific Text</option>
                <option value="timePeriod">Dates Occurring</option>
                <option value="blanks">Blanks</option>
                <option value="noBlanks">No Blanks</option>
                <option value="errors">Errors</option>
                <option value="noErrors">No Errors</option>
              </select>
              {containType === 'cellIs' && (
                <>
                  <select value={op} data-testid="cf-operator" onChange={(e) => setOp(e.target.value as CellIsOperator)}>
                    {CELL_OPS.map((o) => <option key={o.op} value={o.op}>{o.label}</option>)}
                  </select>
                  <RangeField value={f1} onChange={setF1} label="Value" testId="cf-f1" />
                  {(op === 'between' || op === 'notBetween') && (
                    <>
                      <span>and</span>
                      <RangeField value={f2} onChange={setF2} label="Second value" testId="cf-f2" />
                    </>
                  )}
                </>
              )}
              {containType === 'text' && (
                <>
                  <select value={textOp} onChange={(e) => setTextOp(e.target.value as TextOperator)}>
                    {TEXT_OPS.map((o) => <option key={o.op} value={o.op}>{o.label}</option>)}
                  </select>
                  <input value={text} data-testid="cf-rule-text" onChange={(e) => setText(e.target.value)} />
                </>
              )}
              {containType === 'timePeriod' && (
                <select value={period} onChange={(e) => setPeriod(e.target.value as TimePeriod)}>
                  {TIME_PERIODS.map((p) => <option key={p.period} value={p.period}>{p.label}</option>)}
                </select>
              )}
            </div>
          </div>
        )}
        {kind === 'rank' && (
          <div className="sh-row wrap">
            <span>Format values that rank in the:</span>
            <select value={bottom ? 'bottom' : 'top'} onChange={(e) => setBottom(e.target.value === 'bottom')}>
              <option value="top">Top</option>
              <option value="bottom">Bottom</option>
            </select>
            <NumberField value={rank} min={1} max={percent ? 100 : 1000} label="Rank" onChange={setRank} />
            <Check label="% of the selected range" checked={percent} onChange={setPercent} />
          </div>
        )}
        {kind === 'average' && (
          <div className="sh-row wrap">
            <span>Format values that are:</span>
            <select value={avgKind} onChange={(e) => setAvgKind(e.target.value)}>
              <option value="above">above</option>
              <option value="below">below</option>
              <option value="above-eq">equal or above</option>
              <option value="below-eq">equal or below</option>
              <option value="above-1">1 std dev above</option>
              <option value="below-1">1 std dev below</option>
              <option value="above-2">2 std dev above</option>
              <option value="below-2">2 std dev below</option>
              <option value="above-3">3 std dev above</option>
              <option value="below-3">3 std dev below</option>
            </select>
            <span>the average for the selected range</span>
          </div>
        )}
        {kind === 'unique' && (
          <div className="sh-row wrap">
            <span>Format all:</span>
            <select value={unique} onChange={(e) => setUnique(e.target.value as typeof unique)}>
              <option value="duplicateValues">duplicate</option>
              <option value="uniqueValues">unique</option>
            </select>
            <span>values in the selected range</span>
          </div>
        )}
        {kind === 'formula' && (
          <Field label="Format values where this formula is true:" wide>
            <RangeField value={formula} onChange={setFormula} label="Formula" testId="cf-formula" />
          </Field>
        )}
        {needsFormat && <FormatEditor value={format} onChange={setFormat} />}
      </fieldset>
    </Dialog>
  );
}

/** Home > Conditional Formatting > New Rule / More Rules: adds the rule to the selection. */
export function CfRuleDialog({ hint, onClose }: { hint?: string; onClose: () => void }) {
  const ui = useUI();
  const initial = cfInitialRule(hint);
  return (
    <RuleEditorDialog
      title="New Formatting Rule"
      initial={initial}
      hint={hint}
      onClose={onClose}
      onSave={(rule) => {
        commit(ui);
        ui.doc.addCondFormat(rule);
      }}
    />
  );
}

// -----------------------------------------------------------------------------------------------
// Rules Manager

interface Row {
  key: string;
  id?: string;
  rule: CfRule | null;
  raw?: XmlElement;
  ranges: Range[];
  text: string;
  stopIfTrue: boolean;
}

let rowKey = 0;

export function CfManagerDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const [scope, setScope] = useState<'selection' | 'sheet'>('selection');
  const load = (s: typeof scope): Row[] =>
    doc.cfRules(doc.sheet, s === 'selection' ? doc.selection.ranges : undefined).map((e) => ({
      key: `r${rowKey++}`,
      id: e.id,
      rule: e.rule,
      raw: e.raw,
      ranges: e.ranges,
      text: e.ranges.map((r) => rangeText(r, doc.sheet)).join(','),
      stopIfTrue: e.stopIfTrue,
    }));
  const [rows, setRows] = useState<Row[]>(() => load('selection'));
  const [sel, setSel] = useState<string | null>(rows[0]?.key ?? null);
  const [editor, setEditor] = useState<{ row: Row | null } | null>(null);
  const [dirty, setDirty] = useState(false);
  const index = rows.findIndex((r) => r.key === sel);

  const apply = (): boolean => {
    const list = [];
    for (const r of rows) {
      const ranges = r.text.split(',').map((t) => parseRangeText(doc, t)?.range).filter((x): x is Range => !!x);
      if (!ranges.length) {
        ui.notify(`"${r.text}" is not a valid reference.`);
        return false;
      }
      list.push({ id: r.id, rule: r.rule, raw: r.raw, ranges, stopIfTrue: r.stopIfTrue });
    }
    commit(ui);
    doc.applyRuleList(list, scope === 'selection' ? doc.selection.ranges : null);
    setRows(load(scope));
    setDirty(false);
    return true;
  };

  const move = (dir: -1 | 1) => {
    if (index < 0) return;
    const j = index + dir;
    if (j < 0 || j >= rows.length) return;
    const next = rows.slice();
    [next[index], next[j]] = [next[j], next[index]];
    setRows(next);
    setDirty(true);
  };

  return (
    <>
      <Dialog
        title="Conditional Formatting Rules Manager"
        onClose={onClose}
        onOk={() => (dirty ? apply() : true)}
        width={720}
        testId="cf-manager"
        footerLeft={
          <button type="button" className="sh-btn" disabled={!dirty} onClick={() => apply()}>
            Apply
          </button>
        }
      >
        <div className="sh-row">
          <span>Show formatting rules for:</span>
          <select
            value={scope}
            data-testid="cf-manager-scope"
            onChange={(e) => {
              const s = e.target.value as typeof scope;
              setScope(s);
              const next = load(s);
              setRows(next);
              setSel(next[0]?.key ?? null);
            }}
          >
            <option value="selection">Current Selection</option>
            <option value="sheet">This Worksheet</option>
          </select>
        </div>
        <div className="sh-row wrap sh-toolbar-row">
          <button type="button" className="sh-btn" data-testid="cf-manager-new" onClick={() => setEditor({ row: null })}>New Rule...</button>
          <button type="button" className="sh-btn" disabled={index < 0 || !rows[index]?.rule} onClick={() => setEditor({ row: rows[index] })}>Edit Rule...</button>
          <button type="button" className="sh-btn" disabled={index < 0} data-testid="cf-manager-delete" onClick={() => { setRows(rows.filter((r) => r.key !== sel)); setSel(rows[index + 1]?.key ?? rows[index - 1]?.key ?? null); setDirty(true); }}>Delete Rule</button>
          <button type="button" className="sh-btn" disabled={index < 0} onClick={() => { const r = rows[index]; const copy = { ...r, key: `r${rowKey++}`, id: undefined }; setRows([...rows.slice(0, index + 1), copy, ...rows.slice(index + 1)]); setDirty(true); }}>Duplicate Rule</button>
          <button type="button" className="sh-btn icon" aria-label="Move up" disabled={index <= 0} onClick={() => move(-1)}>▲</button>
          <button type="button" className="sh-btn icon" aria-label="Move down" disabled={index < 0 || index >= rows.length - 1} onClick={() => move(1)}>▼</button>
        </div>
        <div className="sh-table" role="table" data-testid="cf-manager-rules">
          <div className="sh-table-head" role="row">
            <span>Rule (applied in order shown)</span>
            <span>Format</span>
            <span>Applies to</span>
            <span>Stop If True</span>
          </div>
          <div className="sh-table-body">
            {rows.map((r) => (
              <div key={r.key} role="row" className={`sh-table-row${r.key === sel ? ' selected' : ''}`} onClick={() => setSel(r.key)} onDoubleClick={() => r.rule && setEditor({ row: r })}>
                <span>{describeRule(r.rule, r.raw ?? ({ type: 'element', name: 'cfRule', attrs: {}, children: [] } as XmlElement))}</span>
                <span><RulePreview rule={r.rule} /></span>
                <span>
                  <input
                    value={r.text}
                    aria-label="Applies to"
                    onChange={(e) => {
                      setRows(rows.map((x) => (x.key === r.key ? { ...x, text: e.target.value } : x)));
                      setDirty(true);
                    }}
                  />
                </span>
                <span>
                  <input
                    type="checkbox"
                    aria-label="Stop If True"
                    checked={r.stopIfTrue}
                    disabled={r.rule?.type === 'colorScale' || r.rule?.type === 'dataBar' || r.rule?.type === 'iconSet'}
                    onChange={(e) => {
                      setRows(rows.map((x) => (x.key === r.key ? { ...x, stopIfTrue: e.target.checked } : x)));
                      setDirty(true);
                    }}
                  />
                </span>
              </div>
            ))}
            {!rows.length && <div className="sh-table-empty">No rules.</div>}
          </div>
        </div>
      </Dialog>
      {editor && (
        <RuleEditorDialog
          title={editor.row ? 'Edit Formatting Rule' : 'New Formatting Rule'}
          initial={editor.row?.rule ?? null}
          onClose={() => setEditor(null)}
          onSave={(rule) => {
            if (editor.row) setRows(rows.map((x) => (x.key === editor.row!.key ? { ...x, rule, raw: undefined } : x)));
            else {
              const row: Row = { key: `r${rowKey++}`, rule, ranges: doc.selection.ranges, text: doc.selection.ranges.map((r) => rangeText(r, doc.sheet)).join(','), stopIfTrue: false };
              setRows([row, ...rows]);
              setSel(row.key);
            }
            setDirty(true);
          }}
        />
      )}
    </>
  );
}
