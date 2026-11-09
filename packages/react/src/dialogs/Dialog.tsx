/**
 * Dialog frame (title bar with close button, OK / Cancel at the bottom right)
 * and the form controls dialogs share: tabs, range fields that can pick cells on the grid, number
 * fields and check boxes.
 */
import { useEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { Modal } from '../Modal';
import { Icon } from '../Icon';
import { type ColorSpec, type IconName, rangeText, resolveColor, type Selection, selectionText, type Sheet } from '@sheetline/core/views';
import { useUI } from '../context';
import { Dropdown } from '../ribbon/Ribbon';
import { ThemePalette } from '../ThemePalette';

export interface DialogProps {
  title: string;
  icon?: IconName;
  onClose: () => void;
  /** OK: return false to keep the dialog open. */
  onOk?: () => boolean | void | Promise<boolean | void>;
  ok?: string;
  cancel?: string;
  width?: number;
  /** Modal dialogs dim the page; small dialogs let the grid be used meanwhile. */
  blocking?: boolean;
  testId?: string;
  children: ReactNode;
  /** Buttons at the left of the footer (Delete, Clear All...). */
  footerLeft?: ReactNode;
  okDisabled?: boolean;
  bodyClassName?: string;
  /** No OK button (Close only). */
  closeOnly?: boolean;
}

export function Dialog({ title, icon, onClose, onOk, ok = 'OK', cancel = 'Cancel', width = 420, blocking = true, testId, children, footerLeft, okDisabled, bodyClassName, closeOnly }: DialogProps) {
  const busy = useRef(false);
  const submit = async () => {
    if (busy.current || okDisabled) return;
    busy.current = true;
    try {
      const keep = (await onOk?.()) === false;
      if (!keep) onClose();
    } finally {
      busy.current = false;
    }
  };
  const width2 = Math.min(width, window.innerWidth - 16);
  return (
    <Modal
      title={title}
      icon={icon}
      onClose={onClose}
      width={width2}
      blocking={blocking}
      testId={testId}
      bodyClassName={`sh-dialog-body${bodyClassName ? ` ${bodyClassName}` : ''}`}
      footer={
        <>
          {footerLeft && <span className="sh-dialog-foot-left">{footerLeft}</span>}
          {!closeOnly && (
            <button type="button" className="sh-btn primary" data-testid={testId ? `${testId}-ok` : undefined} disabled={okDisabled} onClick={() => void submit()}>
              {ok}
            </button>
          )}
          <button type="button" className="sh-btn" data-testid={testId ? `${testId}-cancel` : undefined} onClick={onClose}>
            {closeOnly ? 'Close' : cancel}
          </button>
        </>
      }
    >
      <form
        className="sh-dialog-form"
        onSubmit={(e) => {
          e.preventDefault();
          void submit();
        }}
      >
        {children}
        <button type="submit" hidden aria-hidden tabIndex={-1} />
      </form>
    </Modal>
  );
}

export function Tabs<T extends string>({ tabs, value, onChange, testId }: { tabs: { id: T; label: string }[]; value: T; onChange: (id: T) => void; testId?: string }) {
  return (
    <div className="sh-dtabs" role="tablist" data-testid={testId}>
      {tabs.map((t) => (
        <button key={t.id} type="button" role="tab" aria-selected={value === t.id} className={value === t.id ? 'active' : ''} data-testid={testId ? `${testId}-${t.id}` : undefined} onClick={() => onChange(t.id)}>
          {t.label}
        </button>
      ))}
    </div>
  );
}

export function Check({ label, checked, onChange, testId, disabled }: { label: ReactNode; checked: boolean; onChange: (v: boolean) => void; testId?: string; disabled?: boolean }) {
  return (
    <label className={`sh-check${disabled ? ' disabled' : ''}`}>
      <input type="checkbox" checked={checked} disabled={disabled} data-testid={testId} onChange={(e) => onChange(e.target.checked)} />
      <span>{label}</span>
    </label>
  );
}

export function Radio({ label, checked, onChange, name, testId, disabled }: { label: ReactNode; checked: boolean; onChange: () => void; name: string; testId?: string; disabled?: boolean }) {
  return (
    <label className={`sh-check${disabled ? ' disabled' : ''}`}>
      <input type="radio" name={name} checked={checked} disabled={disabled} data-testid={testId} onChange={onChange} />
      <span>{label}</span>
    </label>
  );
}

export function Field({ label, children, wide }: { label: ReactNode; children: ReactNode; wide?: boolean }) {
  return (
    <label className={`sh-field${wide ? ' wide' : ''}`}>
      <span className="sh-field-label">{label}</span>
      {children}
    </label>
  );
}

export { parseRangeText, rangeText, selectionText } from '@sheetline/core/views';

/**
 * A reference box with a picker button: the dialog hides while cells are selected on the grid
 * (a bar at the top shows the reference), then comes back with the selection filled in.
 */
export function RangeField({ value, onChange, label, testId, home, multi, autoFocus }: { value: string; onChange: (v: string) => void; label?: string; testId?: string; home?: Sheet; multi?: boolean; autoFocus?: boolean }) {
  const ui = useUI();
  const [picking, setPicking] = useState(false);
  const homeSheet = useRef(home ?? ui.doc.sheet);
  // The selection before picking: dialogs act on it (Text to Columns, Data Validation...), so it comes back after.
  const before = useRef<Selection | null>(null);
  useEffect(() => {
    if (!picking) return;
    ui.layer.classList.add('sh-picking');
    return () => ui.layer.classList.remove('sh-picking');
  }, [picking, ui.layer]);
  const live = picking ? (multi ? selectionText(ui.doc, homeSheet.current) : rangeText(ui.doc.bounded(ui.doc.range, ui.doc.sheet) ?? ui.doc.range, ui.doc.sheet, homeSheet.current)) : '';
  const pick = () => {
    before.current = ui.doc.selectionOf(homeSheet.current);
    setPicking(true);
  };
  const done = (accept: boolean) => {
    if (accept) onChange(live);
    setPicking(false);
    if (ui.doc.sheet !== homeSheet.current) ui.doc.setActiveSheet(ui.doc.wb.sheets.indexOf(homeSheet.current));
    if (before.current) ui.doc.setSelection(before.current, homeSheet.current);
  };
  return (
    <span className="sh-range-field">
      <input
        aria-label={label ?? 'Reference'}
        data-testid={testId}
        value={value}
        autoFocus={autoFocus}
        spellCheck={false}
        autoCapitalize="off"
        onChange={(e) => onChange(e.target.value)}
      />
      <button type="button" className="sh-range-pick" title="Select cells" aria-label="Select cells on the sheet" data-testid={testId ? `${testId}-pick` : undefined} onClick={pick}>
        <Icon name="highlight_alt" />
      </button>
      {picking &&
        createPortal(
          <div className="sh-range-picker" role="dialog" aria-label="Select a range" data-testid="range-picker">
            <span className="sh-range-picker-label">{label ?? 'Select a range'}</span>
            <input readOnly value={live} aria-label="Selected range" />
            <button type="button" className="sh-btn primary" data-testid="range-picker-ok" onClick={() => done(true)}>
              <Icon name="check" />
            </button>
            <button type="button" className="sh-btn" aria-label="Cancel" onClick={() => done(false)}>
              <Icon name="close" />
            </button>
          </div>,
          ui.layer,
        )}
    </span>
  );
}

export function NumberField({ value, onChange, min, max, step = 1, testId, label, width = 80, disabled }: { value: number | ''; onChange: (v: number | '') => void; min?: number; max?: number; step?: number; testId?: string; label?: string; width?: number; disabled?: boolean }) {
  return (
    <input
      type="number"
      className="sh-number"
      inputMode="decimal"
      aria-label={label}
      data-testid={testId}
      value={value}
      min={min}
      max={max}
      step={step}
      disabled={disabled}
      style={{ width }}
      onChange={(e) => onChange(e.target.value === '' ? '' : Number(e.target.value))}
    />
  );
}

/** A selectable list (list boxes of Excel dialogs). */
export function ListBox<T>({ items, value, onChange, render, keyOf, testId, height = 160, onActivate, label }: { items: T[]; value: T | null; onChange: (v: T) => void; render: (v: T) => ReactNode; keyOf: (v: T) => string; testId?: string; height?: number; onActivate?: (v: T) => void; label?: string }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    ref.current?.querySelector<HTMLElement>('.selected')?.scrollIntoView({ block: 'nearest' });
  }, [value]);
  const index = value === null ? -1 : items.findIndex((x) => keyOf(x) === keyOf(value));
  return (
    <div
      ref={ref}
      className="sh-listbox"
      role="listbox"
      aria-label={label}
      tabIndex={0}
      data-testid={testId}
      style={{ height }}
      onKeyDown={(e) => {
        if (e.key === 'ArrowDown' && items.length) {
          e.preventDefault();
          onChange(items[Math.min(items.length - 1, index + 1)]);
        } else if (e.key === 'ArrowUp' && items.length) {
          e.preventDefault();
          onChange(items[Math.max(0, index - 1)]);
        } else if (e.key === 'Enter' && value !== null && onActivate) {
          e.preventDefault();
          onActivate(value);
        }
      }}
    >
      {items.map((it) => {
        const k = keyOf(it);
        const sel = value !== null && keyOf(value) === k;
        return (
          <div key={k} role="option" aria-selected={sel} className={`sh-listbox-item${sel ? ' selected' : ''}`} onClick={() => onChange(it)} onDoubleClick={() => onActivate?.(it)}>
            {render(it)}
          </div>
        );
      })}
    </div>
  );
}

/** A colour drop-down button (theme palette) for dialogs. */
export function ColorButton({ value, onChange, noneLabel = 'Automatic', noneSwatch, testId, label, disabled }: { value: ColorSpec | null | undefined; onChange: (c: ColorSpec | null) => void; noneLabel?: string; noneSwatch?: string; testId?: string; label?: string; disabled?: boolean }) {
  const ui = useUI();
  const [anchor, setAnchor] = useState<DOMRect | null>(null);
  const css = value ? `#${resolveColor(value, ui.doc.wb.theme, '000000')}` : noneSwatch ?? 'transparent';
  return (
    <>
      <button type="button" className="sh-color-button" aria-label={label ?? 'Color'} data-testid={testId} disabled={disabled} onClick={(e) => setAnchor((e.currentTarget as HTMLElement).getBoundingClientRect())}>
        <span className="sh-color-chip" style={{ background: css }} />
        <span className="sh-color-name">{value ? css.toUpperCase() : noneLabel}</span>
        <Icon name="arrow_drop_down" />
      </button>
      {anchor && (
        <Dropdown anchor={anchor} onClose={() => setAnchor(null)} className="nested" testId={testId ? `${testId}-palette` : undefined}>
          <ThemePalette
            theme={ui.doc.wb.theme}
            noneLabel={noneLabel}
            noneSwatch={noneSwatch}
            onPick={(c) => {
              onChange(c);
              setAnchor(null);
            }}
          />
        </Dropdown>
      )}
    </>
  );
}
