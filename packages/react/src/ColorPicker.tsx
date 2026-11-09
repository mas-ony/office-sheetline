/**
 * Colour picker pop-up: a palette grid, a spectrum and RGB sliders, with Reset and Done.
 */
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { usePersistent, useUI } from './context';
import { hexToRgb, normalizeHex, paintSpectrum, PALETTE_ROWS, rgbToHex, spectrumPick, withRecent } from '@sheetline/core/views';

function Spectrum({ value, onPick }: { value: string; onPick: (hex: string) => void }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const [point, setPoint] = useState<{ x: number; y: number } | null>(null);
  useEffect(() => {
    if (canvas.current) paintSpectrum(canvas.current);
  }, []);
  void value;
  const pick = (e: React.PointerEvent) => {
    const c = canvas.current;
    if (!c) return;
    const { x, y, hex } = spectrumPick(c, e.clientX, e.clientY);
    setPoint({ x, y });
    if (hex) onPick(hex);
  };
  return (
    <div
      className="sh-color-spectrum"
      onPointerDown={(e) => {
        (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
        pick(e);
      }}
      onPointerMove={(e) => {
        if (e.buttons) pick(e);
      }}
    >
      <canvas ref={canvas} width={232} height={160} />
      {point && <div className="sh-color-point" style={{ left: point.x, top: point.y }} />}
    </div>
  );
}

export interface ColorPickerProps {
  anchor: DOMRect;
  value: string | null;
  onChange: (hex: string | null) => void;
  onClose: () => void;
  /** Applies immediately on every pick (true) or only on Done (false). */
  live?: boolean;
  testId?: string;
}

export function ColorPicker({ anchor, value, onChange, onClose, live = true, testId }: ColorPickerProps) {
  const { layer } = useUI();
  const ref = useRef<HTMLDivElement>(null);
  const [tab, setTab] = useState<'grid' | 'spectrum' | 'sliders'>('grid');
  const [current, setCurrent] = useState<string>(value ?? '#000000');
  const [text, setText] = useState(value ?? '');
  const [recent, setRecent] = usePersistent<string[]>('picker-recent-colors', []);

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    let x = anchor.left;
    let y = anchor.bottom + 4;
    if (x + w > window.innerWidth - 4) x = Math.max(4, window.innerWidth - w - 4);
    if (y + h > window.innerHeight - 4) y = Math.max(4, anchor.top - h - 4);
    el.style.left = `${x}px`;
    el.style.top = `${y}px`;
  }, [anchor]);

  useEffect(() => {
    const down = (e: PointerEvent) => {
      if (!(e.target as HTMLElement).closest?.('.sh-color')) onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    const id = window.setTimeout(() => {
      window.addEventListener('pointerdown', down, true);
      window.addEventListener('keydown', key);
    }, 0);
    return () => {
      window.clearTimeout(id);
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('keydown', key);
    };
  }, [onClose]);

  const pick = (hex: string) => {
    setCurrent(hex);
    setText(hex);
    if (live) onChange(hex);
  };
  const done = () => {
    const hex = normalizeHex(text) ?? current;
    onChange(hex);
    setRecent(withRecent(recent, hex, 12));
    onClose();
  };
  const [r, g, b] = hexToRgb(current);

  return createPortal(
    <div ref={ref} className="sh-color" role="dialog" aria-label="Colour" data-testid={testId}>
      <div className="sh-color-head">
        <button type="button" data-testid="color-reset" onClick={() => { onChange(null); onClose(); }}>Reset</button>
        <button type="button" data-testid="color-done" onClick={done}>Done</button>
      </div>
      <div className="sh-segmented" role="tablist">
        {(['grid', 'spectrum', 'sliders'] as const).map((t) => (
          <button key={t} type="button" role="tab" aria-selected={tab === t} className={tab === t ? 'selected' : ''} onClick={() => setTab(t)}>
            {t[0].toUpperCase() + t.slice(1)}
          </button>
        ))}
      </div>
      {tab === 'grid' && (
        <div className="sh-color-grid" style={{ gridTemplateColumns: `repeat(${PALETTE_ROWS[0].length}, 1fr)` }}>
          {PALETTE_ROWS.flatMap((row, i) =>
            row.map((hex, j) => (
              <button key={`${i}-${j}`} type="button" title={hex} aria-label={hex} data-color={hex} className={current.toUpperCase() === hex ? 'selected' : ''} style={{ background: hex }} onClick={() => pick(hex)} />
            )),
          )}
        </div>
      )}
      {tab === 'spectrum' && <Spectrum value={current} onPick={pick} />}
      {tab === 'sliders' && (
        <div className="sh-color-sliders">
          {(['Red', 'Green', 'Blue'] as const).map((label, k) => {
            const vals = [r, g, b];
            return (
              <label key={label}>
                {label}
                <input
                  type="range"
                  min={0}
                  max={255}
                  value={vals[k]}
                  aria-label={label}
                  onChange={(e) => {
                    const next = [...vals];
                    next[k] = Number(e.target.value);
                    pick(rgbToHex(next[0], next[1], next[2]));
                  }}
                />
                <span>{vals[k]}</span>
              </label>
            );
          })}
        </div>
      )}
      {recent.length > 0 && (
        <div className="sh-color-recent" aria-label="Recent colours">
          {recent.map((hex) => (
            <button key={hex} type="button" title={hex} style={{ background: hex }} onClick={() => pick(hex)} />
          ))}
        </div>
      )}
      <div className="sh-color-foot">
        <span className="swatch" style={{ background: normalizeHex(text) ?? current }} />
        <input
          value={text}
          placeholder="#000000"
          aria-label="Hex colour"
          onChange={(e) => {
            setText(e.target.value);
            const hex = normalizeHex(e.target.value);
            if (hex) setCurrent(hex);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') done();
          }}
        />
      </div>
    </div>,
    layer,
  );
}
