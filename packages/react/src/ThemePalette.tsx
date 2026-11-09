/**
 * Office-style colour palette: an "Automatic" / "No Fill" entry, the theme colours with their five
 * tint rows, the standard colours, recent colours and "More Colors..." (the spectrum picker).
 * Theme picks are returned as theme colours so they follow later theme changes.
 */
import { useRef, useState } from 'react';
import { type ColorSpec, rgbSpec, specToCss, STANDARD_COLORS, type Theme, THEME_COLUMNS, tintLabel, withRecent } from '@sheetline/core/views';
import { ColorPicker } from './ColorPicker';
import { usePersistent } from './context';

export interface ThemePaletteProps {
  theme: Theme;
  onPick: (color: ColorSpec | null) => void;
  /** Label of the empty choice ("No Fill", "Automatic", "No Color"); omitted hides it. */
  noneLabel?: string;
  /** Colour shown on the "Automatic" swatch. */
  noneSwatch?: string;
  testId?: string;
}

export function ThemePalette({ theme, onPick, noneLabel, noneSwatch, testId }: ThemePaletteProps) {
  const [recent, setRecent] = usePersistent<string[]>('recent-colors', []);
  const [more, setMore] = useState<DOMRect | null>(null);
  const moreBtn = useRef<HTMLButtonElement>(null);
  const pickHex = (hex: string) => {
    setRecent(withRecent(recent, hex, 10));
    onPick(rgbSpec(hex) ?? null);
  };
  const swatch = (css: string, title: string, onClick: () => void, key: string) => (
    <button key={key} type="button" className="sh-swatch" style={{ background: css }} title={title} aria-label={title} onClick={onClick} />
  );
  return (
    <div className="sh-palette" data-testid={testId} onMouseDown={(e) => e.preventDefault()}>
      {noneLabel && (
        <button type="button" className="sh-palette-none" data-testid={testId ? `${testId}-none` : undefined} onClick={() => onPick(null)}>
          <span className="sh-swatch" style={{ background: noneSwatch ?? 'transparent' }} />
          {noneLabel}
        </button>
      )}
      <div className="sh-palette-title">Theme Colors</div>
      <div className="sh-palette-grid theme">
        {THEME_COLUMNS.map((col) => swatch(specToCss({ theme: col.theme }, theme), col.name, () => onPick({ theme: col.theme }), `b${col.theme}`))}
      </div>
      <div className="sh-palette-grid tints">
        {[0, 1, 2, 3, 4].map((row) =>
          THEME_COLUMNS.map((col) => {
            const tint = col.tints[row];
            return swatch(specToCss({ theme: col.theme, tint }, theme), `${col.name}, ${tintLabel(tint)}`, () => onPick({ theme: col.theme, tint }), `t${row}-${col.theme}`);
          }),
        )}
      </div>
      <div className="sh-palette-title">Standard Colors</div>
      <div className="sh-palette-grid">
        {STANDARD_COLORS.map((c) => swatch(c.hex, c.name, () => pickHex(c.hex), c.hex))}
      </div>
      {recent.length > 0 && (
        <>
          <div className="sh-palette-title">Recent Colors</div>
          <div className="sh-palette-grid">{recent.map((hex) => swatch(hex, hex, () => pickHex(hex), `r${hex}`))}</div>
        </>
      )}
      <button ref={moreBtn} type="button" className="sh-palette-more" onClick={() => setMore(moreBtn.current!.getBoundingClientRect())}>
        More Colors...
      </button>
      {more && (
        <ColorPicker
          anchor={more}
          value={null}
          live={false}
          onChange={(hex) => {
            if (hex) pickHex(hex);
            else onPick(null);
          }}
          onClose={() => setMore(null)}
        />
      )}
    </div>
  );
}
