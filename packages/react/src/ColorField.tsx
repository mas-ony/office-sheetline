/** A colour text field with a palette button opening the colour picker. */
import { useState } from 'react';
import { ColorPicker } from './ColorPicker';
import { normalizeHex } from '@sheetline/core/views';
import { Icon } from './Icon';

export interface ColorFieldProps {
  value: string;
  onChange: (hex: string) => void;
  placeholder?: string;
  disabled?: boolean;
  testId?: string;
  id?: string;
}

export function ColorField({ value, onChange, placeholder, disabled, testId, id }: ColorFieldProps) {
  const [anchor, setAnchor] = useState<DOMRect | null>(null);
  const [text, setText] = useState<string | null>(null);
  return (
    <div className="sh-range-input sh-color-field">
      <input
        id={id}
        className="sh-input"
        value={text ?? value}
        placeholder={placeholder}
        disabled={disabled}
        spellCheck={false}
        data-testid={testId}
        style={value ? { color: value } : undefined}
        onChange={(e) => {
          setText(e.target.value);
          const hex = normalizeHex(e.target.value);
          if (hex || e.target.value === '') onChange(hex ?? '');
        }}
        onBlur={() => setText(null)}
      />
      <button type="button" aria-label="Pick a colour" disabled={disabled} data-testid={testId ? `${testId}-pick` : undefined} onClick={(e) => setAnchor((e.currentTarget as HTMLElement).getBoundingClientRect())}>
        <Icon name="palette" size={18} />
      </button>
      {anchor && (
        <ColorPicker
          anchor={anchor}
          value={value || null}
          live={false}
          onChange={(hex) => {
            onChange(hex ?? '');
            setText(null);
          }}
          onClose={() => setAnchor(null)}
        />
      )}
    </div>
  );
}
