/** Directives giving form fields React's behaviour: autoFocus, and check boxes that show their value. */
import type { Directive } from 'vue';

/** `v-focus`: focuses the element once mounted; `v-focus="false"` does not. */
export const vFocus: Directive<HTMLElement, boolean | undefined> = {
  mounted(el, binding) {
    if (binding.value !== false) el.focus();
  },
};

type CheckBox = HTMLInputElement & { _alChecked?: boolean };

/**
 * `v-checked`: a check box or radio button showing a value, after every render and also after a
 * click its handler ignored (which leaves the value unchanged).
 */
export const vChecked: Directive<CheckBox, boolean> = {
  mounted(el, { value }) {
    el._alChecked = !!value;
    el.checked = !!value;
    // Runs after the change handlers and the render they cause (both queued before this).
    el.addEventListener('change', () => queueMicrotask(() => (el.checked = !!el._alChecked)));
  },
  updated(el, { value }) {
    el._alChecked = !!value;
    el.checked = !!value;
  },
};

type NumberInput = HTMLInputElement & { _alValue?: number | '' };

/** As React does for number inputs: "1.50" stays while the value is 1.5 (only a different number is written). */
function syncNumber(el: NumberInput): void {
  const v = el._alValue ?? '';
  // eslint-disable-next-line eqeqeq
  if ((v === 0 && el.value === '') || el.value != (v as unknown as string)) el.value = String(v);
}

/** `v-number-value`: the value of a number input, kept like React's controlled number inputs. */
export const vNumberValue: Directive<NumberInput, number | ''> = {
  mounted(el, { value }) {
    el._alValue = value;
    syncNumber(el);
    el.addEventListener('input', () => queueMicrotask(() => syncNumber(el)));
  },
  updated(el, { value }) {
    el._alValue = value;
    syncNumber(el);
  },
};

type TextInput = HTMLInputElement & { _alText?: string };

/**
 * `v-value`: the value of a text input that refuses some input, as React's controlled inputs do:
 * after typing, the field shows the value again (unchanged when the handler refused the text).
 */
export const vValue: Directive<TextInput, string | number> = {
  mounted(el, { value }) {
    el._alText = String(value);
    el.value = el._alText;
    el.addEventListener('input', () =>
      queueMicrotask(() => {
        if (el.value !== el._alText) el.value = el._alText ?? '';
      }),
    );
  },
  updated(el, { value }) {
    el._alText = String(value);
    if (el.value !== el._alText) el.value = el._alText;
  },
};

/** `v-default-value`: the first value of an input the user then edits freely (React's defaultValue). */
export const vDefaultValue: Directive<HTMLInputElement, string> = {
  mounted(el, { value }) {
    el.defaultValue = value;
  },
};
