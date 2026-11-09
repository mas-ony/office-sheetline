<script setup lang="ts">
/** Style: a custom cell style made from the active cell (kept in this browser). */
import { resolveColor, type StylePatch } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { usePersistent, useUI } from '../../context';
import { vFocus } from '../../directives';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';

interface CustomCellStyle {
  name: string;
  patch: StylePatch;
}
type Part = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const styles = usePersistent<CustomCellStyle[]>('custom-cell-styles', []);
const name = ref(`Style ${styles.value.length + 1}`);
const parts = shallowRef<Record<Part, boolean>>({ number: true, alignment: true, font: true, border: true, fill: true, protection: true });
const view = computed(() => {
  const { doc } = ui.value;
  const { r, c } = doc.selection.active;
  const style = doc.styleAt(r, c);
  const theme = doc.wb.theme;
  const describe: Record<Part, string | undefined> = {
    number: style.numFmt,
    alignment: `${style.align?.horizontal ?? 'General'}, ${style.align?.vertical ?? 'Bottom'} Aligned`,
    font: `${style.font.name} ${style.font.size}${style.font.bold ? ', Bold' : ''}${style.font.italic ? ', Italic' : ''}, ${style.font.color ? `#${resolveColor(style.font.color, theme, '000000')}` : 'Text 1'}`,
    border: style.border ? 'Borders' : 'No Borders',
    fill: style.fill ? 'Shaded' : 'No Shading',
    protection: style.locked === false ? 'Unlocked' : 'Locked',
  };
  return { style, describe };
});

const ok = () => {
  const n = name.value.trim();
  if (!n) return false;
  const { style } = view.value;
  const on = parts.value;
  const patch: StylePatch = {};
  if (on.number) patch.numFmt = style.numFmt;
  if (on.alignment) patch.align = { ...(style.align ?? {}) };
  if (on.font) patch.font = { ...style.font };
  if (on.border) patch.border = style.border ? { ...style.border } : null;
  if (on.fill) patch.fill = style.fill ?? null;
  if (on.protection) patch.locked = style.locked !== false;
  styles.value = [...styles.value.filter((s) => s.name !== n), { name: n, patch }];
  return true;
};
</script>

<template>
  <Dialog title="Style" :width="400" test-id="new-cell-style" :on-ok="ok" @close="emit('close')">
    <Field label="Style name:" wide>
      <input v-focus :value="name" data-testid="cell-style-name" @input="name = ($event.target as HTMLInputElement).value" />
    </Field>
    <div class="sh-label">Style Includes (By Example)</div>
    <div v-for="k in (Object.keys(parts) as Part[])" :key="k" class="sh-row">
      <Check :label="k[0].toUpperCase() + k.slice(1)" :checked="parts[k]" @change="parts = { ...parts, [k]: $event }" />
      <span class="sh-hint">{{ view.describe[k] }}</span>
    </div>
    <p class="sh-hint">Custom styles are kept in this browser and appear under Custom in the Cell Styles gallery.</p>
  </Dialog>
</template>
