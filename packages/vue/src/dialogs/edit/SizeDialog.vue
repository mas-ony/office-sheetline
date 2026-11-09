<script setup lang="ts">
/** Row Height, Column Width, Standard Height and Standard Width. */
import { cmd, ptToPx, pxToPt, widthCharsToPx, widthPxToChars } from '@sheetline/core/views';
import { ref } from 'vue';
import { useUI } from '../../context';
import { vFocus } from '../../directives';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';

const props = defineProps<{ what: 'row-height' | 'col-width' | 'default-height' | 'default-width' }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { doc } = ui.value;
const sheet = doc.sheet;
const { r, c } = doc.selection.active;
const initial =
  props.what === 'row-height' ? pxToPt(sheet.rowHeightPx(r)) : props.what === 'default-height' ? sheet.defaultRowHeightPt : props.what === 'col-width' ? widthPxToChars(sheet.colWidthPx(c)) : widthPxToChars(sheet.defaultColWidthPx);
const value = ref(String(Math.round(initial * 100) / 100));
const title = { 'row-height': 'Row Height', 'col-width': 'Column Width', 'default-height': 'Standard Height', 'default-width': 'Standard Width' }[props.what];

const ok = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  const { what } = props;
  const v = Number(value.value);
  const max = what.includes('height') ? 409 : 255;
  if (!Number.isFinite(v) || v < 0 || v > max) {
    ctx.notify(`Enter a number between 0 and ${max}.`);
    return false;
  }
  cmd.commit(ctx);
  const lines = (axis: 'row' | 'col') => {
    const out = new Set<number>();
    for (const x of doc.selection.ranges) {
      const [a, b] = axis === 'row' ? [x.r0, Math.min(x.r1, x.r0 + 100000)] : [x.c0, x.c1];
      for (let i = a; i <= b; i++) out.add(i);
    }
    return [...out];
  };
  if (what === 'row-height') doc.setRowHeight(lines('row'), v);
  else if (what === 'col-width') doc.setColWidth(lines('col'), widthCharsToPx(v));
  else if (what === 'default-height') doc.setDefaultSize({ rowHeightPt: v });
  else doc.setDefaultSize({ colWidthPx: widthCharsToPx(v) });
  return true;
};
</script>

<template>
  <Dialog :title="title" :width="280" test-id="size-dialog" :on-ok="ok" @close="emit('close')">
    <Field :label="`${title}:`">
      <input v-focus :value="value" inputmode="decimal" data-testid="size-input" :style="{ width: '90px' }" @input="value = ($event.target as HTMLInputElement).value" />
    </Field>
    <p v-if="what === 'row-height'" class="sh-hint">Points ({{ Math.round(ptToPx(Number(value) || 0)) }} pixels).</p>
  </Dialog>
</template>
