<script setup lang="ts">
/** Compact Format Cells for rules: font style and colour, underline, strikethrough, border, fill, number format. */
import { type CfFormat, cfHex, cfSpec, type ColorSpec } from '@sheetline/core/views';
import { useUI } from '../../context';
import Check from '../Check.vue';
import ColorButton from '../ColorButton.vue';
import Field from '../Field.vue';
import FormatPreview from './FormatPreview.vue';

const props = defineProps<{ value: CfFormat }>();
const emit = defineEmits<{ change: [format: CfFormat] }>();

const ui = useUI();
const set = (p: Partial<CfFormat>) => emit('change', { ...props.value, ...p });
const hex = (c: ColorSpec | null) => cfHex(c, ui.value.doc.wb.theme);
</script>

<template>
  <fieldset class="sh-group sh-cf-format">
    <legend>Format</legend>
    <div class="sh-row wrap">
      <Check label="Bold" :checked="!!value.bold" @change="set({ bold: $event || undefined })" />
      <Check label="Italic" :checked="!!value.italic" @change="set({ italic: $event || undefined })" />
      <Check label="Underline" :checked="!!value.underline" @change="set({ underline: $event || undefined })" />
      <Check label="Strikethrough" :checked="!!value.strike" @change="set({ strike: $event || undefined })" />
    </div>
    <div class="sh-row wrap">
      <Field label="Font color:">
        <ColorButton :value="cfSpec(value.color)" none-label="Automatic" none-swatch="#000000" @change="set({ color: hex($event) })" />
      </Field>
      <Field label="Fill:">
        <ColorButton :value="cfSpec(value.fill)" none-label="No Color" @change="set({ fill: hex($event) })" />
      </Field>
      <Field label="Border:">
        <ColorButton :value="cfSpec(value.border)" none-label="None" @change="set({ border: hex($event) })" />
      </Field>
    </div>
    <Field label="Number format:">
      <input :value="value.numFmt ?? ''" placeholder="(unchanged)" spellcheck="false" @input="set({ numFmt: ($event.target as HTMLInputElement).value || undefined })" />
    </Field>
    <div class="sh-row">
      <span class="sh-label">Preview:</span>
      <FormatPreview :format="value" />
    </div>
  </fieldset>
</template>
