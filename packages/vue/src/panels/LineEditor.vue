<script setup lang="ts">
/** No line or a solid line with its colour, width and dash. */
import type { ColorSpec, LineDash } from '@sheetline/core/views';
import { useUI } from '../context';
import ColorButton from '../dialogs/ColorButton.vue';
import Field from '../dialogs/Field.vue';
import { vChecked } from '../directives';
import { DASHES, hexOf, type Line, specOf } from './format';
import LazyInput from './LazyInput.vue';

const props = withDefaults(defineProps<{ line: Line | null | undefined; noneLabel?: string }>(), { noneLabel: 'No line' });
const emit = defineEmits<{ change: [line: Line | null] }>();

const ui = useUI();
const solid = () => emit('change', { color: props.line?.color ?? '#595959', width: props.line?.width ?? 1 });
const setColor = (c: ColorSpec | null) => emit('change', { ...props.line!, color: hexOf(ui.value, c) ?? '#595959' });
const setWidth = (v: string) => {
  if (Number(v) > 0) emit('change', { ...props.line!, width: Math.min(20, Number(v)) });
};
const setDash = (e: Event) => {
  const v = (e.target as HTMLSelectElement).value;
  emit('change', { ...props.line!, dash: v === 'solid' ? undefined : (v as LineDash) });
};
</script>

<template>
  <div class="sh-stack tight">
    <div class="sh-row">
      <label class="sh-check"><input v-checked="!line" type="radio" @change="emit('change', null)" /> <span>{{ noneLabel }}</span></label>
      <label class="sh-check"><input v-checked="!!line" type="radio" @change="solid" /> <span>Solid line</span></label>
    </div>
    <template v-if="line">
      <Field label="Color:">
        <ColorButton :value="specOf(line.color)" none-label="Automatic" @change="setColor" />
      </Field>
      <Field label="Width (px):">
        <LazyInput :value="String(line.width)" label="Line width" input-mode="decimal" @commit="setWidth" />
      </Field>
      <Field label="Dash type:">
        <select :value="line.dash ?? 'solid'" @change="setDash">
          <option v-for="d in DASHES" :key="d.dash" :value="d.dash">{{ d.label }}</option>
        </select>
      </Field>
    </template>
  </div>
</template>
