<script setup lang="ts">
/** Height and width of an object, in inches. */
import { objectRect, setObjectSize } from '@sheetline/core/views';
import { computed } from 'vue';
import { useUI } from '../context';
import Field from '../dialogs/Field.vue';
import LazyInput from './LazyInput.vue';
import PaneSection from './PaneSection.vue';

const props = defineProps<{ id: string }>();
const ui = useUI();
const rect = computed(() => objectRect(ui.value, props.id));
const inches = (px: number) => String(Math.round((px / 96) * 100) / 100);
const setHeight = (v: string) => {
  const r = rect.value;
  if (r && Number(v) > 0) setObjectSize(ui.value, props.id, r.w, Number(v) * 96);
};
const setWidth = (v: string) => {
  const r = rect.value;
  if (r && Number(v) > 0) setObjectSize(ui.value, props.id, Number(v) * 96, r.h);
};
</script>

<template>
  <PaneSection v-if="rect" title="Size">
    <Field label="Height (in):">
      <LazyInput :value="inches(rect.h)" label="Height" input-mode="decimal" @commit="setHeight" />
    </Field>
    <Field label="Width (in):">
      <LazyInput :value="inches(rect.w)" label="Width" input-mode="decimal" @commit="setWidth" />
    </Field>
  </PaneSection>
</template>
