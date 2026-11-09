<script setup lang="ts">
/** The format part of a rule: a preset list plus "Custom Format..." (font, border, fill, number format). */
import { type CfFormat, HIGHLIGHT_PRESETS } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import FormatEditor from './FormatEditor.vue';

const props = defineProps<{ value: CfFormat; testId?: string }>();
const emit = defineEmits<{ change: [format: CfFormat] }>();

const custom = ref(false);
const preset = computed(() => HIGHLIGHT_PRESETS.findIndex((p) => JSON.stringify(p.format) === JSON.stringify(props.value)));
const pick = (e: Event) => {
  const v = (e.target as HTMLSelectElement).value;
  if (v === 'custom') custom.value = true;
  else {
    custom.value = false;
    emit('change', { ...HIGHLIGHT_PRESETS[Number(v)].format });
  }
};
</script>

<template>
  <select :data-testid="testId" :value="custom || preset < 0 ? 'custom' : String(preset)" @change="pick">
    <option v-for="(p, i) in HIGHLIGHT_PRESETS" :key="p.label" :value="i">{{ p.label }}</option>
    <option value="custom">Custom Format...</option>
  </select>
  <FormatEditor v-if="custom || preset < 0" :value="value" @change="emit('change', $event)" />
</template>
