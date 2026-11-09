<script setup lang="ts">
/** A threshold of a colour scale or data bar: its type and value. */
import { type Cfvo, CFVO_TYPE_OPTIONS, type CfvoType } from '@sheetline/core/views';
import { computed } from 'vue';

const props = defineProps<{ value: Cfvo; which: 'min' | 'mid' | 'max' }>();
const emit = defineEmits<{ change: [value: Cfvo] }>();

const types = computed(() => CFVO_TYPE_OPTIONS.filter((t) => (props.which === 'min' ? t.type !== 'max' : props.which === 'max' ? t.type !== 'min' : t.type !== 'min' && t.type !== 'max')));
const fixed = computed(() => props.value.type === 'min' || props.value.type === 'max');
const pickType = (e: Event) => {
  const t = (e.target as HTMLSelectElement).value as CfvoType;
  emit('change', { type: t, value: t === 'min' || t === 'max' ? undefined : (props.value.value ?? (t === 'percentile' || t === 'percent' ? '50' : '0')) });
};
</script>

<template>
  <div class="sh-stack tight">
    <select :value="value.type" @change="pickType">
      <option v-for="t in types" :key="t.type" :value="t.type">{{ t.label }}</option>
    </select>
    <input :value="fixed ? (value.type === 'min' ? '(Lowest value)' : '(Highest value)') : (value.value ?? '')" :disabled="fixed" @input="emit('change', { ...value, value: ($event.target as HTMLInputElement).value })" />
  </div>
</template>
