<script setup lang="ts">
/** A number input of a dialog ('' when empty). */
import { vNumberValue } from '../directives';

withDefaults(defineProps<{ value: number | ''; min?: number; max?: number; step?: number; testId?: string; label?: string; width?: number; disabled?: boolean }>(), { step: 1, width: 80 });
const emit = defineEmits<{ change: [value: number | ''] }>();
const changed = (e: Event) => {
  const v = (e.target as HTMLInputElement).value;
  emit('change', v === '' ? '' : Number(v));
};
</script>

<template>
  <input
    v-number-value="value"
    type="number"
    class="sh-number"
    inputmode="decimal"
    :aria-label="label"
    :data-testid="testId"
    :min="min"
    :max="max"
    :step="step"
    :disabled="disabled"
    :style="{ width: `${width}px` }"
    @input="changed"
  />
</template>
