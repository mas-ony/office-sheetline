<script setup lang="ts">
/** Text box that commits on blur or Enter. */
import { ref, watch } from 'vue';

const props = defineProps<{ value: string; label: string; testId?: string; inputMode?: 'decimal' | 'text' }>();
const emit = defineEmits<{ commit: [value: string] }>();

const draft = ref(props.value);
watch(
  () => props.value,
  (v) => (draft.value = v),
);
const blurred = () => {
  if (draft.value !== props.value) emit('commit', draft.value);
};
</script>

<template>
  <input
    :value="draft"
    :aria-label="label"
    :data-testid="testId"
    :inputmode="inputMode"
    @input="draft = ($event.target as HTMLInputElement).value"
    @blur="blurred"
    @keydown.enter="($event.target as HTMLInputElement).blur()"
  />
</template>
