<script setup lang="ts">
/** A check box with its label (the default slot, or `label`). */
import { vChecked } from '../directives';

defineProps<{ label?: string; checked: boolean; testId?: string; disabled?: boolean }>();
const emit = defineEmits<{ change: [value: boolean] }>();
</script>

<template>
  <label :class="`sh-check${disabled ? ' disabled' : ''}`">
    <input v-checked="checked" type="checkbox" :disabled="disabled" :data-testid="testId" @change="emit('change', ($event.target as HTMLInputElement).checked)" />
    <span><slot>{{ label }}</slot></span>
  </label>
</template>
