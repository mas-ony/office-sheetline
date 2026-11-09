<script setup lang="ts">
/** A radio button with its label (the default slot, or `label`). */
import { vChecked } from '../directives';

defineProps<{ label?: string; checked: boolean; name: string; testId?: string; disabled?: boolean }>();
const emit = defineEmits<{ change: [] }>();
</script>

<template>
  <label :class="`sh-check${disabled ? ' disabled' : ''}`">
    <input v-checked="checked" type="radio" :name="name" :disabled="disabled" :data-testid="testId" @change="emit('change')" />
    <span><slot>{{ label }}</slot></span>
  </label>
</template>
