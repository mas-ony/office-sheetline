<script setup lang="ts">
/** Labeled on/off switch. */
defineProps<{ label: string; checked: boolean; testId?: string }>();
const emit = defineEmits<{ change: [value: boolean] }>();
</script>

<template>
  <label class="sh-labeled-switch">
    <input type="checkbox" role="switch" :checked="checked" :data-testid="testId" @change="emit('change', ($event.target as HTMLInputElement).checked)" />
    <span>{{ label }}</span>
  </label>
</template>
