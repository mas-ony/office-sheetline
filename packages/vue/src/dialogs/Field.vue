<script setup lang="ts">
/** A labelled field of a dialog (the label slot, or `label`). */
defineProps<{ label?: string; wide?: boolean }>();
</script>

<template>
  <label :class="`sh-field${wide ? ' wide' : ''}`">
    <span class="sh-field-label"><slot name="label">{{ label }}</slot></span>
    <slot />
  </label>
</template>
