<script setup lang="ts" generic="T extends string">
/** The tabs of a dialog. */
defineProps<{ tabs: { id: T; label: string }[]; value: T; testId?: string }>();
const emit = defineEmits<{ change: [id: T] }>();
</script>

<template>
  <div class="sh-dtabs" role="tablist" :data-testid="testId">
    <button
      v-for="t in tabs"
      :key="t.id"
      type="button"
      role="tab"
      :aria-selected="value === t.id"
      :class="value === t.id ? 'active' : ''"
      :data-testid="testId ? `${testId}-${t.id}` : undefined"
      @click="emit('change', t.id)"
    >{{ t.label }}</button>
  </div>
</template>
