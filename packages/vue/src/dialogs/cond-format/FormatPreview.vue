<script setup lang="ts">
/** How a format looks (preview cell of the dialogs and the manager). */
import type { CfFormat } from '@sheetline/core/views';
import { computed } from 'vue';

const props = withDefaults(defineProps<{ format: CfFormat | null; text?: string }>(), { text: 'AaBbCcYyZz' });
const look = computed(() => {
  const f = props.format;
  if (!f) return {};
  return {
    color: f.color,
    background: f.fill,
    fontWeight: f.bold ? 700 : undefined,
    fontStyle: f.italic ? 'italic' : undefined,
    textDecoration: [f.underline ? 'underline' : '', f.strike ? 'line-through' : ''].join(' ').trim() || undefined,
    outline: f.border ? `1px solid ${f.border}` : undefined,
  };
});
</script>

<template>
  <span v-if="!format" class="sh-cf-preview none">No Format Set</span>
  <span v-else class="sh-cf-preview" :style="look">{{ text }}</span>
</template>
