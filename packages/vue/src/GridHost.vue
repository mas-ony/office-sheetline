<script setup lang="ts">
/** The grid's element; the grid of each workbook is created in it. */
import type { UIState } from '@sheetline/core/views';
import { onMounted, onUnmounted, ref } from 'vue';
import { useUI } from './context';

defineProps<{ state: UIState }>();

const { ui } = useUI().value;
const host = ref<HTMLDivElement | null>(null);
let attached: HTMLDivElement | null = null;
onMounted(() => {
  attached = host.value!;
  ui.attachGrid(attached);
});
onUnmounted(() => {
  if (attached) ui.detachGrid(attached);
});
</script>

<template>
  <div ref="host" :class="`sh-grid-host${state.settings.showHScroll === false ? ' no-hscroll' : ''}${state.settings.showVScroll === false ? ' no-vscroll' : ''}`" data-testid="grid" />
</template>
