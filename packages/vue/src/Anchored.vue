<script setup lang="ts">
/** A fixed pop-up next to a rectangle, closed by clicks outside and Escape. */
import { placeNextTo } from '@sheetline/core/views';
import { onMounted, onUnmounted, onUpdated, ref, watch } from 'vue';
import { useUI } from './context';

const props = defineProps<{ rect: DOMRect; boxClass: string; testId?: string; label?: string; width?: number; below?: boolean }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { layer } = ui.value;
const el = ref<HTMLDivElement | null>(null);
// Placed after every render, as the whole interface renders again on each change.
const place = () => {
  if (el.value) placeNextTo(el.value, props.rect, props.below);
};
onMounted(place);
onUpdated(place);
watch(ui, place, { flush: 'post' });

const down = (e: PointerEvent) => {
  const t = e.target as HTMLElement;
  if (t.closest(`.${props.boxClass.split(' ')[0]}`) || t.closest('.sh-popup') || t.closest('.sh-bottom-sheet') || t.closest('.sh-sheet-backdrop') || t.closest('.sh-modal-layer')) return;
  emit('close');
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Escape' && !layer.querySelector('.sh-modal-layer')) emit('close');
};
const timer = window.setTimeout(() => {
  window.addEventListener('pointerdown', down, true);
  window.addEventListener('keydown', key);
}, 0);
onUnmounted(() => {
  window.clearTimeout(timer);
  window.removeEventListener('pointerdown', down, true);
  window.removeEventListener('keydown', key);
});
</script>

<template>
  <Teleport :to="layer">
    <div ref="el" :class="boxClass" role="dialog" :aria-label="label" :data-testid="testId" :style="width ? { width: `${width}px` } : undefined">
      <slot />
    </div>
  </Teleport>
</template>
