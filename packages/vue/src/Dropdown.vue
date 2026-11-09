<script setup lang="ts">
/** The surface of a drop-down (ribbon menus, galleries, lists), or a bottom sheet on phones. */
import { onMounted, onUnmounted, ref, watch } from 'vue';
import { useUI } from './context';

const props = defineProps<{
  anchor: DOMRect;
  /** Extra classes of the surface. */
  extraClass?: string;
  testId?: string;
  /** Bottom sheet presentation (phones). */
  sheet?: boolean;
  title?: string;
}>();
const emit = defineEmits<{ close: [] }>();

const { layer } = useUI().value;
const el = ref<HTMLDivElement | null>(null);

const place = () => {
  const box = el.value;
  if (!box || props.sheet) return;
  const w = box.offsetWidth;
  const h = box.offsetHeight;
  const { anchor } = props;
  let x = anchor.left;
  let y = anchor.bottom + 1;
  if (x + w > window.innerWidth - 4) x = Math.max(4, window.innerWidth - w - 4);
  if (y + h > window.innerHeight - 4) y = Math.max(4, Math.min(anchor.top - h - 1, window.innerHeight - h - 4));
  box.style.left = `${x}px`;
  box.style.top = `${y}px`;
};
onMounted(place);
watch(() => [props.anchor, props.sheet], place, { flush: 'post' });

const down = (e: PointerEvent) => {
  const t = e.target as HTMLElement | null;
  if (!t || el.value?.contains(t) || t.closest('.sh-color, .sh-popup, .sh-bottom-sheet, .sh-modal-layer, .sh-dropdown.nested')) return;
  emit('close');
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Escape') {
    e.preventDefault();
    emit('close');
  }
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
const backdrop = (e: PointerEvent) => {
  e.preventDefault();
  emit('close');
};
</script>

<template>
  <Teleport :to="layer">
    <template v-if="sheet">
      <div class="sh-sheet-backdrop" @pointerdown="backdrop" />
      <div ref="el" :class="['sh-bottom-sheet sh-sheet-panel', extraClass]" :data-testid="testId" role="dialog" :aria-label="title">
        <div class="sh-sheet-grip" />
        <div v-if="title" class="sh-sheet-title"><span :style="{ width: '8px' }" /><span>{{ title }}</span></div>
        <slot />
      </div>
    </template>
    <div v-else ref="el" :class="['sh-dropdown', extraClass]" :data-testid="testId" @contextmenu.prevent>
      <slot />
    </div>
  </Teleport>
</template>
