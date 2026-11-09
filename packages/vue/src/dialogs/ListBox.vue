<script setup lang="ts" generic="T">
/** A selectable list (list boxes of Excel dialogs); the default slot draws an item. */
import { onMounted, ref, watch } from 'vue';

const props = withDefaults(defineProps<{ items: T[]; value: T | null; keyOf: (v: T) => string; testId?: string; height?: number; onActivate?: (v: T) => void; label?: string }>(), {
  height: 160,
});
const emit = defineEmits<{ change: [value: T] }>();
defineSlots<{ default(props: { item: T }): unknown }>();

const el = ref<HTMLDivElement | null>(null);
const scrollToSelected = () => el.value?.querySelector<HTMLElement>('.selected')?.scrollIntoView({ block: 'nearest' });
onMounted(scrollToSelected);
watch(() => props.value, scrollToSelected, { flush: 'post' });

const selected = (it: T) => props.value !== null && props.keyOf(props.value) === props.keyOf(it);
const key = (e: KeyboardEvent) => {
  const { items, value, keyOf, onActivate } = props;
  const index = value === null ? -1 : items.findIndex((x) => keyOf(x) === keyOf(value));
  if (e.key === 'ArrowDown' && items.length) {
    e.preventDefault();
    emit('change', items[Math.min(items.length - 1, index + 1)]);
  } else if (e.key === 'ArrowUp' && items.length) {
    e.preventDefault();
    emit('change', items[Math.max(0, index - 1)]);
  } else if (e.key === 'Enter' && value !== null && onActivate) {
    e.preventDefault();
    onActivate(value);
  }
};
</script>

<template>
  <div ref="el" class="sh-listbox" role="listbox" :aria-label="label" tabindex="0" :data-testid="testId" :style="{ height: `${height}px` }" @keydown="key">
    <div
      v-for="it in items"
      :key="keyOf(it)"
      role="option"
      :aria-selected="selected(it)"
      :class="`sh-listbox-item${selected(it) ? ' selected' : ''}`"
      @click="emit('change', it)"
      @dblclick="onActivate?.(it)"
    >
      <slot :item="it" />
    </div>
  </div>
</template>
