<script setup lang="ts">
/** Height and width of the selected object (Shape and Picture Format). */
import { type Ctx, objectRect, pxToInches, setObjectSize } from '@sheetline/core/views';
import { computed, reactive, watch } from 'vue';

const props = defineProps<{ ctx: Ctx }>();

// A new object on every change: objects are moved and resized in place.
const view = computed(() => {
  const id = props.ctx.selectedObject;
  return { id, r: id ? objectRect(props.ctx, id) : null };
});
const size = reactive({ h: '', w: '' });
watch(
  [() => view.value.r?.w, () => view.value.r?.h, () => view.value.id],
  () => {
    const { r } = view.value;
    size.w = r ? pxToInches(r.w) : '';
    size.h = r ? pxToInches(r.h) : '';
  },
  { immediate: true },
);
const FIELDS = [
  { label: 'Height:', key: 'h', testId: 'object-height' },
  { label: 'Width:', key: 'w', testId: 'object-width' },
] as const;
const apply = () => {
  const nw = parseFloat(size.w);
  const nh = parseFloat(size.h);
  const { id } = view.value;
  if (id && nw > 0 && nh > 0) setObjectSize(props.ctx, id, nw * 96, nh * 96);
};
</script>

<template>
  <div v-if="view.id && view.r" class="sh-size-fields">
    <label v-for="f in FIELDS" :key="f.key" class="sh-field-inline">
      <span>{{ f.label }}</span>
      <input
        inputmode="decimal"
        :aria-label="f.label"
        :data-testid="f.testId"
        :value="size[f.key]"
        @input="size[f.key] = ($event.target as HTMLInputElement).value"
        @blur="apply"
        @keydown.enter="($event.target as HTMLInputElement).blur()"
      />
      <span>"</span>
    </label>
  </div>
</template>
