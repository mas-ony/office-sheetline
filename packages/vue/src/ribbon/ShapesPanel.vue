<script setup lang="ts">
/** Insert > Shapes and Shape Format > Change Shape. */
import { SHAPE_GALLERY, type ShapeGalleryItem, shapeIconSvg } from '@sheetline/core/views';

defineProps<{ testId: string; itemTestIds?: boolean }>();
const emit = defineEmits<{ close: []; pick: [item: ShapeGalleryItem] }>();

const choose = (it: ShapeGalleryItem) => {
  emit('close');
  emit('pick', it);
};
// The icons (v-html) are generated from the shape geometry, not user content.
const itemTestId = (it: ShapeGalleryItem) => `shape-${it.preset}${it.arrowEnd ? (it.arrowStart ? '-both' : '-end') : ''}`;
</script>

<template>
  <div class="sh-shapes-panel" :data-testid="testId">
    <div v-for="group in SHAPE_GALLERY" :key="group.title">
      <div class="sh-panel-title">{{ group.title }}</div>
      <div class="sh-shape-grid">
        <button
          v-for="(it, i) in group.items"
          :key="`${it.preset}-${i}`"
          type="button"
          :title="it.label"
          :aria-label="it.label"
          :data-testid="itemTestIds ? itemTestId(it) : undefined"
          @click="choose(it)"
          v-html="shapeIconSvg(it)"
        />
      </div>
    </div>
  </div>
</template>
