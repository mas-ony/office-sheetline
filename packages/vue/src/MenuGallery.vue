<script setup lang="ts">
/** A grid of previews in a menu (data bars, colour scales, icon sets...). */
import type { MenuItem } from '@sheetline/core/views';
import PreviewView from './PreviewView.vue';

defineProps<{ item: MenuItem }>();
const emit = defineEmits<{ close: [picked?: boolean] }>();
</script>

<template>
  <div class="sh-menu-gallery" role="group" :aria-label="item.title">
    <div v-if="item.title" class="sh-menu-header">{{ item.title }}</div>
    <div class="grid" :style="{ gridTemplateColumns: `repeat(${item.columns ?? 6}, auto)` }">
      <button
        v-for="(g, k) in item.gallery ?? []"
        :key="k"
        type="button"
        role="menuitem"
        class="sh-gallery-item"
        :title="g.title"
        :aria-label="g.title"
        :data-testid="g.testId"
        @click="
          emit('close', true);
          g.onClick();
        "
      >
        <PreviewView :preview="g.preview" />
      </button>
    </div>
  </div>
</template>
