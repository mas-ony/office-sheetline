<script setup lang="ts">
/** A menu as a bottom sheet (phones): submenus open in place with a back button. */
import type { MenuItem } from '@sheetline/core/views';
import { computed, shallowRef } from 'vue';
import Icon from './Icon.vue';
import MenuGallery from './MenuGallery.vue';
import PreviewView from './PreviewView.vue';

const props = defineProps<{ items: MenuItem[]; title?: string; testId?: string }>();
const emit = defineEmits<{ close: [picked?: boolean] }>();

const stack = shallowRef<{ title?: string; items: MenuItem[] }[]>([{ title: props.title, items: props.items }]);
const top = computed(() => stack.value[stack.value.length - 1]);
const pick = (it: MenuItem) => {
  if (it.submenu) {
    stack.value = [...stack.value, { title: it.title, items: it.submenu }];
    return;
  }
  if (!it.keepOpen) emit('close', true);
  it.onClick?.();
};
</script>

<template>
  <div
    class="sh-sheet-backdrop"
    @pointerdown="
      $event.preventDefault();
      emit('close');
    "
  />
  <div class="sh-bottom-sheet" role="menu" :aria-label="top.title" :data-testid="testId">
    <div class="sh-sheet-grip" />
    <div v-if="stack.length > 1 || top.title" class="sh-sheet-title">
      <button v-if="stack.length > 1" type="button" aria-label="Back" @click="stack = stack.slice(0, -1)">
        <Icon name="arrow_right" class="sh-flip" />
      </button>
      <span v-else :style="{ width: '8px' }" />
      <span>{{ top.title }}</span>
    </div>
    <template v-for="(it, i) in top.items" :key="i">
      <div v-if="it.type === 'line'" class="sh-menu-sep" role="separator" />
      <div v-else-if="it.type === 'header'" class="sh-menu-header">{{ it.title }}</div>
      <MenuGallery v-else-if="it.type === 'gallery'" :item="it" @close="(picked) => emit('close', picked)" />
      <button v-else type="button" role="menuitem" :data-testid="it.testId" :title="it.tip" class="sh-menu-item" :disabled="it.disabled" @click="pick(it)">
        <span class="icon-slot">
          <PreviewView v-if="it.swatch" :preview="it.swatch" />
          <Icon v-else-if="it.icon" :name="it.icon" />
          <Icon v-else-if="it.checked" name="check" class="check" />
        </span>
        <span class="title">{{ it.title }}</span>
        <Icon v-if="it.submenu" name="arrow_right" class="sub" />
      </button>
    </template>
  </div>
</template>
