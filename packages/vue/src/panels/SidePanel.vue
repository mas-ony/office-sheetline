<script setup lang="ts">
/** A side panel at the right of the grid, with its title, actions and close button. */
import { useUI } from '../context';
import Icon from '../Icon.vue';

defineProps<{ title: string; testId: string }>();
const ui = useUI();
</script>

<template>
  <aside class="sh-side-panel" :data-testid="testId" :aria-label="title">
    <div class="sh-side-head">
      <span class="sh-side-title">{{ title }}</span>
      <slot name="actions" />
      <button type="button" class="sh-icon-btn" aria-label="Close panel" data-testid="side-panel-close" @click="ui.setPanel(null)">
        <Icon name="close" />
      </button>
    </div>
    <div class="sh-side-body"><slot /></div>
  </aside>
</template>
