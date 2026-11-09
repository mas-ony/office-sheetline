<script setup lang="ts">
/** A small picture of a menu item or gallery entry. */
import type { Preview } from '@sheetline/core/views';
import Icon from './Icon.vue';
import IconSetPreview from './IconSetPreview.vue';

defineProps<{ preview: Preview }>();
</script>

<template>
  <span v-if="preview.kind === 'svg'" class="sh-svg" :style="{ width: `${preview.width}px`, height: `${preview.height}px` }" v-html="preview.markup" />
  <span v-else-if="preview.kind === 'color'" class="sh-menu-swatch" :style="{ background: preview.css }" />
  <span v-else-if="preview.kind === 'theme'" class="sh-theme-swatch">
    <i v-for="(c, i) in preview.colors" :key="i" :style="{ background: c }" />
  </span>
  <span v-else-if="preview.kind === 'font'" class="sh-font-aa" :style="{ fontFamily: preview.family }">Aa</span>
  <span v-else-if="preview.kind === 'palette'" class="sh-palette-row">
    <i v-for="(c, i) in preview.colors" :key="i" :style="{ background: c }" />
  </span>
  <span v-else-if="preview.kind === 'paste'" class="sh-paste-icon">
    <Icon name="content_paste" class="base" />
    <Icon v-if="preview.badge" :name="preview.badge" class="badge" />
  </span>
  <IconSetPreview v-else-if="preview.kind === 'icon-set'" :set="preview.set" />
</template>
