<script setup lang="ts">
/** An icon of the editor's set, as an inline SVG that takes the text colour (bundled package assets, not user content). */
import { type IconName, iconSvg } from '@sheetline/core/views';

defineProps<{ name: IconName; size?: number; color?: string }>();
</script>

<template>
  <span
    class="mi"
    aria-hidden="true"
    :style="size || color ? { width: size ? `${size}px` : undefined, height: size ? `${size}px` : undefined, color } : undefined"
    v-html="iconSvg(name)"
  />
</template>
