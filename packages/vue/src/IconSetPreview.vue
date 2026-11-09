<script setup lang="ts">
/** The icons of a conditional formatting icon set, drawn on a canvas. */
import { type IconSetName, paintIconSet } from '@sheetline/core/views';
import { onMounted, ref, watch } from 'vue';

const props = defineProps<{ set: IconSetName }>();
const canvas = ref<HTMLCanvasElement | null>(null);
const paint = () => {
  if (canvas.value) paintIconSet(canvas.value, props.set);
};
onMounted(paint);
watch(() => props.set, paint, { flush: 'post' });
</script>

<template>
  <canvas ref="canvas" :style="{ width: `${Number(set[0]) * 16}px`, height: '16px' }" aria-hidden="true" />
</template>
