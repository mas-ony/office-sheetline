<script setup lang="ts">
/** The colour spectrum of the colour picker: drag on it to pick. */
import { paintSpectrum, spectrumPick } from '@sheetline/core/views';
import { onMounted, ref, shallowRef } from 'vue';

const emit = defineEmits<{ pick: [hex: string] }>();
const canvas = ref<HTMLCanvasElement | null>(null);
const point = shallowRef<{ x: number; y: number } | null>(null);
onMounted(() => {
  if (canvas.value) paintSpectrum(canvas.value);
});
const pick = (e: PointerEvent) => {
  const c = canvas.value;
  if (!c) return;
  const { x, y, hex } = spectrumPick(c, e.clientX, e.clientY);
  point.value = { x, y };
  if (hex) emit('pick', hex);
};
const down = (e: PointerEvent) => {
  (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
  pick(e);
};
const move = (e: PointerEvent) => {
  if (e.buttons) pick(e);
};
</script>

<template>
  <div class="sh-color-spectrum" @pointerdown="down" @pointermove="move">
    <canvas ref="canvas" width="232" height="160" />
    <div v-if="point" class="sh-color-point" :style="{ left: `${point.x}px`, top: `${point.y}px` }" />
  </div>
</template>
