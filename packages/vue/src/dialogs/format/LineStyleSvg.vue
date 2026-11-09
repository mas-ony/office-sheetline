<script setup lang="ts">
/** A border line style drawn as a short line ("None" for no line). */
import { BORDER_LINE_PREVIEW, type BorderLineStyle } from '@sheetline/core/views';
import { computed } from 'vue';

const props = withDefaults(defineProps<{ lineStyle: BorderLineStyle | 'none'; color?: string; w?: number }>(), { color: '#000000', w: 70 });
const spec = computed(() => (props.lineStyle === 'none' ? '' : BORDER_LINE_PREVIEW[props.lineStyle]));
const stroke = computed(() => {
  const [width, ...dash] = spec.value.split(' ').map(Number);
  return { width, dash: dash.length ? dash.join(' ') : undefined };
});
</script>

<template>
  <span v-if="lineStyle === 'none'" class="sh-line-none">None</span>
  <svg v-else-if="spec === 'd'" :width="w" :height="10">
    <line x1="2" y1="3.5" :x2="w - 2" y2="3.5" :stroke="color" stroke-width="1" />
    <line x1="2" y1="6.5" :x2="w - 2" y2="6.5" :stroke="color" stroke-width="1" />
  </svg>
  <svg v-else :width="w" :height="10">
    <line x1="2" y1="5" :x2="w - 2" y2="5" :stroke="color" :stroke-width="stroke.width" :stroke-dasharray="stroke.dash" />
  </svg>
</template>
