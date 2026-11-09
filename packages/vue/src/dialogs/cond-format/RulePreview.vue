<script setup lang="ts">
/** How a rule looks: its colour scale, data bar, icons or format. */
import { cfIconUrl, type CfRule, iconCount } from '@sheetline/core/views';
import { computed } from 'vue';
import FormatPreview from './FormatPreview.vue';

const props = defineProps<{ rule: CfRule | null }>();
const icons = computed(() => {
  const r = props.rule;
  if (r?.type !== 'iconSet') return [];
  const n = iconCount(r.iconSet);
  return Array.from({ length: n }, (_, i) => cfIconUrl(r.iconSet, r.reverse ? i : n - 1 - i));
});
</script>

<template>
  <span v-if="!rule" class="sh-cf-preview none">Unknown rule</span>
  <span v-else-if="rule.type === 'colorScale'" class="sh-cf-preview scale" :style="{ background: `linear-gradient(90deg, ${rule.points.map((p) => p.color).join(', ')})` }" />
  <span v-else-if="rule.type === 'dataBar'" class="sh-cf-preview bar" :style="{ background: rule.gradient ? `linear-gradient(90deg, ${rule.color} 0, #fff 70%)` : rule.color, width: '60px' }" />
  <span v-else-if="rule.type === 'iconSet'" class="sh-cf-preview icons">
    <img v-for="(src, i) in icons" :key="i" alt="" :src="src" width="16" height="16" />
  </span>
  <FormatPreview v-else :format="rule.format" />
</template>
