<script setup lang="ts">
/**
 * Office-style colour palette: an "Automatic" / "No Fill" entry, the theme colours with their five
 * tint rows, the standard colours, recent colours and "More Colors..." (the spectrum picker).
 * Theme picks are returned as theme colours so they follow later theme changes.
 */
import { type ColorSpec, rgbSpec, specToCss, STANDARD_COLORS, type Theme, THEME_COLUMNS, tintLabel, withRecent } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import ColorPicker from './ColorPicker.vue';
import { usePersistent } from './context';

defineProps<{
  theme: Theme;
  /** Label of the empty choice ("No Fill", "Automatic", "No Color"); omitted hides it. */
  noneLabel?: string;
  /** Colour shown on the "Automatic" swatch. */
  noneSwatch?: string;
  testId?: string;
}>();
const emit = defineEmits<{ pick: [color: ColorSpec | null] }>();

const recent = usePersistent<string[]>('recent-colors', []);
const more = shallowRef<DOMRect | null>(null);
const moreBtn = ref<HTMLButtonElement | null>(null);
const pickHex = (hex: string) => {
  recent.value = withRecent(recent.value, hex, 10);
  emit('pick', rgbSpec(hex) ?? null);
};
const fromPicker = (hex: string | null) => {
  if (hex) pickHex(hex);
  else emit('pick', null);
};
</script>

<template>
  <div class="sh-palette" :data-testid="testId" @mousedown.prevent>
    <button v-if="noneLabel" type="button" class="sh-palette-none" :data-testid="testId ? `${testId}-none` : undefined" @click="emit('pick', null)">
      <span class="sh-swatch" :style="{ background: noneSwatch ?? 'transparent' }" />{{ noneLabel }}
    </button>
    <div class="sh-palette-title">Theme Colors</div>
    <div class="sh-palette-grid theme">
      <button
        v-for="col in THEME_COLUMNS"
        :key="`b${col.theme}`"
        type="button"
        class="sh-swatch"
        :style="{ background: specToCss({ theme: col.theme }, theme) }"
        :title="col.name"
        :aria-label="col.name"
        @click="emit('pick', { theme: col.theme })"
      />
    </div>
    <div class="sh-palette-grid tints">
      <template v-for="row in [0, 1, 2, 3, 4]" :key="row">
        <button
          v-for="col in THEME_COLUMNS"
          :key="`t${row}-${col.theme}`"
          type="button"
          class="sh-swatch"
          :style="{ background: specToCss({ theme: col.theme, tint: col.tints[row] }, theme) }"
          :title="`${col.name}, ${tintLabel(col.tints[row])}`"
          :aria-label="`${col.name}, ${tintLabel(col.tints[row])}`"
          @click="emit('pick', { theme: col.theme, tint: col.tints[row] })"
        />
      </template>
    </div>
    <div class="sh-palette-title">Standard Colors</div>
    <div class="sh-palette-grid">
      <button v-for="c in STANDARD_COLORS" :key="c.hex" type="button" class="sh-swatch" :style="{ background: c.hex }" :title="c.name" :aria-label="c.name" @click="pickHex(c.hex)" />
    </div>
    <template v-if="recent.length > 0">
      <div class="sh-palette-title">Recent Colors</div>
      <div class="sh-palette-grid">
        <button v-for="hex in recent" :key="`r${hex}`" type="button" class="sh-swatch" :style="{ background: hex }" :title="hex" :aria-label="hex" @click="pickHex(hex)" />
      </div>
    </template>
    <button ref="moreBtn" type="button" class="sh-palette-more" @click="more = moreBtn!.getBoundingClientRect()">More Colors...</button>
    <ColorPicker v-if="more" :anchor="more" :value="null" :live="false" @change="fromPicker" @close="more = null" />
  </div>
</template>
