<script setup lang="ts">
/** A colour drop-down button (theme palette) for dialogs. */
import { type ColorSpec, resolveColor } from '@sheetline/core/views';
import { computed, shallowRef } from 'vue';
import { useUI } from '../context';
import Dropdown from '../Dropdown.vue';
import Icon from '../Icon.vue';
import ThemePalette from '../ThemePalette.vue';

const props = withDefaults(defineProps<{ value?: ColorSpec | null; noneLabel?: string; noneSwatch?: string; testId?: string; label?: string; disabled?: boolean }>(), {
  value: undefined,
  noneLabel: 'Automatic',
});
const emit = defineEmits<{ change: [color: ColorSpec | null] }>();

const ui = useUI();
const anchor = shallowRef<DOMRect | null>(null);
const css = computed(() => (props.value ? `#${resolveColor(props.value, ui.value.doc.wb.theme, '000000')}` : (props.noneSwatch ?? 'transparent')));
const pick = (c: ColorSpec | null) => {
  emit('change', c);
  anchor.value = null;
};
</script>

<template>
  <button type="button" class="sh-color-button" :aria-label="label ?? 'Color'" :data-testid="testId" :disabled="disabled" @click="anchor = ($event.currentTarget as HTMLElement).getBoundingClientRect()">
    <span class="sh-color-chip" :style="{ background: css }" />
    <span class="sh-color-name">{{ value ? css.toUpperCase() : noneLabel }}</span>
    <Icon name="arrow_drop_down" />
  </button>
  <Dropdown v-if="anchor" :anchor="anchor" extra-class="nested" :test-id="testId ? `${testId}-palette` : undefined" @close="anchor = null">
    <ThemePalette :theme="ui.doc.wb.theme" :none-label="noneLabel" :none-swatch="noneSwatch" @pick="pick" />
  </Dropdown>
</template>
