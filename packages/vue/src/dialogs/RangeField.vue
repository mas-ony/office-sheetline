<script setup lang="ts">
/**
 * A reference box with a picker button: the dialog hides while cells are selected on the grid
 * (a bar at the top shows the reference), then comes back with the selection filled in.
 */
import { rangeText, type Selection, selectionText, type Sheet } from '@sheetline/core/views';
import { computed, onUnmounted, ref, watch } from 'vue';
import { useUI } from '../context';
import { vFocus, vValue } from '../directives';
import Icon from '../Icon.vue';

const props = defineProps<{ value: string; label?: string; testId?: string; home?: Sheet; multi?: boolean; autoFocus?: boolean }>();
const emit = defineEmits<{ change: [value: string] }>();

const ui = useUI();
const picking = ref(false);
const homeSheet = props.home ?? ui.value.doc.sheet;
// The selection before picking: dialogs act on it (Text to Columns, Data Validation...), so it comes back after.
let before: Selection | null = null;
watch(picking, (on) => ui.value.layer.classList.toggle('sh-picking', on));
onUnmounted(() => {
  if (picking.value) ui.value.layer.classList.remove('sh-picking');
});
const live = computed(() => {
  if (!picking.value) return '';
  const { doc } = ui.value;
  return props.multi ? selectionText(doc, homeSheet) : rangeText(doc.bounded(doc.range, doc.sheet) ?? doc.range, doc.sheet, homeSheet);
});
const pick = () => {
  before = ui.value.doc.selectionOf(homeSheet);
  picking.value = true;
};
const done = (accept: boolean) => {
  if (accept) emit('change', live.value);
  picking.value = false;
  const { doc } = ui.value;
  if (doc.sheet !== homeSheet) doc.setActiveSheet(doc.wb.sheets.indexOf(homeSheet));
  if (before) doc.setSelection(before, homeSheet);
};
</script>

<template>
  <span class="sh-range-field">
    <input
      v-focus="autoFocus"
      v-value="value"
      :aria-label="label ?? 'Reference'"
      :data-testid="testId"
      spellcheck="false"
      autocapitalize="off"
      @input="emit('change', ($event.target as HTMLInputElement).value)"
    />
    <button type="button" class="sh-range-pick" title="Select cells" aria-label="Select cells on the sheet" :data-testid="testId ? `${testId}-pick` : undefined" @click="pick">
      <Icon name="highlight_alt" />
    </button>
    <Teleport v-if="picking" :to="ui.layer">
      <div class="sh-range-picker" role="dialog" aria-label="Select a range" data-testid="range-picker">
        <span class="sh-range-picker-label">{{ label ?? 'Select a range' }}</span>
        <input readonly :value="live" aria-label="Selected range" />
        <button type="button" class="sh-btn primary" data-testid="range-picker-ok" @click="done(true)">
          <Icon name="check" />
        </button>
        <button type="button" class="sh-btn" aria-label="Cancel" @click="done(false)">
          <Icon name="close" />
        </button>
      </div>
    </Teleport>
  </span>
</template>
