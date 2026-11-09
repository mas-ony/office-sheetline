<script setup lang="ts">
/** Move Chart: to a new chart sheet, or as an object in another sheet. */
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Radio from '../Radio.vue';

const props = defineProps<{ chartId: string }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { doc } = ui.value;
const sheets = computed(() => ui.value.doc.wb.sheets.filter((s) => s.state === 'visible' && !s.chartsheet));
const where = ref<'new' | 'object'>('object');
let k = 1;
while (doc.wb.sheets.some((s) => s.name.toLowerCase() === `chart${k}`)) k++;
const newName = ref(`Chart${k}`);
const target = ref(doc.sheet.name);

const ok = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  const from = doc.sheet;
  if (where.value === 'new') {
    const created = doc.addSheet(newName.value.trim() || undefined, doc.activeIndex);
    if (!created) {
      ctx.notify('That sheet name is not valid or is already used.');
      return false;
    }
    doc.moveChartToSheet(props.chartId, created, from);
    doc.setActiveSheet(doc.wb.sheets.indexOf(created));
    return true;
  }
  const t = doc.wb.sheetByName(target.value);
  if (t && t !== from) {
    doc.moveChartToSheet(props.chartId, t, from);
    doc.setActiveSheet(doc.wb.sheets.indexOf(t));
  }
  ctx.selectObject(null);
  return true;
};
</script>

<template>
  <Dialog title="Move Chart" :width="420" test-id="move-chart" :on-ok="ok" @close="emit('close')">
    <p class="sh-lead">Choose where you want the chart to be placed:</p>
    <div class="sh-row">
      <Radio name="move-chart" label="New sheet:" :checked="where === 'new'" test-id="move-chart-new" @change="where = 'new'" />
      <input :value="newName" :disabled="where !== 'new'" aria-label="New sheet name" @input="newName = ($event.target as HTMLInputElement).value" />
    </div>
    <div class="sh-row">
      <Radio name="move-chart" label="Object in:" :checked="where === 'object'" @change="where = 'object'" />
      <select :value="target" :disabled="where !== 'object'" aria-label="Sheet" data-testid="move-chart-sheet" @change="target = ($event.target as HTMLSelectElement).value">
        <option v-for="s in sheets" :key="s.uid" :value="s.name">{{ s.name }}</option>
      </select>
    </div>
  </Dialog>
</template>
