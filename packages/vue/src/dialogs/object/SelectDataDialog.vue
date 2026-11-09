<script setup lang="ts">
/** Select Data Source: the chart's range, its series (add, edit, remove, reorder) and category labels. */
import { absRef, applySubtype, chartFromRange, type ChartSeries, chartSeriesName, localRef, parseRangeText, readChartData, subtypeOf, switchedSeries } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Icon from '../../Icon.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import ListBox from '../ListBox.vue';
import RangeField from '../RangeField.vue';

const props = defineProps<{ chartId: string }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const chart = computed(() => ui.value.doc.sheet.charts.find((c) => c.id === props.chartId));
const first = chart.value;
const home = ui.value.doc.sheet;
const range = ref(first?.range ? localRef(first.range, home) : '');
const series = shallowRef<ChartSeries[]>(first?.series.map((s) => ({ ...s })) ?? []);
const categories = ref(first?.categories ? localRef(first.categories, home) : '');
const seriesIn = ref(first?.seriesIn ?? 'columns');
const headers = ref(first?.headers ?? true);
const sel = ref<number | null>(series.value.length ? 0 : null);
const edit = shallowRef<{ index: number | null; name: string; values: string } | null>(null);
const empty = ref<'gap' | 'zero' | 'span'>('gap');
const indexes = computed(() => series.value.map((_, i) => i));
const keyOf = (i: number) => String(i);

const qualify = (text: string) => {
  const hit = parseRangeText(ui.value.doc, text);
  return hit ? absRef(hit.sheet, hit.range) : null;
};
const setRange = (text: string) => {
  range.value = text;
  const c = chart.value;
  const { doc } = ui.value;
  const hit = parseRangeText(doc, text);
  if (!c || !hit) return;
  const def = chartFromRange(doc, hit.sheet, hit.range, c.type);
  series.value = applySubtype(def, subtypeOf(c)).series;
  categories.value = def.categories ? localRef(def.categories, doc.sheet) : '';
  seriesIn.value = def.seriesIn;
  headers.value = def.headers;
  sel.value = 0;
};
const categoryHint = () => {
  const c = chart.value;
  if (!c) return '';
  const { doc } = ui.value;
  return `${readChartData(doc, { ...c, categories: qualify(categories.value) ?? undefined }).categories.slice(0, 6).join(', ')}${readChartData(doc, c).categories.length > 6 ? ', …' : ''}`;
};

const ok = () => {
  const ctx = ui.value;
  const c = chart.value;
  if (!c) return;
  const cat = categories.value.trim() ? qualify(categories.value) : undefined;
  if (categories.value.trim() && !cat) {
    ctx.notify('The category labels reference isn\'t valid.');
    return false;
  }
  const hit = parseRangeText(ctx.doc, range.value);
  ctx.doc.updateChart(c.id, { range: hit ? absRef(hit.sheet, hit.range) : c.range, series: series.value, categories: cat ?? undefined, seriesIn: seriesIn.value, headers: headers.value });
  return true;
};
const switchRowColumn = () => {
  const c = chart.value;
  if (!c) return;
  const { doc } = ui.value;
  const next = seriesIn.value === 'rows' ? 'columns' : 'rows';
  const sw = switchedSeries(doc, { ...c, range: qualify(range.value) ?? c.range, seriesIn: seriesIn.value, headers: headers.value, categories: categories.value ? (qualify(categories.value) ?? undefined) : undefined, series: series.value }, next);
  if (sw.series) {
    series.value = applySubtype({ ...c, series: sw.series }, subtypeOf(c)).series;
    categories.value = sw.categories ? localRef(sw.categories, doc.sheet) : '';
    seriesIn.value = next;
    headers.value = !!sw.headers;
  }
};
const editSelected = () => {
  const i = sel.value;
  if (i === null) return;
  const s = series.value[i];
  const { doc } = ui.value;
  edit.value = { index: i, name: s.nameRef ? localRef(s.nameRef, doc.sheet) : (s.name ?? ''), values: localRef(s.values, doc.sheet) };
};
const removeSelected = () => {
  const i = sel.value;
  if (i === null) return;
  const count = series.value.length;
  series.value = series.value.filter((_, k) => k !== i);
  sel.value = count > 1 ? Math.max(0, i - 1) : null;
};
const moveUp = () => {
  const i = sel.value;
  if (!i) return;
  const n = series.value.slice();
  [n[i - 1], n[i]] = [n[i], n[i - 1]];
  series.value = n;
  sel.value = i - 1;
};
const moveDown = () => {
  const i = sel.value;
  if (i === null) return;
  const n = series.value.slice();
  [n[i + 1], n[i]] = [n[i], n[i + 1]];
  series.value = n;
  sel.value = i + 1;
};
const setEdit = (patch: { name?: string; values?: string }) => {
  if (edit.value) edit.value = { ...edit.value, ...patch };
};
const editOk = () => {
  const e = edit.value;
  if (!e) return;
  const ctx = ui.value;
  const values = qualify(e.values);
  if (!values) {
    ctx.notify('The series values reference isn\'t valid.');
    return false;
  }
  const nameRef = e.name.trim() && parseRangeText(ctx.doc, e.name) ? (qualify(e.name) ?? undefined) : undefined;
  const count = series.value.length;
  const s: ChartSeries = { ...(e.index !== null ? series.value[e.index] : {}), values, nameRef, name: nameRef ? undefined : e.name || `Series ${count + 1}` };
  if (e.index === null) {
    series.value = [...series.value, s];
    sel.value = count;
  } else series.value = series.value.map((x, i) => (i === e.index ? s : x));
  return true;
};
</script>

<template>
  <Dialog v-if="chart" title="Select Data Source" :width="600" test-id="select-data" :on-ok="ok" @close="emit('close')">
    <Field label="Chart data range:" wide>
      <RangeField :value="range" label="Chart data range" test-id="select-data-range" @change="setRange" />
    </Field>
    <div class="sh-row center">
      <button type="button" class="sh-btn" data-testid="select-data-switch" @click="switchRowColumn"><Icon name="swap_horiz" /> Switch Row/Column</button>
    </div>
    <div class="sh-two-col even">
      <fieldset class="sh-group">
        <legend>Legend Entries (Series)</legend>
        <div class="sh-row">
          <button type="button" class="sh-btn" data-testid="series-add" @click="edit = { index: null, name: '', values: '' }">Add</button>
          <button type="button" class="sh-btn" :disabled="sel === null" @click="editSelected">Edit</button>
          <button type="button" class="sh-btn" :disabled="sel === null" data-testid="series-remove" @click="removeSelected">Remove</button>
          <button type="button" class="sh-btn icon" aria-label="Move up" :disabled="!sel" @click="moveUp">▲</button>
          <button type="button" class="sh-btn icon" aria-label="Move down" :disabled="sel === null || sel >= series.length - 1" @click="moveDown">▼</button>
        </div>
        <ListBox v-slot="{ item }" :items="indexes" :value="sel" :key-of="keyOf" :height="150" test-id="series-list" label="Series" @change="sel = $event">{{ chartSeriesName(ui.doc, series[item], item) }}</ListBox>
      </fieldset>
      <fieldset class="sh-group">
        <legend>Horizontal (Category) Axis Labels</legend>
        <RangeField :value="categories" label="Axis label range" test-id="select-data-categories" @change="categories = $event" />
        <div class="sh-hint">Values: {{ categoryHint() }}</div>
      </fieldset>
    </div>
    <Field label="Show empty cells as:">
      <select :value="empty" @change="empty = ($event.target as HTMLSelectElement).value as typeof empty">
        <option value="gap">Gaps</option>
        <option value="zero">Zero</option>
        <option value="span">Connect data points with line</option>
      </select>
    </Field>
  </Dialog>
  <Dialog v-if="chart && edit" title="Edit Series" :width="380" test-id="edit-series" :on-ok="editOk" @close="edit = null">
    <Field label="Series name:" wide>
      <RangeField :value="edit.name" label="Series name" test-id="edit-series-name" @change="setEdit({ name: $event })" />
    </Field>
    <Field label="Series values:" wide>
      <RangeField :value="edit.values" label="Series values" test-id="edit-series-values" @change="setEdit({ values: $event })" />
    </Field>
  </Dialog>
</template>
