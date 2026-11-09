<script setup lang="ts">
/** Insert Chart, and Change Chart Type: categories, subtypes and previews of the data both ways. */
import {
  applySubtype,
  CHART_CATEGORIES,
  type ChartCategory,
  chartFromRange,
  chartPreviewSvg,
  chartSeriesName,
  chartSource,
  type ChartSubtype,
  COMBO_SERIES_KINDS,
  comboKinds,
  forcedChart,
  insertChart,
  ONE_WAY_CHARTS,
  type SeriesKind,
  subtypeOf,
} from '@sheetline/core/views';
import { computed, onMounted, ref, shallowRef, watch } from 'vue';
import { useUI } from '../../context';
import { vChecked } from '../../directives';
import Icon from '../../Icon.vue';
import Dialog from '../Dialog.vue';

const props = defineProps<{ category?: string; changeId?: string }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const doc = computed(() => ui.value.doc);
const existing = computed(() => (props.changeId ? (ui.value.doc.sheet.charts.find((c) => c.id === props.changeId) ?? null) : null));
const first = existing.value;
const cat = ref((first ? CHART_CATEGORIES.find((c) => c.subtypes.includes(subtypeOf(first)))?.id : props.category) ?? 'column');
const catDef = computed(() => CHART_CATEGORIES.find((c) => c.id === cat.value) ?? CHART_CATEGORIES[0]);
const subId = ref<string>(first ? subtypeOf(first).id : catDef.value.subtypes[0].id);
const sub = computed<ChartSubtype>(() => catDef.value.subtypes.find((s) => s.id === subId.value) ?? catDef.value.subtypes[0]);
const subType = computed(() => sub.value.type);
const source = first ? null : chartSource(ui.value);
const orientation = ref<0 | 1>(0);

// The chosen category stays in view (the list is a sideways strip on phones).
const cats = ref<HTMLDivElement | null>(null);
const showCategory = () => {
  const strip = cats.value;
  const el = strip?.querySelector<HTMLElement>('[aria-selected="true"]');
  if (el && strip && strip.scrollWidth > strip.clientWidth) strip.scrollLeft = Math.max(0, el.offsetLeft - strip.offsetLeft - 8);
};
onMounted(showCategory);
watch(cat, showCategory, { flush: 'post' });

const variants = computed(() => {
  if (existing.value || !source) return null;
  const d = doc.value;
  const a = chartFromRange(d, d.sheet, source, subType.value);
  if (ONE_WAY_CHARTS.has(subType.value)) return [a];
  return [a, forcedChart(d, d.sheet, source, subType.value, a.seriesIn === 'rows' ? 'columns' : 'rows')];
});
const seriesCount = (first ?? variants.value?.[0])?.series.length ?? 0;
const custom = shallowRef<{ kind: SeriesKind; secondary?: boolean }[]>(
  first?.type === 'combo' ? first.series.map((s, i) => ({ kind: s.kind ?? (i === first.series.length - 1 ? 'line' : 'column'), secondary: s.secondaryAxis })) : comboKinds('column-line', Math.max(1, seriesCount)),
);
const kinds = computed(() => (sub.value.combo === 'custom' ? custom.value : undefined));
const chosen = computed(() => existing.value ?? variants.value?.[Math.min(orientation.value, (variants.value?.length ?? 1) - 1)] ?? null);
const candidate = computed(() => (chosen.value ? applySubtype(chosen.value, sub.value, kinds.value) : null));
const previews = computed(() => {
  const { doc } = ui.value;
  return variants.value?.map((v) => chartPreviewSvg(doc, applySubtype(v, sub.value, kinds.value), 220, 150)) ?? null;
});

const pickCategory = (c: ChartCategory) => {
  cat.value = c.id;
  subId.value = c.subtypes[0].id;
  orientation.value = 0;
};
const setKind = (i: number, kind: string) => {
  const next = custom.value.slice();
  next[i] = { ...(next[i] ?? {}), kind: kind as SeriesKind };
  custom.value = next;
};
const setSecondary = (i: number, on: boolean) => {
  const next = custom.value.slice();
  next[i] = { kind: next[i]?.kind ?? 'column', secondary: on };
  custom.value = next;
};
const ok = () => {
  const ctx = ui.value;
  const c = candidate.value;
  if (!c) {
    ctx.notify('Select the data for the chart first.');
    return false;
  }
  const ex = existing.value;
  if (ex) {
    ctx.doc.updateChart(ex.id, { type: c.type, series: c.series, markers: c.markers, scatterStyle: c.scatterStyle, stockKind: c.stockKind });
    return true;
  }
  const id = insertChart(ctx, sub.value, kinds.value);
  if (id && orientation.value === 1 && variants.value?.[1]) ctx.doc.updateChart(id, { seriesIn: c.seriesIn, headers: c.headers, categories: c.categories, series: c.series });
  return id !== null;
};
</script>

<template>
  <Dialog :title="existing ? 'Change Chart Type' : 'Insert Chart'" :on-ok="ok" :width="760" test-id="insert-chart" body-class="sh-chart-dialog" @close="emit('close')">
    <div class="sh-chart-dialog-grid">
      <div ref="cats" class="sh-chart-cats" role="listbox" aria-label="Chart categories" data-testid="chart-categories">
        <button v-for="c in CHART_CATEGORIES" :key="c.id" type="button" role="option" :aria-selected="c.id === cat" :class="c.id === cat ? 'selected' : ''" :data-testid="`chart-cat-${c.id}`" @click="pickCategory(c)">
          <Icon :name="c.icon" />
          <span>{{ c.label }}</span>
        </button>
      </div>
      <div class="sh-chart-main">
        <div class="sh-chart-subtypes" role="listbox" aria-label="Chart subtypes">
          <button v-for="s in catDef.subtypes" :key="s.id" type="button" role="option" :aria-selected="s.id === sub.id" :title="s.label" :class="s.id === sub.id ? 'selected' : ''" :data-testid="`chart-sub-${s.id}`" @click="subId = s.id">
            <Icon :name="s.icon" :size="28" />
          </button>
        </div>
        <div class="sh-chart-subtitle">{{ sub.label }}</div>
        <div v-if="sub.combo === 'custom' && chosen" class="sh-combo-table">
          <div class="sh-combo-head"><span>Series Name</span><span>Chart Type</span><span>Secondary Axis</span></div>
          <div v-for="(s, i) in chosen.series" :key="i" class="sh-combo-row">
            <span>{{ chartSeriesName(ui.doc, s, i) }}</span>
            <select :value="custom[i]?.kind ?? 'column'" @change="setKind(i, ($event.target as HTMLSelectElement).value)">
              <option v-for="k in COMBO_SERIES_KINDS" :key="k.kind" :value="k.kind">{{ k.label }}</option>
            </select>
            <input v-checked="!!custom[i]?.secondary" type="checkbox" aria-label="Secondary Axis" @change="setSecondary(i, ($event.target as HTMLInputElement).checked)" />
          </div>
        </div>
        <div v-if="previews" class="sh-chart-previews">
          <button
            v-for="(svg, i) in previews"
            :key="i"
            type="button"
            :class="`sh-chart-preview${orientation === i ? ' selected' : ''}`"
            :data-testid="`chart-preview-${i}`"
            :title="variants?.[i]?.seriesIn === 'rows' ? 'Series in rows' : 'Series in columns'"
            @click="orientation = i as 0 | 1"
            v-html="svg"
          />
        </div>
        <div v-else-if="candidate" class="sh-chart-preview single" v-html="chartPreviewSvg(ui.doc, candidate, 420, 240)" />
        <p v-else class="sh-hint">Select the data you want to chart first: a range with labels in the first row or column works best.</p>
      </div>
    </div>
  </Dialog>
</template>
