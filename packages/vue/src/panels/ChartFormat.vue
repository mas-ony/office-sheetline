<script setup lang="ts">
/** The Format pane of a chart: chart area, title, legend, axes and series. */
import { type ColorSpec, type LineDash, readChartData, type SeriesKind, type SheetChart, type Side, type Trendline } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import { useUI } from '../context';
import Check from '../dialogs/Check.vue';
import ColorButton from '../dialogs/ColorButton.vue';
import Field from '../dialogs/Field.vue';
import { vChecked } from '../directives';
import { DASHES, hexOf, specOf } from './format';
import LazyInput from './LazyInput.vue';
import LineEditor from './LineEditor.vue';
import PaneSection from './PaneSection.vue';
import SidePanel from './SidePanel.vue';
import SizeSection from './SizeSection.vue';

type ChartPart = 'area' | 'title' | 'legend' | 'haxis' | 'vaxis' | `series:${number}`;

const props = defineProps<{ chart: SheetChart }>();

const ui = useUI();
const part = ref<ChartPart>('area');
const TRENDLINES = ['none', 'linear', 'exponential', 'logarithmic', 'polynomial', 'power', 'moving-average'];
// A new object on every change, as the data follows the cells.
const view = computed(() => {
  const chart = props.chart;
  const data = readChartData(ui.value.doc, chart);
  const pie = ['pie', 'doughnut', 'treemap', 'sunburst', 'funnel'].includes(chart.type);
  const seriesIndex = part.value.startsWith('series:') ? Number(part.value.slice(7)) : -1;
  const series = seriesIndex >= 0 ? chart.series[seriesIndex] : undefined;
  const lineLike = chart.type.endsWith('line') || chart.type === 'radar' || chart.type === 'scatter' || (chart.type === 'combo' && series?.kind === 'line');
  return { chart, data, pie, seriesIndex, series, lineLike };
});
const update = (patch: Partial<SheetChart>) => ui.value.doc.updateChart(props.chart.id, patch);
const setSeries = (patch: Partial<SheetChart['series'][number]>) => {
  const { chart, seriesIndex } = view.value;
  update({ series: chart.series.map((s, i) => (i === seriesIndex ? { ...s, ...patch } : s)) });
};
const selectValue = (e: Event) => (e.target as HTMLSelectElement).value;
const hex = (c: ColorSpec | null) => hexOf(ui.value, c);
const numberOrUndefined = (v: string) => (v.trim() === '' || !Number.isFinite(Number(v)) ? undefined : Number(v));
const trendLabel = (t: string) => (t === 'none' ? 'None' : t[0].toUpperCase() + t.slice(1).replace('-', ' '));
const setLineWidth = (v: string) => {
  if (Number(v) > 0) setSeries({ lineWidth: Math.min(20, Number(v)) });
};
</script>

<template>
  <SidePanel title="Format Chart" test-id="format-pane">
    <Field label="Chart Options:" wide>
      <select :value="part" data-testid="format-chart-part" @change="part = selectValue($event) as ChartPart">
        <option value="area">Chart Area</option>
        <option value="title">Chart Title</option>
        <option value="legend">Legend</option>
        <option v-if="!view.pie" value="haxis">Horizontal (Category) Axis</option>
        <option v-if="!view.pie" value="vaxis">Vertical (Value) Axis</option>
        <option v-for="(s, i) in chart.series" :key="i" :value="`series:${i}`">Series "{{ view.data.series[i]?.name ?? s.name ?? i + 1 }}"</option>
      </select>
    </Field>
    <template v-if="part === 'area'">
      <PaneSection title="Fill">
        <div class="sh-row">
          <label class="sh-check"><input v-checked="chart.fill === null" type="radio" @change="update({ fill: null })" /> <span>No fill</span></label>
          <label class="sh-check"><input v-checked="chart.fill !== null" type="radio" @change="update({ fill: chart.fill ?? undefined })" /> <span>Solid fill</span></label>
        </div>
        <Field v-if="chart.fill !== null" label="Color:">
          <ColorButton :value="specOf(chart.fill ?? '#FFFFFF')" test-id="chart-fill" none-label="Automatic" @change="update({ fill: hex($event) ?? undefined })" />
        </Field>
      </PaneSection>
      <PaneSection title="Border">
        <LineEditor :line="chart.border === undefined ? { color: '#D9D9D9', width: 1 } : chart.border" @change="update({ border: $event })" />
      </PaneSection>
      <PaneSection title="Data Labels and Gridlines">
        <Check label="Data labels" :checked="chart.valueLabels" @change="update({ valueLabels: $event })" />
        <Check v-if="!view.pie" label="Major horizontal gridlines" :checked="chart.gridlines" @change="update({ gridlines: $event })" />
        <Check v-if="!view.pie" label="Major vertical gridlines" :checked="!!chart.vGridlines" @change="update({ vGridlines: $event || undefined })" />
      </PaneSection>
      <SizeSection :id="chart.id" />
      <PaneSection title="Properties" :open="false">
        <Field label="Name:" wide>
          <LazyInput :value="chart.name" label="Chart name" @commit="$event.trim() && update({ name: $event.trim() })" />
        </Field>
      </PaneSection>
    </template>
    <PaneSection v-if="part === 'title'" title="Title">
      <Check label="Show chart title" :checked="!!chart.title" @change="update({ title: $event ? chart.title || 'Chart Title' : undefined })" />
      <Field label="Text:" wide>
        <LazyInput :value="chart.title ?? ''" label="Chart title" test-id="chart-title-input" @commit="update({ title: $event || undefined })" />
      </Field>
      <Field label="Subtitle:" wide>
        <LazyInput :value="chart.subtitle ?? ''" label="Subtitle" @commit="update({ subtitle: $event || undefined })" />
      </Field>
      <Field label="Position:">
        <select :value="chart.titlePosition ?? 'top'" @change="update({ titlePosition: selectValue($event) as Side })">
          <option value="top">Top</option>
          <option value="bottom">Bottom</option>
          <option value="left">Left</option>
          <option value="right">Right</option>
        </select>
      </Field>
    </PaneSection>
    <PaneSection v-if="part === 'legend'" title="Legend">
      <Check label="Show legend" :checked="chart.showLegend" test-id="chart-legend-show" @change="update({ showLegend: $event })" />
      <Field label="Position:">
        <select :value="chart.legendPosition ?? 'bottom'" @change="update({ legendPosition: selectValue($event) as Side, showLegend: true })">
          <option value="bottom">Bottom</option>
          <option value="top">Top</option>
          <option value="left">Left</option>
          <option value="right">Right</option>
        </select>
      </Field>
      <Field label="Title:" wide>
        <LazyInput :value="chart.legendTitle ?? ''" label="Legend title" @commit="update({ legendTitle: $event || undefined })" />
      </Field>
    </PaneSection>
    <PaneSection v-if="part === 'haxis'" title="Horizontal Axis">
      <Check label="Show axis" :checked="!chart.hideXAxis" @change="update({ hideXAxis: !$event || undefined })" />
      <Field label="Axis title:" wide>
        <LazyInput :value="chart.xAxisTitle ?? ''" label="Horizontal axis title" @commit="update({ xAxisTitle: $event || undefined })" />
      </Field>
    </PaneSection>
    <PaneSection v-if="part === 'vaxis'" title="Vertical Axis">
      <Check label="Show axis" :checked="!chart.hideYAxis" @change="update({ hideYAxis: !$event || undefined })" />
      <Field label="Axis title:" wide>
        <LazyInput :value="chart.yAxisTitle ?? ''" label="Vertical axis title" @commit="update({ yAxisTitle: $event || undefined })" />
      </Field>
      <Field label="Minimum:">
        <LazyInput :value="chart.yMin === undefined ? '' : String(chart.yMin)" label="Minimum" input-mode="decimal" test-id="chart-ymin" @commit="update({ yMin: numberOrUndefined($event) })" />
      </Field>
      <Field label="Maximum:">
        <LazyInput :value="chart.yMax === undefined ? '' : String(chart.yMax)" label="Maximum" input-mode="decimal" @commit="update({ yMax: numberOrUndefined($event) })" />
      </Field>
      <Field label="Secondary axis title:" wide>
        <LazyInput :value="chart.secondaryAxisTitle ?? ''" label="Secondary axis title" @commit="update({ secondaryAxisTitle: $event || undefined })" />
      </Field>
    </PaneSection>
    <template v-if="view.series">
      <PaneSection title="Series Options">
        <Field v-if="chart.type === 'combo'" label="Chart type:">
          <select :value="view.series.kind ?? 'column'" @change="setSeries({ kind: selectValue($event) as SeriesKind })">
            <option value="column">Clustered Column</option>
            <option value="stacked-column">Stacked Column</option>
            <option value="line">Line</option>
            <option value="area">Area</option>
            <option value="stacked-area">Stacked Area</option>
          </select>
        </Field>
        <Check v-if="!view.pie" label="Plot on secondary axis" :checked="!!view.series.secondaryAxis" @change="setSeries({ secondaryAxis: $event || undefined })" />
        <Field v-if="!view.pie" label="Trendline:">
          <select :value="view.series.trendline ?? 'none'" @change="setSeries({ trendline: selectValue($event) === 'none' ? undefined : (selectValue($event) as Trendline) })">
            <option v-for="t in TRENDLINES" :key="t" :value="t">{{ trendLabel(t) }}</option>
          </select>
        </Field>
      </PaneSection>
      <PaneSection :title="view.lineLike ? 'Line' : 'Fill'">
        <Field label="Color:">
          <ColorButton :value="specOf(view.series.color)" none-label="Automatic" test-id="series-color" @change="setSeries({ color: hex($event) ?? undefined })" />
        </Field>
        <template v-if="view.lineLike">
          <Field label="Width (px):">
            <LazyInput :value="String(view.series.lineWidth ?? 2.5)" label="Line width" input-mode="decimal" @commit="setLineWidth" />
          </Field>
          <Field label="Dash type:">
            <select :value="view.series.lineStyle ?? 'solid'" @change="setSeries({ lineStyle: selectValue($event) === 'solid' ? undefined : (selectValue($event) as LineDash) })">
              <option v-for="d in DASHES" :key="d.dash" :value="d.dash">{{ d.label }}</option>
            </select>
          </Field>
          <Check label="Markers" :checked="view.series.markers ?? !!chart.markers" @change="setSeries({ markers: $event })" />
          <Check label="Smoothed line" :checked="!!view.series.smooth" @change="setSeries({ smooth: $event || undefined })" />
        </template>
      </PaneSection>
    </template>
  </SidePanel>
</template>
