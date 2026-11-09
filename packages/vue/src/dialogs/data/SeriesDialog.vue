<script setup lang="ts">
/** Fill Series: rows or columns, the series type and date unit, trend, step and stop values. */
import { commit, type SeriesOptions } from '@sheetline/core/views';
import { ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import Radio from '../Radio.vue';

const emit = defineEmits<{ close: [] }>();

const UNITS = ['day', 'weekday', 'month', 'year'] as const;

const ui = useUI();
const range0 = ui.value.doc.range;
const opts = shallowRef<SeriesOptions>({ by: range0.r1 - range0.r0 >= range0.c1 - range0.c0 ? 'columns' : 'rows', type: 'linear', unit: 'day', step: 1, trend: false });
const step = ref('1');
const stop = ref('');
const set = (p: Partial<SeriesOptions>) => (opts.value = { ...opts.value, ...p });
const unitLabel = (u: (typeof UNITS)[number]) => (u === 'weekday' ? 'Weekday' : u[0].toUpperCase() + u.slice(1));

const ok = () => {
  const ctx = ui.value;
  const s = Number(step.value);
  const st = stop.value.trim() === '' ? undefined : Number(stop.value);
  if (!Number.isFinite(s) || (st !== undefined && !Number.isFinite(st))) {
    ctx.notify('Enter numbers for the step and stop values.');
    return false;
  }
  commit(ctx);
  ctx.doc.fillSeriesRange(ctx.doc.range, { ...opts.value, step: s, stop: st });
  return true;
};
</script>

<template>
  <Dialog title="Series" :width="420" test-id="series-dialog" :on-ok="ok" @close="emit('close')">
    <div class="sh-row top wrap">
      <fieldset class="sh-group">
        <legend>Series in</legend>
        <Radio name="series-in" label="Rows" :checked="opts.by === 'rows'" @change="set({ by: 'rows' })" />
        <Radio name="series-in" label="Columns" :checked="opts.by === 'columns'" @change="set({ by: 'columns' })" />
      </fieldset>
      <fieldset class="sh-group">
        <legend>Type</legend>
        <Radio name="series-type" label="Linear" :checked="opts.type === 'linear'" @change="set({ type: 'linear' })" />
        <Radio name="series-type" label="Growth" :checked="opts.type === 'growth'" @change="set({ type: 'growth' })" />
        <Radio name="series-type" label="Date" :checked="opts.type === 'date'" @change="set({ type: 'date' })" />
        <Radio name="series-type" label="AutoFill" :checked="opts.type === 'autofill'" @change="set({ type: 'autofill' })" />
      </fieldset>
      <fieldset class="sh-group">
        <legend>Date unit</legend>
        <Radio v-for="u in UNITS" :key="u" name="series-unit" :label="unitLabel(u)" :checked="opts.unit === u" :disabled="opts.type !== 'date'" @change="set({ unit: u })" />
      </fieldset>
    </div>
    <Check label="Trend" :checked="opts.trend" :disabled="opts.type === 'date' || opts.type === 'autofill'" @change="set({ trend: $event })" />
    <div class="sh-row">
      <Field label="Step value:">
        <input :value="step" inputmode="decimal" :style="{ width: '80px' }" data-testid="series-step" @input="step = ($event.target as HTMLInputElement).value" />
      </Field>
      <Field label="Stop value:">
        <input :value="stop" inputmode="decimal" :style="{ width: '80px' }" data-testid="series-stop" @input="stop = ($event.target as HTMLInputElement).value" />
      </Field>
    </div>
  </Dialog>
</template>
