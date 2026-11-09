<script setup lang="ts">
/** Convert Text to Columns Wizard: data type, delimiters or break positions, then column formats and destination. */
import { commit, makeRange, parseCellRef, rangeText, splitText, type TextToColumnsOptions } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import { vValue } from '../../directives';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import Radio from '../Radio.vue';
import RangeField from '../RangeField.vue';

const emit = defineEmits<{ close: [] }>();

type ColumnFormat = 'general' | 'text' | 'skip';

const ui = useUI();
const range = computed(() => {
  const { doc } = ui.value;
  return doc.bounded(doc.range, doc.sheet) ?? doc.range;
});
const step = ref(1);
const opts = shallowRef<TextToColumnsOptions>({ type: 'delimited', delimiters: { tab: true, semicolon: false, comma: false, space: false, other: '' }, consecutive: false, qualifier: '"', breaks: [10] });
const dest = ref(rangeText(makeRange(range.value.r0, range.value.c0), ui.value.doc.sheet));
const formats = shallowRef<ColumnFormat[]>([]);
const col = ref(0);
const sample = computed(() => {
  const { doc } = ui.value;
  const r = range.value;
  const rows: string[] = [];
  for (let i = r.r0; i <= Math.min(r.r1, r.r0 + 7); i++) rows.push(doc.displayText(i, r.c0));
  return rows;
});
const split = computed(() => sample.value.map((t) => (t === '' ? [] : splitText(t, opts.value))));
const columns = computed(() => Array.from({ length: Math.max(1, ...split.value.map((s) => s.length)) }, (_, i) => i));

const set = (p: Partial<TextToColumnsOptions>) => (opts.value = { ...opts.value, ...p });
const setDelimiters = (p: Partial<TextToColumnsOptions['delimiters']>) => set({ delimiters: { ...opts.value.delimiters, ...p } });
const setBreaks = (text: string) => set({ breaks: text.split(/[,\s]+/).map(Number).filter((n) => n > 0) });
const setFormat = (f: ColumnFormat) => {
  const next = formats.value.slice();
  next[col.value] = f;
  formats.value = next;
};
const formatLabel = (i: number) => (formats.value[i] === 'skip' ? 'Skip' : formats.value[i] === 'text' ? 'Text' : 'General');

const finish = async () => {
  const ctx = ui.value;
  const { doc } = ctx;
  const r = range.value;
  const at = parseCellRef(dest.value.replace(/\$/g, '').replace(/^.*!/, ''));
  const full: TextToColumnsOptions = { ...opts.value, destination: at ? { r: at.r, c: at.c } : undefined, formats: formats.value };
  if (doc.textToColumnsOverwrites(r, full) && !(await ctx.confirm('Do you want to replace the contents of the destination cells?', { title: 'Text to Columns', ok: 'OK' }))) return false;
  commit(ctx);
  const error = doc.textToColumns(r, full);
  if (error) {
    ctx.notify(error);
    return false;
  }
  return true;
};
const ok = async () => {
  if (step.value < 3) {
    step.value++;
    return false;
  }
  return finish();
};
const finishNow = () => void finish().then((done) => done && emit('close'));
</script>

<template>
  <Dialog v-if="range.c0 !== range.c1" title="Convert Text to Columns Wizard" close-only :width="420" test-id="text-to-columns" @close="emit('close')">
    <p>This can convert only one column at a time. The range can be many rows tall but no more than one column wide. Try again by selecting cells in one column only.</p>
  </Dialog>
  <Dialog v-else :title="`Convert Text to Columns Wizard - Step ${step} of 3`" :width="600" test-id="text-to-columns" :ok="step < 3 ? 'Next >' : 'Finish'" :on-ok="ok" @close="emit('close')">
    <fieldset v-if="step === 1" class="sh-group">
      <legend>Original data type</legend>
      <Radio name="t2c-type" :checked="opts.type === 'delimited'" test-id="t2c-delimited" @change="set({ type: 'delimited' })"><span><b>Delimited</b> - Characters such as commas or tabs separate each field.</span></Radio>
      <Radio name="t2c-type" :checked="opts.type === 'fixed'" test-id="t2c-fixed" @change="set({ type: 'fixed' })"><span><b>Fixed width</b> - Fields are aligned in columns with spaces between each field.</span></Radio>
    </fieldset>
    <div v-if="step === 2 && opts.type === 'delimited'" class="sh-row wrap top">
      <fieldset class="sh-group">
        <legend>Delimiters</legend>
        <Check label="Tab" :checked="opts.delimiters.tab" @change="setDelimiters({ tab: $event })" />
        <Check label="Semicolon" :checked="opts.delimiters.semicolon" @change="setDelimiters({ semicolon: $event })" />
        <Check label="Comma" :checked="opts.delimiters.comma" test-id="t2c-comma" @change="setDelimiters({ comma: $event })" />
        <Check label="Space" :checked="opts.delimiters.space" test-id="t2c-space" @change="setDelimiters({ space: $event })" />
        <div class="sh-row">
          <Check label="Other:" :checked="!!opts.delimiters.other" @change="setDelimiters({ other: $event ? opts.delimiters.other || '|' : '' })" />
          <input :value="opts.delimiters.other" maxlength="1" :style="{ width: '32px' }" aria-label="Other delimiter" @input="setDelimiters({ other: ($event.target as HTMLInputElement).value })" />
        </div>
      </fieldset>
      <div class="sh-stack">
        <Check label="Treat consecutive delimiters as one" :checked="opts.consecutive" @change="set({ consecutive: $event })" />
        <Field label="Text qualifier:">
          <select :value="opts.qualifier" @change="set({ qualifier: ($event.target as HTMLSelectElement).value as TextToColumnsOptions['qualifier'] })">
            <option value='"'>"</option>
            <option value="'">'</option>
            <option value="">{none}</option>
          </select>
        </Field>
      </div>
    </div>
    <Field v-if="step === 2 && opts.type === 'fixed'" label="Break positions (characters, separated by commas):" wide>
      <input v-value="opts.breaks.join(', ')" @input="setBreaks(($event.target as HTMLInputElement).value)" />
    </Field>
    <div v-if="step === 3" class="sh-row wrap top">
      <fieldset class="sh-group">
        <legend>Column data format</legend>
        <Radio name="t2c-format" label="General" :checked="(formats[col] ?? 'general') === 'general'" @change="setFormat('general')" />
        <Radio name="t2c-format" label="Text" :checked="formats[col] === 'text'" @change="setFormat('text')" />
        <Radio name="t2c-format" label="Do not import column (skip)" :checked="formats[col] === 'skip'" @change="setFormat('skip')" />
      </fieldset>
      <Field label="Destination:">
        <RangeField :value="dest" label="Destination" test-id="t2c-destination" @change="dest = $event" />
      </Field>
    </div>
    <div class="sh-label">Data preview</div>
    <div class="sh-t2c-preview" data-testid="t2c-preview">
      <table>
        <thead v-if="step === 3">
          <tr>
            <th v-for="i in columns" :key="i" :class="i === col ? 'selected' : ''" @click="col = i">{{ formatLabel(i) }}</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(cells, r) in split" :key="r">
            <td v-if="step === 1">{{ sample[r] }}</td>
            <template v-else>
              <td v-for="i in columns" :key="i" :class="step === 3 && i === col ? 'selected' : ''" @click="col = i">{{ cells[i] ?? '' }}</td>
            </template>
          </tr>
        </tbody>
      </table>
    </div>
    <template #footer-left>
      <button type="button" class="sh-btn" :disabled="step === 1" @click="step--">&lt; Back</button>
      <button v-if="step < 3" type="button" class="sh-btn" data-testid="text-to-columns-finish" @click="finishNow">Finish</button>
    </template>
  </Dialog>
</template>
