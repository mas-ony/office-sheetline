<script setup lang="ts">
/** Page Setup: Page, Margins, Header/Footer (with custom sections) and Sheet. */
import { colName, HEADER_PRESETS, type HeaderFooterSet, hfSections, hfText, MARGIN_PRESETS, type PageSetup, PAPER, parseRangeText } from '@sheetline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../../context';
import { vValue } from '../../directives';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import Radio from '../Radio.vue';
import RangeField from '../RangeField.vue';
import Tabs from '../Tabs.vue';
import HeaderFooterEditor from './HeaderFooterEditor.vue';

type PageTab = 'page' | 'margins' | 'header' | 'sheet';
type Margin = keyof PageSetup['margins'];

const props = withDefaults(defineProps<{ tab?: PageTab }>(), { tab: 'page' });
const emit = defineEmits<{ close: [] }>();

const TABS: { id: PageTab; label: string }[] = [
  { id: 'page', label: 'Page' },
  { id: 'margins', label: 'Margins' },
  { id: 'header', label: 'Header/Footer' },
  { id: 'sheet', label: 'Sheet' },
];
const MARGINS: Margin[] = ['top', 'header', 'left', 'right', 'bottom', 'footer'];
const QUALITIES = [150, 300, 600, 1200];

const ui = useUI();
const { doc } = ui.value;
const sheet = doc.sheet;
const tab = ref<PageTab>(props.tab);
const p = shallowRef<PageSetup>({ ...sheet.page, margins: { ...sheet.page.margins }, headerFooter: { ...sheet.page.headerFooter } });
const titles = doc.printTitles();
const area = doc.printArea();
const printArea = ref(area ? `$${colName(area.c0)}$${area.r0 + 1}:$${colName(area.c1)}$${area.r1 + 1}` : '');
const rowsTop = ref(titles.rows ? `$${titles.rows[0] + 1}:$${titles.rows[1] + 1}` : '');
const colsLeft = ref(titles.cols ? `$${colName(titles.cols[0])}:$${colName(titles.cols[1])}` : '');
const hfEdit = ref<'header' | 'footer' | null>(null);
const val = (e: Event) => (e.target as HTMLInputElement).value;

const set = (patch: Partial<PageSetup>) => (p.value = { ...p.value, ...patch });
const setM = (k: Margin, v: string) => {
  const n = Number(v);
  if (Number.isFinite(n) && n >= 0) p.value = { ...p.value, margins: { ...p.value.margins, [k]: n } };
};
const setHF = (patch: Partial<HeaderFooterSet>) => (p.value = { ...p.value, headerFooter: { ...p.value.headerFooter, ...patch } });
const preview = (code: string | undefined) => {
  const { doc } = ui.value;
  const s = hfSections(code);
  const ctx = { page: 1, pages: 1, file: doc.fileName || 'Book1', sheet: doc.sheet.name, now: new Date() };
  return [hfText(s.left, ctx), hfText(s.center, ctx), hfText(s.right, ctx)];
};
const apply = (): boolean => {
  const { doc } = ui.value;
  let areaRange = null;
  if (printArea.value.trim()) {
    const hit = parseRangeText(doc, printArea.value);
    if (!hit) {
      ui.value.notify('The print area reference is not valid.');
      return false;
    }
    areaRange = hit.range;
  }
  const top = rowsTop.value.trim();
  const left = colsLeft.value.trim();
  const rows = top ? /^\$?(\d+):\$?(\d+)$/.exec(top.replace(/^.*!/, '')) : null;
  const cols = left ? /^\$?([A-Za-z]+):\$?([A-Za-z]+)$/.exec(left.replace(/^.*!/, '')) : null;
  if ((top && !rows) || (left && !cols)) {
    ui.value.notify('Print titles must be whole rows ($1:$2) or whole columns ($A:$B).');
    return false;
  }
  const colIndex = (s: string) => [...s.toUpperCase()].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  doc.transact('Page Setup', () => {
    doc.setPageSetup(p.value);
    const before = doc.printArea();
    if (JSON.stringify(before) !== JSON.stringify(areaRange)) doc.setPrintArea(areaRange);
    doc.setPrintTitles(rows ? [Number(rows[1]) - 1, Number(rows[2]) - 1] : null, cols ? [colIndex(cols[1]), colIndex(cols[2])] : null);
  });
  return true;
};
const print = () => {
  if (!apply()) return;
  emit('close');
  ui.value.openBackstage('print');
};
const presets = (code: string | undefined) =>
  HEADER_PRESETS.some((h) => h.code === (code ?? '')) ? HEADER_PRESETS : [...HEADER_PRESETS, { label: preview(code).filter(Boolean).join('  ') || '(custom)', code: code ?? '' }];
const headerPresets = computed(() => presets(p.value.headerFooter.oddHeader));
const footerPresets = computed(() => presets(p.value.headerFooter.oddFooter));
const saveHF = (code: string) => setHF(hfEdit.value === 'header' ? { oddHeader: code || undefined } : { oddFooter: code || undefined });
</script>

<template>
  <Dialog title="Page Setup" :width="560" test-id="page-setup" :on-ok="apply" @close="emit('close')">
    <template #footer-left>
      <button type="button" class="sh-btn" @click="print">Print...</button>
      <button type="button" class="sh-btn" data-testid="page-setup-preview" @click="print">Print Preview</button>
    </template>
    <Tabs :value="tab" :tabs="TABS" test-id="page-setup-tabs" @change="tab = $event" />
    <div class="sh-tab-body" :style="{ minHeight: '280px' }">
      <div v-if="tab === 'page'" class="sh-stack">
        <fieldset class="sh-group">
          <legend>Orientation</legend>
          <div class="sh-row">
            <Radio name="ps-orient" label="Portrait" :checked="p.orientation === 'portrait'" test-id="ps-portrait" @change="set({ orientation: 'portrait' })" />
            <Radio name="ps-orient" label="Landscape" :checked="p.orientation === 'landscape'" test-id="ps-landscape" @change="set({ orientation: 'landscape' })" />
          </div>
        </fieldset>
        <fieldset class="sh-group">
          <legend>Scaling</legend>
          <div class="sh-row">
            <Radio name="ps-scale" label="Adjust to:" :checked="!p.fitToPage" @change="set({ fitToPage: false })" />
            <input v-value="p.scale" inputmode="numeric" :style="{ width: '60px' }" :disabled="p.fitToPage" aria-label="Scale" data-testid="ps-scale" @input="set({ scale: Math.max(10, Math.min(400, Number(val($event)) || 100)) })" />
            <span>% normal size</span>
          </div>
          <div class="sh-row">
            <Radio name="ps-scale" label="Fit to:" :checked="p.fitToPage" test-id="ps-fit" @change="set({ fitToPage: true })" />
            <input v-value="p.fitToWidth || ''" inputmode="numeric" :style="{ width: '50px' }" :disabled="!p.fitToPage" aria-label="Pages wide" @input="set({ fitToWidth: Number(val($event)) || 0 })" />
            <span>page(s) wide by</span>
            <input v-value="p.fitToHeight || ''" inputmode="numeric" :style="{ width: '50px' }" :disabled="!p.fitToPage" aria-label="Pages tall" @input="set({ fitToHeight: Number(val($event)) || 0 })" />
            <span>tall</span>
          </div>
        </fieldset>
        <Field label="Paper size:">
          <select :value="p.paperSize" data-testid="ps-paper" @change="set({ paperSize: Number(val($event)) })">
            <option v-for="x in PAPER" :key="x.code" :value="x.code">{{ x.label }} {{ x.size }}</option>
          </select>
        </Field>
        <Field label="Print quality:">
          <select :value="p.printQuality ?? 600" @change="set({ printQuality: Number(val($event)) })">
            <option v-for="q in QUALITIES" :key="q" :value="q">{{ q }} dpi</option>
          </select>
        </Field>
        <Field label="First page number:">
          <input v-value="p.firstPageNumber ?? 'Auto'" :style="{ width: '80px' }" @input="set({ firstPageNumber: Number(val($event)) || undefined })" />
        </Field>
      </div>
      <div v-if="tab === 'margins'" class="sh-stack">
        <div class="sh-margins-grid">
          <Field v-for="k in MARGINS" :key="k" :label="`${k[0].toUpperCase()}${k.slice(1)}:`">
            <input v-value="p.margins[k]" inputmode="decimal" :style="{ width: '70px' }" :data-testid="`ps-margin-${k}`" @input="setM(k, val($event))" />
          </Field>
        </div>
        <div class="sh-row wrap">
          <button v-for="m in MARGIN_PRESETS" :key="m.name" type="button" class="sh-btn" @click="set({ margins: { ...m.margins } })">{{ m.label }}</button>
        </div>
        <fieldset class="sh-group">
          <legend>Center on page</legend>
          <Check label="Horizontally" :checked="p.centerH" @change="set({ centerH: $event })" />
          <Check label="Vertically" :checked="p.centerV" @change="set({ centerV: $event })" />
        </fieldset>
        <p class="sh-hint">Margins are in inches.</p>
      </div>
      <div v-if="tab === 'header'" class="sh-stack">
        <div class="sh-hf-preview">
          <span v-for="(t, i) in preview(p.headerFooter.oddHeader)" :key="i">{{ t }}</span>
        </div>
        <Field label="Header:" wide>
          <select :value="p.headerFooter.oddHeader ?? ''" data-testid="ps-header" @change="setHF({ oddHeader: val($event) || undefined })">
            <option v-for="(h, i) in headerPresets" :key="i" :value="h.code">{{ h.label }}</option>
          </select>
        </Field>
        <div class="sh-row">
          <button type="button" class="sh-btn" data-testid="ps-custom-header" @click="hfEdit = 'header'">Custom Header...</button>
          <button type="button" class="sh-btn" @click="hfEdit = 'footer'">Custom Footer...</button>
        </div>
        <Field label="Footer:" wide>
          <select :value="p.headerFooter.oddFooter ?? ''" data-testid="ps-footer" @change="setHF({ oddFooter: val($event) || undefined })">
            <option v-for="(h, i) in footerPresets" :key="i" :value="h.code">{{ h.label }}</option>
          </select>
        </Field>
        <div class="sh-hf-preview">
          <span v-for="(t, i) in preview(p.headerFooter.oddFooter)" :key="i">{{ t }}</span>
        </div>
        <Check label="Different odd and even pages" :checked="!!p.headerFooter.differentOddEven" @change="setHF({ differentOddEven: $event || undefined })" />
        <Check label="Different first page" :checked="!!p.headerFooter.differentFirst" @change="setHF({ differentFirst: $event || undefined })" />
      </div>
      <div v-if="tab === 'sheet'" class="sh-stack">
        <Field label="Print area:" wide>
          <RangeField :value="printArea" label="Print area" test-id="ps-print-area" @change="printArea = $event" />
        </Field>
        <fieldset class="sh-group">
          <legend>Print titles</legend>
          <Field label="Rows to repeat at top:" wide>
            <RangeField :value="rowsTop" label="Rows to repeat at top" test-id="ps-rows-top" @change="rowsTop = $event.replace(/^\$?[A-Z]+(\$?\d+):\$?[A-Z]+(\$?\d+)$/i, '$1:$2')" />
          </Field>
          <Field label="Columns to repeat at left:" wide>
            <RangeField :value="colsLeft" label="Columns to repeat at left" test-id="ps-cols-left" @change="colsLeft = $event.replace(/^(\$?[A-Z]+)\$?\d+:(\$?[A-Z]+)\$?\d+$/i, '$1:$2')" />
          </Field>
        </fieldset>
        <fieldset class="sh-group">
          <legend>Print</legend>
          <Check label="Gridlines" :checked="p.gridLines" test-id="ps-gridlines" @change="set({ gridLines: $event })" />
          <Check label="Black and white" :checked="p.blackAndWhite" @change="set({ blackAndWhite: $event })" />
          <Check label="Row and column headings" :checked="p.headings" @change="set({ headings: $event })" />
        </fieldset>
        <fieldset class="sh-group">
          <legend>Page order</legend>
          <Radio name="ps-order" label="Down, then over" :checked="p.pageOrder === 'downThenOver'" @change="set({ pageOrder: 'downThenOver' })" />
          <Radio name="ps-order" label="Over, then down" :checked="p.pageOrder === 'overThenDown'" @change="set({ pageOrder: 'overThenDown' })" />
        </fieldset>
      </div>
    </div>
  </Dialog>
  <HeaderFooterEditor
    v-if="hfEdit"
    :title="hfEdit === 'header' ? 'Header' : 'Footer'"
    :code="(hfEdit === 'header' ? p.headerFooter.oddHeader : p.headerFooter.oddFooter) ?? ''"
    @close="hfEdit = null"
    @save="saveHF"
  />
</template>
