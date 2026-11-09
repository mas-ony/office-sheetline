<script setup lang="ts">
/** File > Print: what to print, orientation, paper, margins and scaling, with the pages previewed. */
import {
  buildPages,
  canvasMeasure,
  fileBaseName,
  loadPrintFonts,
  MARGIN_PRESETS,
  type Measure,
  type Page,
  pagesToPdf,
  pageSvg,
  PAPER,
  paperOf,
  type PrintChoice,
  printOptionsFor,
  printPages,
  printScalingKind,
  rasterize,
  saveFile,
} from '@sheetline/core/views';
import { computed, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue';
import { useUI } from '../context';
import { vNumberValue, vValue } from '../directives';
import Icon from '../Icon.vue';
import type { PrintOption } from './pages';
import PrintSelect from './PrintSelect.vue';

const emit = defineEmits<{ close: [] }>();

const PT_TO_PX = 96 / 72;

const ui = useUI();
const choice = shallowRef<PrintChoice>({ scope: 'sheet' });
const pages = shallowRef<Page[] | null>(null);
const current = ref(0);
const busy = ref(false);
const fit = ref(true);
let measure: Measure | null = null;
const pane = ref<HTMLDivElement | null>(null);
const paneSize = shallowRef({ w: 600, h: 700 });

// A new object on every change: the page setup is updated in place.
const view = computed(() => {
  const { doc, state } = ui.value;
  return { doc, sheet: doc.sheet, page: doc.sheet.page, version: state.version, title: fileBaseName(doc.fileName) };
});
let optsMemo: { key: unknown[]; value: ReturnType<typeof printOptionsFor> } | null = null;
const opts = computed(() => {
  const { doc, sheet, version } = view.value;
  const key = [doc, sheet, choice.value, version];
  if (optsMemo && optsMemo.key.every((k, i) => k === key[i])) return optsMemo.value;
  optsMemo = { key, value: printOptionsFor(doc, sheet, choice.value) };
  return optsMemo.value;
});
watch(
  opts,
  (o, _old, onCleanup) => {
    let alive = true;
    const { doc } = view.value;
    const t = setTimeout(async () => {
      await loadPrintFonts();
      if (!alive) return;
      measure ??= canvasMeasure();
      const built = buildPages(doc, o, measure, { title: doc.fileName });
      pages.value = built;
      current.value = Math.min(current.value, built.length - 1);
    }, 30);
    onCleanup(() => {
      alive = false;
      clearTimeout(t);
    });
  },
  { immediate: true },
);

let observer: ResizeObserver | null = null;
onMounted(() => {
  const el = pane.value;
  if (!el) return;
  observer = new ResizeObserver(() => (paneSize.value = { w: el.clientWidth, h: el.clientHeight }));
  observer.observe(el);
  paneSize.value = { w: el.clientWidth, h: el.clientHeight };
});
onUnmounted(() => observer?.disconnect());

const pg = computed(() => pages.value?.[current.value]);
const scale = computed(() => {
  const p = pg.value;
  if (!p) return 1;
  return fit.value ? Math.min((paneSize.value.w - 40) / (p.width * PT_TO_PX), (paneSize.value.h - 40) / (p.height * PT_TO_PX)) : 1;
});
// Generated from the document by the print layout (escaped text).
const markup = computed(() => (pg.value ? pageSvg(pg.value, { width: '100%', height: '100%' }) : ''));

const download = async () => {
  if (!pages.value || busy.value) return;
  busy.value = true;
  try {
    const title = view.value.title;
    const bytes = await pagesToPdf(pages.value, { rasterize, measure: measure ?? canvasMeasure() }, { title });
    saveFile(bytes, `${title}.pdf`, 'application/pdf');
  } catch (e) {
    ui.value.notify(`The PDF could not be created: ${(e as Error).message}`);
  } finally {
    busy.value = false;
  }
};

const scaling = computed(() => {
  const { page } = view.value;
  return printScalingKind(page.fitToPage, page.fitToWidth, page.fitToHeight);
});
const marginPreset = computed(() => MARGIN_PRESETS.find((m) => JSON.stringify(m.margins) === JSON.stringify(view.value.page.margins))?.name ?? 'custom');

const SCOPES: PrintOption[] = [
  { value: 'sheet', label: 'Print Active Sheets', hint: 'Only print the active sheets' },
  { value: 'workbook', label: 'Print Entire Workbook', hint: 'Print the entire workbook' },
  { value: 'selection', label: 'Print Selection', hint: 'Only print the current selection' },
];
const ORIENTATIONS: PrintOption[] = [
  { value: 'portrait', label: 'Portrait Orientation' },
  { value: 'landscape', label: 'Landscape Orientation' },
];
const PAPERS: PrintOption[] = PAPER.map((p) => ({ value: String(p.code), label: p.label, hint: p.size }));
const margins = computed<PrintOption[]>(() => {
  const { page } = view.value;
  return [
    ...MARGIN_PRESETS.map((m) => ({ value: m.name, label: `${m.label} Margins`, hint: `Top: ${m.margins.top}" Bottom: ${m.margins.bottom}" Left: ${m.margins.left}" Right: ${m.margins.right}"` })),
    { value: 'custom', label: 'Custom Margins...', hint: marginPreset.value === 'custom' ? `Top: ${page.margins.top}" Left: ${page.margins.left}"` : '' },
  ];
});
const scalings = computed<PrintOption[]>(() => {
  const { page } = view.value;
  return [
    { value: 'none', label: page.fitToPage || page.scale === 100 ? 'No Scaling' : `Custom Scaling (${page.scale}%)`, hint: 'Print sheets at their actual size' },
    { value: 'sheet', label: 'Fit Sheet on One Page', hint: 'Shrink the printout so that it fits on one page' },
    { value: 'columns', label: 'Fit All Columns on One Page', hint: 'Shrink the printout so that it is one page wide' },
    { value: 'rows', label: 'Fit All Rows on One Page', hint: 'Shrink the printout so that it is one page high' },
    { value: 'custom', label: 'Custom Scaling Options...' },
  ];
});

const setMargins = (v: string) => {
  if (v === 'custom') ui.value.openDialog({ kind: 'page-setup', tab: 'margins' });
  else view.value.doc.setPageSetup({ margins: { ...MARGIN_PRESETS.find((m) => m.name === v)!.margins } });
};
const setScaling = (v: string) => {
  const { doc } = view.value;
  if (v === 'none') doc.setPageSetup({ fitToPage: false, scale: 100 });
  else if (v === 'sheet') doc.setPageSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 1 });
  else if (v === 'columns') doc.setPageSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 0 });
  else if (v === 'rows') doc.setPageSetup({ fitToPage: true, fitToWidth: 0, fitToHeight: 1 });
  else ui.value.openDialog({ kind: 'page-setup', tab: 'page' });
};
const pageNumber = (e: Event) => (e.target as HTMLInputElement).value;
const setFrom = (e: Event) => (choice.value = { ...choice.value, pageFrom: pageNumber(e) ? Number(pageNumber(e)) : undefined });
const setTo = (e: Event) => (choice.value = { ...choice.value, pageTo: pageNumber(e) ? Number(pageNumber(e)) : undefined });
const goTo = (e: Event) => {
  const n = Number(pageNumber(e));
  if (pages.value && n >= 1 && n <= pages.value.length) current.value = n - 1;
};
</script>

<template>
  <div class="sh-print">
    <div class="sh-print-settings" data-testid="print-options">
      <h1>Print</h1>
      <div class="sh-print-actions">
        <button type="button" class="sh-bs-action big" data-testid="print-print" :disabled="!pages" @click="pages && printPages(pages)">
          <Icon name="print" />
          <span>Print</span>
        </button>
        <button type="button" class="sh-bs-action big" data-testid="print-download" :disabled="!pages || busy" @click="download">
          <Icon :name="busy ? 'hourglass_empty' : 'picture_as_pdf'" />
          <span>{{ busy ? 'Creating…' : 'Save as PDF' }}</span>
        </button>
      </div>
      <h2>Settings</h2>
      <PrintSelect label="What to print" icon="grid_on" test-id="print-scope" :value="choice.scope" :options="SCOPES" @change="choice = { ...choice, scope: $event as PrintChoice['scope'] }" />
      <div class="sh-print-pages">
        <span>Pages:</span>
        <input v-number-value="choice.pageFrom ?? ''" type="number" :min="1" class="sh-input" aria-label="Pages from" data-testid="print-from" @input="setFrom" />
        <span>to</span>
        <input v-number-value="choice.pageTo ?? ''" type="number" :min="1" class="sh-input" aria-label="Pages to" data-testid="print-to" @input="setTo" />
      </div>
      <PrintSelect
        label="Orientation"
        :icon="view.page.orientation === 'landscape' ? 'crop_landscape' : 'crop_portrait'"
        test-id="print-orientation"
        :value="view.page.orientation"
        :options="ORIENTATIONS"
        @change="view.doc.setPageSetup({ orientation: $event as 'portrait' | 'landscape' })"
      />
      <PrintSelect label="Paper size" icon="description" test-id="print-paper" :value="String(view.page.paperSize)" :options="PAPERS" @change="view.doc.setPageSetup({ paperSize: Number($event) })" />
      <PrintSelect label="Margins" icon="margin" test-id="print-margins" :value="marginPreset" :options="margins" @change="setMargins" />
      <PrintSelect label="Scaling" icon="fit_screen" test-id="print-scaling" :value="scaling" :options="scalings" @change="setScaling" />
      <button type="button" class="sh-link" data-testid="print-page-setup" @click="ui.openDialog({ kind: 'page-setup', tab: 'page' })">Page Setup</button>
      <div class="sh-hint" :style="{ marginTop: '12px' }">{{ paperOf(view.page.paperSize).label }} · {{ pages?.length ?? 0 }} page{{ pages?.length === 1 ? '' : 's' }}</div>
      <button type="button" class="sh-link" @click="emit('close')">Back to the sheet</button>
    </div>
    <div class="sh-print-preview">
      <div ref="pane" class="sh-print-canvas" data-testid="print-preview">
        <div v-if="!pages" class="sh-print-wait">Preparing pages…</div>
        <div
          v-if="pg"
          class="sh-print-page"
          data-testid="print-page"
          :style="{ width: `${pg.width * PT_TO_PX * scale}px`, height: `${pg.height * PT_TO_PX * scale}px` }"
          v-html="markup"
        />
      </div>
      <div class="sh-print-nav">
        <button type="button" aria-label="Previous page" data-testid="print-prev" :disabled="current <= 0" @click="current--"><Icon name="arrow_left" /></button>
        <input v-value="pages ? current + 1 : 0" class="sh-input" aria-label="Current page" @input="goTo" />
        <span data-testid="print-page-count">of {{ pages?.length ?? 0 }}</span>
        <button type="button" aria-label="Next page" data-testid="print-next" :disabled="!pages || current >= pages.length - 1" @click="current++"><Icon name="arrow_right" /></button>
        <span class="grow" />
        <button type="button" :aria-label="fit ? 'Actual size' : 'Zoom to page'" :title="fit ? 'Actual size' : 'Zoom to page'" :class="fit ? 'on' : ''" @click="fit = !fit"><Icon name="fit_screen" /></button>
      </div>
    </div>
  </div>
</template>
