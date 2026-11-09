<script setup lang="ts">
/** Find and Replace, with Go To (names and references). */
import { cellName, cmd, type SearchHit } from '@sheetline/core/views';
import { computed, onUnmounted, ref, shallowRef, watch } from 'vue';
import { useUI } from '../../context';
import { vFocus } from '../../directives';
import Check from '../Check.vue';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import ListBox from '../ListBox.vue';
import Tabs from '../Tabs.vue';

type FindTab = 'find' | 'replace' | 'goto';

const props = withDefaults(defineProps<{ tab?: FindTab }>(), { tab: 'find' });
const emit = defineEmits<{ close: [] }>();

const TABS: { id: FindTab; label: string }[] = [
  { id: 'find', label: 'Find' },
  { id: 'replace', label: 'Replace' },
  { id: 'goto', label: 'Go To' },
];

const ui = useUI();
const tab = ref<FindTab>(props.tab);
const query = ref('');
const replacement = ref('');
const options = ref(false);
const within = ref<'sheet' | 'workbook'>('sheet');
const byCols = ref(false);
const lookIn = ref<'formulas' | 'values'>('formulas');
const matchCase = ref(false);
const whole = ref(false);
const hits = shallowRef<SearchHit[] | null>(null);
const index = ref(-1);
const all = shallowRef<SearchHit[] | null>(null);
const goto = ref('');
const opts = computed(() => ({ matchCase: matchCase.value, wholeCell: whole.value, inFormulas: lookIn.value === 'formulas', allSheets: within.value === 'workbook' }));
const val = (e: Event) => (e.target as HTMLInputElement).value;

watch([query, opts, byCols], () => {
  hits.value = null;
  index.value = -1;
});

const clearHit = (grid: typeof ui.value.grid) => {
  if (grid) {
    grid.searchHit = undefined;
    grid.requestRender();
  }
};
watch(() => ui.value.grid, (_grid, old) => clearHit(old));
onUnmounted(() => clearHit(ui.value.grid));

const search = () => {
  const { doc } = ui.value;
  let list = doc.find(query.value, opts.value);
  if (byCols.value) list = [...list].sort((a, b) => doc.wb.sheets.indexOf(a.sheet) - doc.wb.sheets.indexOf(b.sheet) || a.c - b.c || a.r - b.r);
  hits.value = list;
  return list;
};

const show = (hit: SearchHit) => {
  const { doc, grid } = ui.value;
  const si = doc.wb.sheets.indexOf(hit.sheet);
  if (si !== doc.activeIndex) doc.setActiveSheet(si);
  doc.selectCell(hit.r, hit.c);
  grid?.scrollIntoView(hit.r, hit.c);
  if (grid) {
    grid.searchHit = { r: hit.r, c: hit.c };
    grid.requestRender();
  }
};

const next = () => {
  if (!query.value) return;
  const { doc } = ui.value;
  const list = hits.value ?? search();
  if (!list.length) {
    ui.value.notify(`We couldn't find what you were looking for.`);
    return;
  }
  // Start after the active cell like Excel.
  let k = index.value + 1;
  if (index.value < 0) {
    const { r, c } = doc.selection.active;
    const pos = list.findIndex((h) => h.sheet === doc.sheet && (byCols.value ? h.c > c || (h.c === c && h.r > r) : h.r > r || (h.r === r && h.c > c)));
    k = pos < 0 ? 0 : pos;
  }
  k %= list.length;
  index.value = k;
  show(list[k]);
};

const replaceOne = () => {
  const { doc } = ui.value;
  const list = hits.value ?? search();
  const at = index.value;
  const cur = list[at];
  const { r, c } = doc.selection.active;
  if (cur && cur.sheet === doc.sheet && cur.r === r && cur.c === c) {
    doc.replace(cur, query.value, replacement.value, opts.value);
    const fresh = search();
    index.value = at - 1;
    if (fresh.length) {
      const k = Math.min(Math.max(0, at), fresh.length - 1);
      index.value = k;
      show(fresh[k]);
    }
    return;
  }
  next();
};

const replaceAll = () => {
  const ctx = ui.value;
  cmd.commit(ctx);
  const n = ctx.doc.replaceAll(query.value, replacement.value, opts.value);
  hits.value = null;
  ctx.notify(n ? `All done. We made ${n} replacement${n > 1 ? 's' : ''}.` : `We couldn't find anything to replace.`);
};

const names = computed(() => ui.value.doc.wb.names.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.')).map((n) => n.name));
const same = (x: string) => x;
const optionsLabel = computed(() => (options.value ? 'Options <<' : 'Options >>'));

const findKey = (e: KeyboardEvent) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    next();
  }
};
const activate = (x: string) => {
  if (ui.value.doc.goTo(x)) emit('close');
};
const gotoKey = (e: KeyboardEvent) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    if (ui.value.doc.goTo(goto.value)) emit('close');
    else ui.value.notify('Reference isn\'t valid.');
  }
};
const gotoOk = () => {
  const { doc, grid } = ui.value;
  if (doc.goTo(goto.value)) {
    grid?.scrollIntoView(doc.selection.active.r, doc.selection.active.c);
    emit('close');
  } else ui.value.notify('Reference isn\'t valid.');
};
</script>

<template>
  <Dialog :title="tab === 'goto' ? 'Go To' : 'Find and Replace'" close-only :blocking="false" :width="520" test-id="find-dialog" @close="emit('close')">
    <Tabs :value="tab" :tabs="TABS" test-id="find-tabs" @change="tab = $event" />
    <div class="sh-tab-body">
      <div v-if="tab !== 'goto'" class="sh-stack">
        <Field label="Find what:" wide>
          <input v-focus :value="query" data-testid="find-input" @input="query = val($event)" @keydown="findKey" />
        </Field>
        <Field v-if="tab === 'replace'" label="Replace with:" wide>
          <input :value="replacement" data-testid="replace-input" @input="replacement = val($event)" />
        </Field>
        <div v-if="options" class="sh-row wrap top">
          <div class="sh-stack">
            <Field label="Within:">
              <select :value="within" @change="within = val($event) as typeof within">
                <option value="sheet">Sheet</option>
                <option value="workbook">Workbook</option>
              </select>
            </Field>
            <Field label="Search:">
              <select :value="byCols ? 'cols' : 'rows'" @change="byCols = val($event) === 'cols'">
                <option value="rows">By Rows</option>
                <option value="cols">By Columns</option>
              </select>
            </Field>
            <Field label="Look in:">
              <select :value="lookIn" @change="lookIn = val($event) as typeof lookIn">
                <option value="formulas">Formulas</option>
                <option value="values">Values</option>
              </select>
            </Field>
          </div>
          <div class="sh-stack">
            <Check label="Match case" :checked="matchCase" test-id="find-match-case" @change="matchCase = $event" />
            <Check label="Match entire cell contents" :checked="whole" test-id="find-whole" @change="whole = $event" />
          </div>
        </div>
        <div class="sh-row wrap">
          <button type="button" class="sh-btn" @click="options = !options">{{ optionsLabel }}</button>
          <span class="grow" />
          <template v-if="tab === 'replace'">
            <button type="button" class="sh-btn" data-testid="replace-all" :disabled="!query" @click="replaceAll">Replace All</button>
            <button type="button" class="sh-btn" data-testid="replace-one" :disabled="!query" @click="replaceOne">Replace</button>
          </template>
          <button type="button" class="sh-btn" data-testid="find-all" :disabled="!query" @click="all = search()">Find All</button>
          <button type="button" class="sh-btn primary" data-testid="find-next" :disabled="!query" @click="next">Find Next</button>
        </div>
        <div v-if="all" class="sh-table finds" role="table" data-testid="find-results">
          <div class="sh-table-head" role="row"><span>Sheet</span><span>Cell</span><span>Value</span></div>
          <div class="sh-table-body">
            <div v-for="(h, i) in all" :key="i" role="row" class="sh-table-row" @click="show(h)">
              <span>{{ h.sheet.name }}</span>
              <span>{{ cellName(h.r, h.c, true, true) }}</span>
              <span>{{ h.text }}</span>
            </div>
          </div>
          <div class="sh-hint">{{ all.length }} cell(s) found</div>
        </div>
      </div>
      <div v-else class="sh-stack">
        <div class="sh-label">Go to:</div>
        <ListBox v-slot="{ item }" :items="names" :value="names.includes(goto) ? goto : null" :key-of="same" :height="140" label="Names" :on-activate="activate" @change="goto = $event">{{ item }}</ListBox>
        <Field label="Reference:" wide>
          <input :value="goto" data-testid="goto-input" @input="goto = val($event)" @keydown="gotoKey" />
        </Field>
        <div class="sh-row">
          <button type="button" class="sh-btn" @click="ui.openDialog({ kind: 'goto-special' })">Special...</button>
          <span class="grow" />
          <button type="button" class="sh-btn primary" data-testid="goto-ok" @click="gotoOk">OK</button>
        </div>
      </div>
    </div>
  </Dialog>
</template>
