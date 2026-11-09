<script setup lang="ts">
/** The name box: the selection's reference, a field to go to a cell or name, and the list of defined names. */
import { computed, ref, shallowRef, watch } from 'vue';
import { useUI } from './context';
import Dropdown from './Dropdown.vue';
import Icon from './Icon.vue';

const ui = useUI();
const name = ref(ui.value.state.rangeRef);
const focused = ref(false);
const list = shallowRef<DOMRect | null>(null);
const wrap = ref<HTMLDivElement | null>(null);
watch([() => ui.value.state.rangeRef, focused], ([rangeRef, f]) => {
  if (!f) name.value = rangeRef;
});
const names = computed(() => ui.value.doc.wb.names.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.')));

const goTo = () => {
  const { doc, grid } = ui.value;
  const text = name.value.trim();
  if (text && !doc.goTo(text)) ui.value.notify(`“${text}” is not a valid cell reference or name.`);
  focused.value = false;
  grid?.focusIdle();
};
const focus = (e: FocusEvent) => {
  focused.value = true;
  (e.target as HTMLInputElement).select();
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    goTo();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    name.value = ui.value.state.rangeRef;
    ui.value.grid?.focusIdle();
  }
};
const toggle = () => (list.value = list.value ? null : wrap.value!.getBoundingClientRect());
const pick = (n: string) => {
  list.value = null;
  ui.value.doc.goTo(n);
  ui.value.grid?.focusIdle();
};
const manage = () => {
  list.value = null;
  ui.value.openDialog({ kind: 'name-manager' });
};
</script>

<template>
  <div ref="wrap" class="sh-namebox">
    <input
      aria-label="Name box"
      data-testid="name-box"
      :value="name"
      spellcheck="false"
      autocapitalize="characters"
      autocomplete="off"
      enterkeyhint="go"
      @focus="focus"
      @blur="focused = false"
      @input="name = ($event.target as HTMLInputElement).value"
      @keydown="key"
    />
    <button type="button" class="sh-namebox-caret" aria-label="Defined names" data-testid="names-button" :aria-expanded="!!list" @mousedown.prevent @click="toggle">
      <Icon name="arrow_drop_down" />
    </button>
    <Dropdown v-if="list" :anchor="list" test-id="names-list" :sheet="ui.phone" title="Names" @close="list = null">
      <div class="sh-names" role="listbox" aria-label="Defined names">
        <div v-if="names.length === 0" class="sh-hint" :style="{ padding: '6px 12px' }">No defined names</div>
        <button v-for="n in names" :key="`${n.name}-${n.localSheet ?? ''}`" type="button" class="sh-combo-item" @click="pick(n.name)">{{ n.name }}</button>
        <button v-if="!ui.state.readOnly" type="button" class="sh-combo-item more" data-testid="names-manage" @click="manage">Name Manager...</button>
      </div>
    </Dropdown>
  </div>
</template>
