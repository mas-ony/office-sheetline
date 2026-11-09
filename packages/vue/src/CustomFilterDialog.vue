<script setup lang="ts">
/** Excel's Custom AutoFilter dialog: two conditions joined by And / Or. */
import { type ColumnFilter, FILTER_OPERATORS, type FilterCondition, type FilterOperator } from '@sheetline/core/views';
import { ref } from 'vue';
import { vChecked } from './directives';
import Modal from './Modal.vue';

const props = defineProps<{ title: string; initial: ColumnFilter | undefined }>();
const emit = defineEmits<{ apply: [filter: ColumnFilter]; close: [] }>();

const conditions = ref<[FilterCondition, FilterCondition]>([
  props.initial?.conditions?.[0] ?? { op: 'equal', value: '' },
  props.initial?.conditions?.[1] ?? { op: 'equal', value: '' },
]);
const and = ref(props.initial?.and ?? true);
const setOp = (i: 0 | 1, e: Event) => (conditions.value[i] = { ...conditions.value[i], op: (e.target as HTMLSelectElement).value as FilterOperator });
const setValue = (i: 0 | 1, e: Event) => (conditions.value[i] = { ...conditions.value[i], value: (e.target as HTMLInputElement).value });
const ok = () => emit('apply', { conditions: conditions.value.filter((x) => x.value !== ''), and: and.value });
</script>

<template>
  <Modal title="Custom AutoFilter" :width="460" blocking test-id="custom-filter-dialog" @close="emit('close')">
    <div class="sh-form">
      <div>Show rows where: <b>{{ title }}</b></div>
      <template v-for="n in [1, 2] as const" :key="n">
        <div v-if="n === 2" class="sh-form-row">
          <label><input v-checked="and" type="radio" @change="and = true" /> And</label>
          <label><input v-checked="!and" type="radio" @change="and = false" /> Or</label>
        </div>
        <div class="sh-form-row">
          <select class="sh-select" :value="conditions[n - 1].op" :aria-label="`Operator ${n}`" :data-testid="`custom-filter-op${n}`" @change="setOp((n - 1) as 0 | 1, $event)">
            <option v-for="o in FILTER_OPERATORS" :key="o.op" :value="o.op">{{ o.label }}</option>
          </select>
          <input class="sh-input" :value="conditions[n - 1].value" :aria-label="`Value ${n}`" :data-testid="`custom-filter-value${n}`" @input="setValue((n - 1) as 0 | 1, $event)" />
        </div>
      </template>
      <div class="sh-hint">Use ? to represent any single character. Use * to represent any series of characters.</div>
    </div>
    <template #footer>
      <button type="button" class="sh-btn-primary" data-testid="custom-filter-ok" @click="ok">OK</button>
      <button type="button" class="sh-btn-secondary" @click="emit('close')">Cancel</button>
    </template>
  </Modal>
</template>
