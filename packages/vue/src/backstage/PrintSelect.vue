<script setup lang="ts">
/** A print setting: a native select drawn as Excel's setting button (icon, choice and hint). */
import type { IconName } from '@sheetline/core/views';
import { computed } from 'vue';
import Icon from '../Icon.vue';
import type { PrintOption } from './pages';

const props = defineProps<{ label: string; icon: IconName; value: string; options: PrintOption[]; testId: string }>();
const emit = defineEmits<{ change: [value: string] }>();
const sel = computed(() => props.options.find((o) => o.value === props.value) ?? props.options[0]);
</script>

<template>
  <label class="sh-print-select" :title="label">
    <span class="sr-only">{{ label }}</span>
    <span class="face" aria-hidden="true">
      <span class="ic"><Icon :name="icon" /></span>
      <span class="txt">
        <span class="t">{{ sel.label }}</span>
        <span v-if="sel.hint" class="h">{{ sel.hint }}</span>
      </span>
    </span>
    <select :aria-label="label" :data-testid="testId" :value="value" @change="emit('change', ($event.target as HTMLSelectElement).value)">
      <option v-for="o in options" :key="o.value" :value="o.value">{{ o.label }}</option>
    </select>
  </label>
</template>
