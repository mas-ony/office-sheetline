<script setup lang="ts">
/** A ribbon control with fields of its own. */
import type { Ctx, ViewSpec } from '@sheetline/core/views';
import ObjectSize from './ObjectSize.vue';
import TableName from './TableName.vue';

defineProps<{ spec: ViewSpec; ctx: Ctx }>();
</script>

<template>
  <TableName v-if="spec.kind === 'table-name'" :ctx="ctx" />
  <ObjectSize v-else-if="spec.kind === 'object-size'" :ctx="ctx" />
  <div v-else-if="spec.kind === 'sheet-options'" class="sh-sheet-options">
    <div v-for="col in spec.columns" :key="col.head" class="sh-sheet-options-col">
      <span class="sh-sheet-options-head">{{ col.head }}</span>
      <label v-for="c in col.items" :key="c.id" class="sh-check">
        <input type="checkbox" :data-testid="`rb-${c.id}`" :checked="!!c.checked?.(ctx)" @change="c.run?.(ctx)" />
        <span>{{ c.label }}</span>
      </label>
    </div>
  </div>
  <span v-else-if="spec.kind === 'separator'" class="sh-tsep" />
</template>
