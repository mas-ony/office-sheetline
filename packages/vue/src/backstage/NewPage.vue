<script setup lang="ts">
/** File > New: a blank workbook (after asking to save the modified one). */
import { useUI } from '../context';

const emit = defineEmits<{ close: [] }>();
const ui = useUI();
const create = async () => {
  const ctx = ui.value;
  if (ctx.doc.modified) {
    const answer = await ctx.ask('Warning: Current file has been modified. Do you want to save your changes?', { title: 'Spreadsheet', ok: 'Yes', alt: 'No', cancel: 'Cancel' });
    if (answer === 'cancel') return;
    if (answer === 'ok') ctx.files.save();
  }
  emit('close');
  ctx.files.newWorkbook();
};
</script>

<template>
  <h1>New</h1>
  <div class="sh-bs-tiles">
    <button type="button" class="sh-bs-tile" data-testid="new-blank" aria-label="Blank workbook" @click="create">
      <span class="sh-blank-preview" aria-hidden="true">
        <svg viewBox="0 0 200 150" width="200" height="150">
          <rect x="0.5" y="0.5" width="199" height="149" fill="#fff" stroke="#9e9e9e" />
          <rect x="0.5" y="0.5" width="199" height="18" fill="#f3f3f3" stroke="#d4d4d4" />
          <rect x="0.5" y="0.5" width="24" height="149" fill="#f3f3f3" stroke="#d4d4d4" />
          <line v-for="i in [1, 2, 3, 4, 5, 6]" :key="`h${i}`" x1="24" :y1="18 + i * 19" x2="200" :y2="18 + i * 19" stroke="#e0e0e0" />
          <line v-for="i in [1, 2]" :key="`v${i}`" :x1="24 + i * 62" y1="18" :x2="24 + i * 62" y2="150" stroke="#e0e0e0" />
          <rect x="25" y="19" width="61" height="18" fill="none" stroke="#217346" stroke-width="2" />
          <text v-for="(l, i) in ['A', 'B', 'C']" :key="l" :x="55 + i * 62" y="13" font-size="9" text-anchor="middle" fill="#555">{{ l }}</text>
          <text v-for="i in [1, 2, 3, 4, 5, 6, 7]" :key="i" x="12" :y="32 + (i - 1) * 19" font-size="9" text-anchor="middle" fill="#555">{{ i }}</text>
        </svg>
      </span>
      <span>Blank workbook</span>
    </button>
  </div>
</template>
