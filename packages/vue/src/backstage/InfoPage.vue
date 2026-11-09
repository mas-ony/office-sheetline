<script setup lang="ts">
/** File > Info: the document properties, dates, file and protection. */
import { DOC_PROPERTIES } from '@sheetline/core/views';
import { useUI } from '../context';
import { vDefaultValue } from '../directives';

const ui = useUI();
const date = (iso?: string) => (iso ? new Date(iso).toLocaleString() : '');
const blurred = (key: string, e: FocusEvent) => {
  const { doc } = ui.value;
  const v = (e.target as HTMLInputElement).value.trim();
  if ((doc.wb.props[key] ?? '') === v) return;
  if (v) doc.wb.props[key] = v;
  else delete doc.wb.props[key];
  doc.modified = true;
};
</script>

<template>
  <h1>Info</h1>
  <div class="sh-info">
    <div class="sh-info-props">
      <h2>Properties</h2>
      <label v-for="p in DOC_PROPERTIES" :key="p.key" class="sh-info-row">
        <span>{{ p.label }}</span>
        <input v-default-value="ui.doc.wb.props[p.key] ?? ''" class="sh-input" :placeholder="p.placeholder" :data-testid="`info-${p.key}`" @blur="blurred(p.key, $event)" />
      </label>
    </div>
    <div class="sh-info-side">
      <h2>Related Dates</h2>
      <div class="sh-info-row"><span>Last Modified</span><span>{{ date(ui.doc.wb.props.modified) }}</span></div>
      <div class="sh-info-row"><span>Created</span><span>{{ date(ui.doc.wb.props.created) }}</span></div>
      <h2>Related People</h2>
      <div class="sh-info-row"><span>Last Modified By</span><span>{{ ui.doc.wb.props.lastModifiedBy ?? '' }}</span></div>
      <h2>File</h2>
      <div class="sh-info-row"><span>Name</span><span data-testid="file-name">{{ ui.doc.fileName }}</span></div>
      <div class="sh-info-row"><span>Sheets</span><span>{{ ui.doc.wb.sheets.length }}</span></div>
      <h2>Protection</h2>
      <button type="button" class="sh-btn-secondary" data-testid="info-protect-workbook" @click="ui.openDialog({ kind: 'protect-workbook' })">{{ ui.doc.wb.protection?.structure ? 'Unprotect Workbook...' : 'Protect Workbook...' }}</button>
    </div>
  </div>
</template>
