<script setup lang="ts">
/**
 * The File backstage: New, Open, Save, Import, Export, Print (settings and live preview) and Info
 * (document properties), shown over the whole window.
 */
import { BACKSTAGE_PAGES as PAGES } from '@sheetline/core/views';
import { onMounted, onUnmounted } from 'vue';
import { useUI } from '../context';
import Icon from '../Icon.vue';
import ExportPage from './ExportPage.vue';
import ImportPage from './ImportPage.vue';
import InfoPage from './InfoPage.vue';
import NewPage from './NewPage.vue';
import OpenPage from './OpenPage.vue';
import PrintPage from './PrintPage.vue';
import SavePage from './SavePage.vue';

defineProps<{ page: string }>();
const emit = defineEmits<{ page: [page: string]; close: [] }>();

const ui = useUI();
const key = (e: KeyboardEvent) => {
  if (e.key === 'Escape' && !ui.value.layer.querySelector('.sh-modal-layer')) emit('close');
};
onMounted(() => window.addEventListener('keydown', key));
onUnmounted(() => window.removeEventListener('keydown', key));
const close = () => emit('close');
</script>

<template>
  <div class="sh-backstage" data-testid="backstage" role="dialog" aria-label="File">
    <nav class="sh-bs-nav" role="menu" aria-label="File">
      <button type="button" class="sh-bs-back" aria-label="Back" data-testid="backstage-back" @click="close">
        <Icon name="arrow_back" />
      </button>
      <div v-for="(p, i) in PAGES" :key="p.id">
        <div v-if="i > 0 && PAGES[i - 1].group !== p.group" class="sh-bs-sep" />
        <button
          type="button"
          role="menuitemradio"
          :aria-checked="page === p.id"
          :class="`sh-bs-item${page === p.id ? ' active' : ''}`"
          :data-testid="`backstage-${p.id}`"
          @click="emit('page', p.id)"
        >{{ p.label }}</button>
      </div>
    </nav>
    <section :class="`sh-bs-content${page === 'print' ? ' print' : ''}`" role="tabpanel" :aria-label="PAGES.find((p) => p.id === page)?.label">
      <NewPage v-if="page === 'new'" @close="close" />
      <OpenPage v-if="page === 'open'" @close="close" />
      <SavePage v-if="page === 'save'" @close="close" />
      <ImportPage v-if="page === 'import'" @close="close" />
      <ExportPage v-if="page === 'export'" @close="close" />
      <PrintPage v-if="page === 'print'" @close="close" />
      <InfoPage v-if="page === 'info'" />
    </section>
    <div v-if="ui.phone" class="sh-bs-phone-pad" />
  </div>
</template>
