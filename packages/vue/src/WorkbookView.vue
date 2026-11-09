<script setup lang="ts">
/** What the editor shows for one workbook (drawn afresh for each workbook opened). */
import { RIBBON_TABS, type UIState } from '@sheetline/core/views';
import Backstage from './backstage/Backstage.vue';
import CommentsBox from './CommentsBox.vue';
import { useUI } from './context';
import ConfirmDialog from './dialogs/ConfirmDialog.vue';
import DialogHost from './dialogs/DialogHost.vue';
import PromptDialog from './dialogs/PromptDialog.vue';
import { vChecked } from './directives';
import FilterPopup from './FilterPopup.vue';
import FormulaAssist from './FormulaAssist.vue';
import FormulaBar from './FormulaBar.vue';
import GridHost from './GridHost.vue';
import Icon from './Icon.vue';
import KeyBar from './KeyBar.vue';
import Menu from './Menu.vue';
import NoteEditor from './NoteEditor.vue';
import SidePanels from './panels/SidePanels.vue';
import Ribbon from './ribbon/Ribbon.vue';
import SheetBar from './SheetBar.vue';
import StatusBar from './StatusBar.vue';

defineProps<{ state: UIState }>();

const { ui, layer } = useUI().value;
</script>

<template>
  <Ribbon
    :tabs="RIBBON_TABS"
    :active="ui.currentTab()"
    :layout="state.phone ? 'ribbon' : state.layout"
    :collapsed="state.collapsed && !state.phone"
    @tab="ui.showTab"
    @file="ui.openBackstage('new')"
    @collapse="ui.setCollapsed"
  >
    <template #extra>
      <button
        type="button"
        class="sh-theme-toggle"
        :title="state.theme === 'dark' ? 'Light theme' : 'Dark theme'"
        aria-label="Switch Theme"
        data-testid="theme-toggle"
        @click="ui.setTheme(state.theme === 'dark' ? 'light' : 'dark')"
      >
        <Icon :name="state.theme === 'dark' ? 'light_mode' : 'dark_mode'" />
      </button>
      <label v-if="!state.phone" class="sh-switch" title="Switch Layout (ribbon / toolbar)">
        <input v-checked="state.layout === 'toolbar'" type="checkbox" role="switch" aria-label="Switch Layout" data-testid="layout-toggle" @change="ui.setLayout(($event.target as HTMLInputElement).checked ? 'toolbar' : 'ribbon')" />
        <span />
      </label>
    </template>
  </Ribbon>
  <FormulaBar v-if="state.view.formulaBar" />
  <div class="sh-main">
    <GridHost :state="state" />
    <SidePanels v-if="state.panel" />
  </div>
  <SheetBar v-if="state.view.tabStrip" @tab-menu="ui.openTabMenu" />
  <StatusBar v-if="state.view.statusBar" />
  <Backstage v-if="state.backstage" :page="state.backstage" @page="ui.setBackstagePage" @close="ui.closeBackstage()" />
  <FormulaAssist />
  <Teleport v-if="state.phone && state.edit.active" :to="layer"><KeyBar /></Teleport>
  <Menu v-if="state.menu" :items="state.menu.items" :at="state.menu.at" :sheet="state.phone" :title="state.menu.title" :test-id="state.menu.testId" @close="ui.closeMenu" />
  <DialogHost v-if="state.dialog" :spec="state.dialog" @close="ui.closeDialog()" />
  <CommentsBox v-if="state.comments" :r="state.comments.r" :c="state.comments.c" :rect="state.comments.rect" :start-new="state.comments.startNew" @close="ui.closeComments" />
  <NoteEditor v-if="state.note" :r="state.note.r" :c="state.note.c" :rect="state.note.rect" @close="ui.closeNote" />
  <FilterPopup v-if="state.filter" :col="state.filter.col" :at="state.filter.at" :table-id="state.filter.tableId" @close="ui.closeFilter" />
  <ConfirmDialog
    v-if="state.pending?.kind === 'confirm'"
    :message="state.pending.message"
    :title="state.pending.opts?.title"
    :ok="state.pending.opts?.ok"
    :cancel="state.pending.opts?.cancel"
    :danger="state.pending.opts?.danger"
    @answer="(v) => ui.answer(v)"
  />
  <ConfirmDialog
    v-if="state.pending?.kind === 'ask'"
    :message="state.pending.message"
    :title="state.pending.opts?.title"
    :ok="state.pending.opts?.ok ?? 'Yes'"
    :alt="state.pending.opts?.alt ?? 'No'"
    :cancel="state.pending.opts?.cancel ?? 'Cancel'"
    @answer="(v, alt) => ui.answer(v ? 'ok' : alt ? 'alt' : 'cancel')"
  />
  <PromptDialog v-if="state.pending?.kind === 'prompt'" :message="state.pending.message" :value="state.pending.value" :title="state.pending.opts?.title" :ok="state.pending.opts?.ok" @answer="(v) => ui.answer(v)" />
  <Teleport v-if="state.toast" :to="layer">
    <div class="sh-toast" role="status" data-testid="toast">
      <span>{{ state.toast }}</span>
      <button type="button" @click="ui.dismissToast">OK</button>
    </div>
  </Teleport>
</template>
