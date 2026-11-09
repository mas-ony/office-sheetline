<script setup lang="ts">
/** Shows the dialog described by a DialogSpec. */
import type { DialogSpec } from '@sheetline/core/views';
import CfManagerDialog from './cond-format/CfManagerDialog.vue';
import CfPresetDialog from './cond-format/CfPresetDialog.vue';
import CfRuleDialog from './cond-format/CfRuleDialog.vue';
import DataTableDialog from './data/DataTableDialog.vue';
import GoalSeekDialog from './data/GoalSeekDialog.vue';
import GroupDialog from './data/GroupDialog.vue';
import OutlineSettingsDialog from './data/OutlineSettingsDialog.vue';
import RemoveDuplicatesDialog from './data/RemoveDuplicatesDialog.vue';
import ScenariosDialog from './data/ScenariosDialog.vue';
import SeriesDialog from './data/SeriesDialog.vue';
import SortDialog from './data/SortDialog.vue';
import SubtotalDialog from './data/SubtotalDialog.vue';
import TextToColumnsDialog from './data/TextToColumnsDialog.vue';
import ValidationDialog from './data/ValidationDialog.vue';
import FindDialog from './edit/FindDialog.vue';
import GotoSpecialDialog from './edit/GotoSpecialDialog.vue';
import HyperlinkDialog from './edit/HyperlinkDialog.vue';
import InsertCellsDialog from './edit/InsertCellsDialog.vue';
import PasteSpecialDialog from './edit/PasteSpecialDialog.vue';
import SizeDialog from './edit/SizeDialog.vue';
import ZoomDialog from './edit/ZoomDialog.vue';
import FormatCellsDialog from './format/FormatCellsDialog.vue';
import InsertFunctionDialog from './formula/InsertFunctionDialog.vue';
import NameManagerDialog from './formula/NameManagerDialog.vue';
import NamesFromSelectionDialog from './formula/NamesFromSelectionDialog.vue';
import NewNameDialog from './formula/NewNameDialog.vue';
import AltTextDialog from './object/AltTextDialog.vue';
import CreateTableDialog from './object/CreateTableDialog.vue';
import InsertChartDialog from './object/InsertChartDialog.vue';
import MoveChartDialog from './object/MoveChartDialog.vue';
import ResizeTableDialog from './object/ResizeTableDialog.vue';
import SelectDataDialog from './object/SelectDataDialog.vue';
import ShapeTextDialog from './object/ShapeTextDialog.vue';
import SparklineDialog from './object/SparklineDialog.vue';
import SymbolDialog from './object/SymbolDialog.vue';
import CommandPalette from './settings/CommandPalette.vue';
import CustomizeColorsDialog from './settings/CustomizeColorsDialog.vue';
import NewCellStyleDialog from './settings/NewCellStyleDialog.vue';
import SheetSettingsDialog from './settings/SheetSettingsDialog.vue';
import ShortcutsDialog from './settings/ShortcutsDialog.vue';
import SpreadSettingsDialog from './settings/SpreadSettingsDialog.vue';
import MoveCopySheetDialog from './sheet/MoveCopySheetDialog.vue';
import PageSetupDialog from './sheet/PageSetupDialog.vue';
import ProtectSheetDialog from './sheet/ProtectSheetDialog.vue';
import ProtectWorkbookDialog from './sheet/ProtectWorkbookDialog.vue';
import UnhideSheetDialog from './sheet/UnhideSheetDialog.vue';
import UnprotectSheetDialog from './sheet/UnprotectSheetDialog.vue';

// Insert Cells and Delete Cells share one instance, as the React interface keeps it between them.
defineProps<{ spec: DialogSpec }>();
const emit = defineEmits<{ close: [] }>();
const close = () => emit('close');
</script>

<template>
  <FormatCellsDialog v-if="spec.kind === 'format-cells'" :tab="spec.tab" @close="close" />
  <CfPresetDialog v-else-if="spec.kind === 'cf-preset'" :preset="spec.preset" @close="close" />
  <CfRuleDialog v-else-if="spec.kind === 'cf-rule'" :hint="spec.ruleType" @close="close" />
  <CfManagerDialog v-else-if="spec.kind === 'cf-manager'" @close="close" />
  <CreateTableDialog v-else-if="spec.kind === 'create-table'" :table-style="spec.style" @close="close" />
  <ResizeTableDialog v-else-if="spec.kind === 'resize-table'" :table-id="spec.tableId" @close="close" />
  <InsertChartDialog v-else-if="spec.kind === 'insert-chart'" :category="spec.category" :change-id="spec.changeId" @close="close" />
  <SelectDataDialog v-else-if="spec.kind === 'select-data'" :chart-id="spec.chartId" @close="close" />
  <MoveChartDialog v-else-if="spec.kind === 'move-chart'" :chart-id="spec.chartId" @close="close" />
  <SparklineDialog v-else-if="spec.kind === 'sparkline'" :type="spec.type" :group-id="spec.groupId" :single="spec.single" @close="close" />
  <HyperlinkDialog v-else-if="spec.kind === 'hyperlink'" @close="close" />
  <InsertFunctionDialog v-else-if="spec.kind === 'insert-function'" :category="spec.category" @close="close" />
  <NameManagerDialog v-else-if="spec.kind === 'name-manager'" @close="close" />
  <NewNameDialog v-else-if="spec.kind === 'new-name'" :name="spec.name" :scope="spec.scope" @close="close" />
  <NamesFromSelectionDialog v-else-if="spec.kind === 'names-from-selection'" @close="close" />
  <SortDialog v-else-if="spec.kind === 'sort'" @close="close" />
  <TextToColumnsDialog v-else-if="spec.kind === 'text-to-columns'" @close="close" />
  <ValidationDialog v-else-if="spec.kind === 'validation'" @close="close" />
  <RemoveDuplicatesDialog v-else-if="spec.kind === 'remove-duplicates'" @close="close" />
  <GroupDialog v-else-if="spec.kind === 'group'" :ungroup="spec.ungroup" @close="close" />
  <SubtotalDialog v-else-if="spec.kind === 'subtotal'" @close="close" />
  <OutlineSettingsDialog v-else-if="spec.kind === 'outline-settings'" @close="close" />
  <GoalSeekDialog v-else-if="spec.kind === 'goal-seek'" @close="close" />
  <DataTableDialog v-else-if="spec.kind === 'data-table'" @close="close" />
  <ScenariosDialog v-else-if="spec.kind === 'scenarios'" @close="close" />
  <SeriesDialog v-else-if="spec.kind === 'series'" @close="close" />
  <FindDialog v-else-if="spec.kind === 'find'" :tab="spec.tab" @close="close" />
  <GotoSpecialDialog v-else-if="spec.kind === 'goto-special'" @close="close" />
  <ZoomDialog v-else-if="spec.kind === 'zoom'" @close="close" />
  <PageSetupDialog v-else-if="spec.kind === 'page-setup'" :tab="spec.tab" @close="close" />
  <InsertCellsDialog v-else-if="spec.kind === 'insert-cells' || spec.kind === 'delete-cells'" :remove="spec.kind === 'delete-cells'" @close="close" />
  <SizeDialog v-else-if="spec.kind === 'size'" :what="spec.what" @close="close" />
  <ProtectSheetDialog v-else-if="spec.kind === 'protect-sheet'" @close="close" />
  <UnprotectSheetDialog v-else-if="spec.kind === 'unprotect-sheet'" @close="close" />
  <ProtectWorkbookDialog v-else-if="spec.kind === 'protect-workbook'" @close="close" />
  <MoveCopySheetDialog v-else-if="spec.kind === 'move-copy-sheet'" :index="spec.index" @close="close" />
  <UnhideSheetDialog v-else-if="spec.kind === 'unhide-sheet'" @close="close" />
  <PasteSpecialDialog v-else-if="spec.kind === 'paste-special'" @close="close" />
  <SpreadSettingsDialog v-else-if="spec.kind === 'spread-settings'" :tab="spec.tab" @close="close" />
  <SheetSettingsDialog v-else-if="spec.kind === 'sheet-settings'" :tab="spec.tab" @close="close" />
  <CustomizeColorsDialog v-else-if="spec.kind === 'customize-colors'" @close="close" />
  <AltTextDialog v-else-if="spec.kind === 'alt-text'" :id="spec.id" @close="close" />
  <CommandPalette v-else-if="spec.kind === 'command-palette'" @close="close" />
  <NewCellStyleDialog v-else-if="spec.kind === 'new-cell-style'" @close="close" />
  <ShortcutsDialog v-else-if="spec.kind === 'shortcuts'" @close="close" />
  <SymbolDialog v-else-if="spec.kind === 'symbol'" @close="close" />
  <ShapeTextDialog v-else-if="spec.kind === 'shape-text'" :id="spec.id" @close="close" />
</template>
