/** Shows the dialog described by a DialogSpec. */
import type { DialogSpec } from '@sheetline/core/views';
import { FormatCellsDialog } from './FormatCells';
import { CfManagerDialog, CfPresetDialog, CfRuleDialog } from './CondFormatDialogs';
import { AltTextDialog, CreateTableDialog, InsertChartDialog, MoveChartDialog, ResizeTableDialog, SelectDataDialog, ShapeTextDialog, SparklineDialog, SymbolDialog } from './ObjectDialogs';
import { DataTableDialog, GoalSeekDialog, GroupDialog, OutlineSettingsDialog, RemoveDuplicatesDialog, ScenariosDialog, SeriesDialog, SortDialog, SubtotalDialog, TextToColumnsDialog, ValidationDialog } from './DataDialogs';
import { InsertFunctionDialog, NameManagerDialog, NamesFromSelectionDialog, NewNameDialog } from './FormulaDialogs';
import { FindDialog, GotoSpecialDialog, HyperlinkDialog, InsertCellsDialog, PasteSpecialDialog, SizeDialog, ZoomDialog } from './EditDialogs';
import { MoveCopySheetDialog, PageSetupDialog, ProtectSheetDialog, ProtectWorkbookDialog, UnhideSheetDialog, UnprotectSheetDialog } from './SheetDialogs';
import { CommandPalette, CustomizeColorsDialog, NewCellStyleDialog, SheetSettingsDialog, ShortcutsDialog, SpreadSettingsDialog } from './SettingsDialogs';

export function DialogHost({ spec, onClose }: { spec: DialogSpec; onClose: () => void }) {
  switch (spec.kind) {
    case 'format-cells':
      return <FormatCellsDialog tab={spec.tab} onClose={onClose} />;
    case 'cf-preset':
      return <CfPresetDialog preset={spec.preset} onClose={onClose} />;
    case 'cf-rule':
      return <CfRuleDialog hint={spec.ruleType} onClose={onClose} />;
    case 'cf-manager':
      return <CfManagerDialog onClose={onClose} />;
    case 'create-table':
      return <CreateTableDialog style={spec.style} onClose={onClose} />;
    case 'resize-table':
      return <ResizeTableDialog tableId={spec.tableId} onClose={onClose} />;
    case 'insert-chart':
      return <InsertChartDialog category={spec.category} changeId={spec.changeId} onClose={onClose} />;
    case 'select-data':
      return <SelectDataDialog chartId={spec.chartId} onClose={onClose} />;
    case 'move-chart':
      return <MoveChartDialog chartId={spec.chartId} onClose={onClose} />;
    case 'sparkline':
      return <SparklineDialog type={spec.type} groupId={spec.groupId} single={spec.single} onClose={onClose} />;
    case 'hyperlink':
      return <HyperlinkDialog onClose={onClose} />;
    case 'insert-function':
      return <InsertFunctionDialog category={spec.category} onClose={onClose} />;
    case 'name-manager':
      return <NameManagerDialog onClose={onClose} />;
    case 'new-name':
      return <NewNameDialog name={spec.name} scope={spec.scope} onClose={onClose} />;
    case 'names-from-selection':
      return <NamesFromSelectionDialog onClose={onClose} />;
    case 'sort':
      return <SortDialog onClose={onClose} />;
    case 'text-to-columns':
      return <TextToColumnsDialog onClose={onClose} />;
    case 'validation':
      return <ValidationDialog onClose={onClose} />;
    case 'remove-duplicates':
      return <RemoveDuplicatesDialog onClose={onClose} />;
    case 'group':
      return <GroupDialog ungroup={spec.ungroup} onClose={onClose} />;
    case 'subtotal':
      return <SubtotalDialog onClose={onClose} />;
    case 'outline-settings':
      return <OutlineSettingsDialog onClose={onClose} />;
    case 'goal-seek':
      return <GoalSeekDialog onClose={onClose} />;
    case 'data-table':
      return <DataTableDialog onClose={onClose} />;
    case 'scenarios':
      return <ScenariosDialog onClose={onClose} />;
    case 'series':
      return <SeriesDialog onClose={onClose} />;
    case 'find':
      return <FindDialog tab={spec.tab} onClose={onClose} />;
    case 'goto-special':
      return <GotoSpecialDialog onClose={onClose} />;
    case 'zoom':
      return <ZoomDialog onClose={onClose} />;
    case 'page-setup':
      return <PageSetupDialog tab={spec.tab} onClose={onClose} />;
    case 'insert-cells':
      return <InsertCellsDialog onClose={onClose} />;
    case 'delete-cells':
      return <InsertCellsDialog remove onClose={onClose} />;
    case 'size':
      return <SizeDialog what={spec.what} onClose={onClose} />;
    case 'protect-sheet':
      return <ProtectSheetDialog onClose={onClose} />;
    case 'unprotect-sheet':
      return <UnprotectSheetDialog onClose={onClose} />;
    case 'protect-workbook':
      return <ProtectWorkbookDialog onClose={onClose} />;
    case 'move-copy-sheet':
      return <MoveCopySheetDialog index={spec.index} onClose={onClose} />;
    case 'unhide-sheet':
      return <UnhideSheetDialog onClose={onClose} />;
    case 'paste-special':
      return <PasteSpecialDialog onClose={onClose} />;
    case 'spread-settings':
      return <SpreadSettingsDialog tab={spec.tab} onClose={onClose} />;
    case 'sheet-settings':
      return <SheetSettingsDialog tab={spec.tab} onClose={onClose} />;
    case 'customize-colors':
      return <CustomizeColorsDialog onClose={onClose} />;
    case 'alt-text':
      return <AltTextDialog id={spec.id} onClose={onClose} />;
    case 'command-palette':
      return <CommandPalette onClose={onClose} />;
    case 'new-cell-style':
      return <NewCellStyleDialog onClose={onClose} />;
    case 'shortcuts':
      return <ShortcutsDialog onClose={onClose} />;
    case 'symbol':
      return <SymbolDialog onClose={onClose} />;
    case 'shape-text':
      return <ShapeTextDialog id={spec.id} onClose={onClose} />;
    default:
      return null;
  }
}
