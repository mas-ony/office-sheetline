/** @sheetline/vue: the Sheetline spreadsheet editor as a Vue component. */
export { default as SheetlineEditor } from './SheetlineEditor.vue';
export { mount, type MountedEditor } from './mount';
export type {
  SheetlineOptions,
  AppSettings,
  DocumentInput,
  DocumentSource,
  ExportFormat,
  FileActions,
  Layout,
  SpreadsheetApi,
  ThemeName,
  ViewSettings,
} from '@sheetline/core';
