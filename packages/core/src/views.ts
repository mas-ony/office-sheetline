/**
 * What the official interfaces (@sheetline/react and @sheetline/vue) are built from: the public
 * API plus the interface's commands, ribbon and menu data, and the engine pieces its dialogs use.
 * It follows those packages and is not a stable API of its own.
 */
export * from './index';

// Engine pieces the dialogs, panels and the backstage use.
export { type PasteOptions, type PasteWhat, type SearchHit, type Selection } from './core/document';
export {
  absRef, chartFromRange, guessLayout, localRef, readChartData, resolveRef, seriesForLayout, trimRange,
} from './core/chartData';
export {
  type CellIsOperator, type CfFormat, type CfRule, type Cfvo, type CfvoType, COLOR_SCALES, DATA_BAR_COLORS,
  defaultThresholds, describeRule, HIGHLIGHT_PRESETS, ICON_SET_GROUPS, iconCount, type IconSetName,
  type TextOperator, TIME_PERIODS, type TimePeriod,
} from './core/condFormats';
export {
  CUSTOM_LISTS, type SeriesOptions, type SortLevel, splitText, type TextToColumnsOptions,
} from './core/doc/dataTools';
export { SUBTOTAL_FUNCTIONS } from './core/doc/outline';
export { type SpecialKind } from './core/doc/select';
export { formatValue } from './core/format/display';
export { tokenize } from './core/formula/tokenizer';
export { type ColorSpec, resolveColor, rgbSpec, type Theme, type ThemeSlot } from './core/model/color';
export {
  type ChartSeries, type LineDash, type SeriesKind, type SheetChart, type SheetShape, type Side,
  type Trendline,
} from './core/model/drawing';
export {
  HEADER_PRESETS, type HeaderFooterSet, hfJoin, hfSections, hfText, MARGIN_PRESETS, type PageSetup, PAPER,
  paperOf,
} from './core/model/page';
export { DEFAULT_PROTECTION, PROTECTION_LABELS, type ProtectionOptions } from './core/model/protection';
export { type Scenario } from './core/model/scenario';
export {
  type ColumnFilter, type DataValidation, type FilterCondition, type FilterOperator, ptToPx, pxToPt,
  type Sheet, type ValidationOperator, widthCharsToPx, widthPxToChars,
} from './core/model/sheet';
export {
  BORDER_LINE_STYLES, type BorderLine, type BorderLineStyle, type CellStyle, type HAlign, type StylePatch,
  type UnderlineType, type VAlign,
} from './core/model/style';
export { type DefinedName } from './core/model/workbook';
export {
  cellName, colName, makeRange, parseCellRef, quoteSheetName, type Range, rangesIntersect,
} from './core/refs';
export { galleryStyles, tableStyleLabel, tableStylePreviewSvg } from './core/tableStyles';
export { THEME_SLOT_LABELS } from './core/themes';
export { validationAt } from './core/validation';
export { CellError } from './core/values';
export { type XmlElement } from './core/xml/xml';
export { REF_COLORS } from './grid/GridView';
export { chartSvg } from './grid/chartRender';
export { drawIcon } from './grid/decor';
export { SHAPE_GALLERY, type GalleryItem as ShapeGalleryItem, shapeIconSvg } from './grid/shapes';
export { canvasMeasure, loadPrintFonts, printPages, rasterize, saveFile } from './print/browser';
export { buildPages, type Measure, type Page } from './print/pages';
export { pagesToPdf } from './print/pdf';
export { type PrintChoice, printOptionsFor } from './print/setup';
export { pageSvg } from './print/svg';

// The interface without its components: commands, menus, the ribbon's data, palettes and icons.
export * from './ui/cellStyles';
export * from './ui/charts';
export * from './ui/commands';
export * from './ui/condFormatRules';
export * as cmd from './ui/commands';
export * from './ui/dataDialogs';
export * from './ui/docState';
export * from './ui/dialogs';
export * from './ui/dom';
export * from './ui/formatCells';
export * from './ui/formulaAssist';
export * from './ui/functionCatalog';
export * from './ui/icons';
export * from './ui/menuItems';
export * from './ui/objects';
export * from './ui/palette';
export * from './ui/popovers';
export * from './ui/rangeText';
export * from './ui/ribbon/contextTabs';
export * from './ui/ribbon/formulasTab';
export * from './ui/ribbon/galleries';
export { themeTiles } from './ui/ribbon/pageLayoutTab';
export * from './ui/ribbon/tabs';
export * from './ui/ribbon/types';
export * from './ui/statusBar';
