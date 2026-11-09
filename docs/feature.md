# Sheetline features

What Sheetline does, area by area. [The README](../README.md) covers running it, the components
and the tests.

## Platforms and files

| Feature | Details |
| --- | --- |
| Browsers | Tested in Chrome on desktop and on Android |
| Desktop | Ribbon or one-row toolbar layout, light and dark theme |
| Phones | Ribbon groups, menus and context menus as bottom sheets, long press for the cell menu, pinch zoom, a key bar above the keyboard, full-screen backstage and dialogs, charts placed under the data |
| Plain-HTTP LAN address | No secure-context API is required; clipboard, file pickers, sharing and offline caching are used only where available |
| Start | A blank "Book1"; `?src=<url>` opens a file, and `?new=1` skips the recovery offer |
| File format | Excel files are the native format: Open, Save, and Save back to the opened file where the browser allows |
| Open XLSX, XLSM, XLTX, XLTM | From the File backstage, by drag and drop and with the system's "Open with" for the installed app |
| Open XLS (Excel 97-2003) and XLSB | Own BIFF8 and BIFF12 readers; saved as XLSX |
| CSV | Import and export with delimiter and encoding detection; TSV |
| Save as Excel | XLSX, XLTX, XLSM and XLTM; the opened package is patched, so its unknown parts and elements and anything the editor does not change are kept |
| PDF | Export and Print through its own PDF writer: vector text and lines; pictures, charts, data bars, icons and sparklines embedded as images |
| New on a modified workbook | Asks to save (Yes, No, Cancel) |
| Autosave and recovery | A copy in IndexedDB, offered back with a restore banner |
| Installable app, offline | PWA with file handlers (offline caching needs HTTPS or localhost) |

## Ribbon and layout

| Feature | Details |
| --- | --- |
| Tabs | File, Home, Insert, Page Layout, Formulas, Data, Review, View, Settings |
| Contextual tabs | Table Design, Chart Design, Shape Format, picture and sparkline tabs |
| Ribbon controls | Group launchers and a caret that collapses the ribbon |
| Toolbar layout | A one-row simplified ribbon, with an overflow "…" for controls that do not fit |
| Light and dark theme | A moon/sun switch beside the ribbon tabs; cells keep their colours, grid headers and the selection turn dark |
| Command Palette | Finds commands, including items inside ribbon menus ("Freeze Panes › Freeze Top Row"); Alt+Q |
| Keyboard shortcuts | Excel's shortcuts |
| Formula bar | Name Box, ✕ ✓ fx, an expandable input; coloured references; argument tips |
| Status bar | Ready/Enter/Edit, Average/Count/Sum (chosen with a right click), zoom slider |
| Sheet tabs | Navigation, all-sheets list, "+", and a tab menu: Insert, Delete, Rename, Move or Copy, Protect, Tab Color, Hide, Unhide |
| File backstage | New, Open, Save, Import, Export, Print, Info |

## Home

| Feature | Details |
| --- | --- |
| Undo and redo | With lists; every edit, including charts, shapes, outlines and themes |
| Clipboard | Cut, Copy, Paste options and Paste Special (All, Formulas, Values, Formatting, Transpose, Column widths, Link, …); Format Painter (a double click keeps it on) |
| Font | Font, size, grow/shrink, bold, italic, underline, double underline, colours |
| Borders | 13 presets and More Borders |
| Alignment | Alignment, orientation, indent, wrap, merge menu |
| Number formats | Format list, accounting, percent, comma, decimals |
| Conditional Format | Highlight, Top/Bottom, Data Bars, Color Scales, Icon Sets, New Rule, Clear, Manage: a rule editor with Excel's rule types and a Rules Manager; written to XLSX (with the Excel 2010 extension) |
| Format Table | Light, Medium and Dark galleries |
| Cell Styles | Excel's built-in styles; new styles are kept in the browser, not in the file |
| Cells | Insert, Delete and Format menus: insert or delete cells with shifting; row height, column width, autofit, hide/unhide, protect, lock |
| Editing | Sum, Fill (with the Fill Series dialog), Clear, Sort & Filter, Find, Go To Special |

## Insert

| Feature | Details |
| --- | --- |
| Table | Create Table with the Medium 2 style and filter buttons; Table Design: name, resize, convert to range, Allow Auto Expand, style options, styles, remove duplicates |
| Insert Chart | 14 categories (including Statistic: histogram, Pareto, box and whisker) and 41 subtypes; custom combo |
| New chart | Right of the data (under it on phones), 480 x 300, "Chart Title", Office 2013 colours |
| Chart Design | Elements, quick layouts, colours, styles, switch row/column, select data, change type, move chart, alt text, Format Pane; saved as chart parts (chartex for sunburst, treemap, funnel, waterfall, histogram, Pareto, box and whisker); charts in opened files are drawn, edited and written back unchanged until edited |
| Sparklines | Line, column and win/loss groups with markers, colours, axis, grouping; saved as the Excel 2010 extension |
| Pictures | Insert and Picture Format: change picture, alt text, size, order |
| Shapes | 8 categories, 125 presets, text boxes, Shape Format (fill, outline, styles, text, rotate, align, size) |
| Hyperlink | Web page, this document, e-mail |
| Threaded comments and notes | With a comments panel |
| Symbol | A dialog to insert symbols |

## Page Layout and printing

| Feature | Details |
| --- | --- |
| Themes | 21 themes, 22 colour and 21 font schemes, Customize Colors; cells with theme colours and fonts follow; the theme part is written to the file |
| Page setup | Margins, orientation, size, print area, breaks, print titles; the Page Setup dialog (Page, Margins, Header/Footer, Sheet) with a custom header and footer editor |
| Scale to Fit, sheet options | Gridlines and headings, for viewing and printing |
| Print backstage | Settings and a live page preview; Save as PDF next to Print |
| Printed content | Cells, data bars, icon sets, sparklines, pictures, charts, shapes, headers and footers |

## Formulas

| Feature | Details |
| --- | --- |
| Functions | 499 with signatures and descriptions (dynamic arrays, LET, LAMBDA and its helpers, GROUPBY, PIVOTBY, financial functions for odd periods, FILTERXML, …) |
| Insert Function | The Insert Function dialog and the function library menus |
| Suggestions and argument tips | While typing, in the cell and in the formula bar |
| Names | Name Manager, Define Name, Create from Selection; sheet-scoped names and comments |
| Calculation | Show Formulas, calculation options, Calculate Now / Sheet |

## Data

| Feature | Details |
| --- | --- |
| Sort | A→Z, Z→A, and the Sort dialog with levels: values, cell colour, font colour, custom lists; by rows or columns |
| Filter | Filter, Clear, Reapply; value lists, conditions, colours, search; saved to XLSX |
| Text to Columns | Delimited and fixed width with preview |
| Data Validation | Settings, Input Message, Error Alert; Circle Invalid Data marks invalid cells |
| Remove Duplicates | Removes repeated rows, comparing the columns chosen |
| Outline | Group, Ungroup, Auto Outline, Subtotal, Show/Hide Detail; outline buttons and levels; saved to XLSX |
| What-If Analysis | Scenario Manager, Goal Seek, Data Table |

## Review, View and Settings

| Feature | Details |
| --- | --- |
| Threaded comments | New, delete, previous, next, panel |
| Protection | Protect sheet and workbook; passwords stored as Excel hashes |
| Show/hide | Headers, gridlines, tab strip, new tab button, formula bar, status bar, zero values |
| Focus Cell | Highlights the row and column of the active cell |
| Zoom | Zoom dialog, 100%, Zoom to Selection; Ctrl+wheel and pinch |
| Freeze Panes | Freeze Panes, Top Row, First Column |
| Selection Pane | Show/hide, rename, order |
| Spread Settings | General, Scroll Bars, Calculation, Tab Strip |
| Sheet Settings | General, Grid Lines, Headers, Sheet Tab |

## Programming interface

The editor is embedded with the `SheetlineEditor` component of `@sheetline/react` or
`@sheetline/vue`, or with their `mount()`. Its page API (`packages/core/src/api/api.ts`), which
the apps install as `window.spreadsheetEditor`, addresses cells with A1 references (`'B2'`,
`'Sheet2!A1:C3'`). `ed.doc` gives access to the native document model
(`packages/core/src/core/document.ts`) and `ed.grid` to the grid.

| Feature | Details |
| --- | --- |
| Components | `SheetlineEditor` for React and for Vue, each written for its framework, with the same props and events (Vue: `v-model:theme`, `v-model:layout`), and `mount(element, options)` for other pages; styles scoped to the editor, several editors on a page |
| Open a file | `ed.open(blob, name)` |
| Layout and theme | `ed.setLayout('ribbon' / 'toolbar')`, `ed.setTheme('light' / 'dark')` |
| Values, formulas, data | `getValue`, `getText`, `setValue`, `getData`, `setData` |
| Styles and formats | `setStyle(ref, cssProperty, value)`, `setFormat(ref, code)` |
| Rows, columns, sheets | `insertRow`, `insertColumn`, `deleteRow`, `deleteColumn`, `orderBy`, `setFreezeRows/Columns`, `createWorksheet`, `renameWorksheet`, `deleteWorksheet`, `worksheets` |
| Events | `on('change' / 'selection' / 'activeSheet' / 'sheets' / 'file' / 'open')` |
| Save and export | `download('xlsx' / 'xltx' / 'xlsm' / 'xltm' / 'csv' / 'pdf')`, `save()` |

## Limitations

- Excel 5.0/95 `.xls` files are not supported; XLS and XLSB files are saved as XLSX. Encrypted
  (password-protected) files cannot be opened or written.
- PivotTables, slicers, form controls, camera pictures, external data connections and sheet
  background pictures are not supported, and the ribbon has no KeyTips (Excel's keyboard
  shortcuts work instead).
- What-If Data Table writes values instead of a live `TABLE()` array.
- Cell styles created in the editor are kept in the browser, not in the file.
- Charts from files that use features the chart model does not cover (3-D charts) are drawn
  approximately; they are written back unchanged unless edited.
- Text that the standard PDF fonts cannot show (for example CJK) is embedded in the PDF as
  pictures.
- The ribbon cannot be configured.
