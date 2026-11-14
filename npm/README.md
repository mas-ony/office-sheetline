# Sheetline

Sheetline is an Excel (XLSX) editor for desktop and mobile that runs in the browser: a ribbon with
contextual tabs, a toolbar layout, light and dark themes, a File backstage, dialogs and context
menus, on its own canvas grid, formula engine and XLSX reader and writer. On phones it has a touch
layout, and it also works over a plain HTTP address on the local network.

This package is the built app: static files to put on a web server. It has no dependencies and
nothing to build. The React and Vue packages, for putting the editor inside an app instead, are
[`@sheetline/react`](https://www.npmjs.com/package/@sheetline/react) and
[`@sheetline/vue`](https://www.npmjs.com/package/@sheetline/vue).

## Install

```bash
npm install sheetline
```

## Host

The app is the package's `dist` folder: `index.html`, `assets/`, the service worker, the manifest
and the icons. Its paths are relative, so it can be served from any path. Copy the folder into a
site, or serve it as it is:

```bash
cp -r node_modules/sheetline/dist public/sheets               # then open /sheets/
python3 -m http.server -d node_modules/sheetline/dist 8080    # then open http://localhost:8080
```

Serve it over HTTP: browsers do not run its scripts from a `file://` address.

Over HTTPS (or on `localhost`) the app also works offline and can be installed, and an installed
app opens spreadsheet files from the system's "Open with". Over plain HTTP, such as a local
network address, the clipboard API, file pickers, the share sheet and offline use are not
available, and the app works without them.

## Files

It opens XLSX, XLSM, XLTX and XLTM (saving patches the opened package, so everything the editor
does not change is kept), XLS (Excel 97-2003), XLSB, CSV and TSV, by drag and drop or from the File
backstage. It saves XLSX, XLTX, XLSM, XLTM, CSV and PDF. Files are opened, edited and saved in the
browser; the app sends nothing to a server.

## Address options

| Option | |
| --- | --- |
| `?src=<url>` | Opens a workbook from a URL the page can fetch (its own server, or one that allows it with CORS) |
| `?new=1` | Starts with a blank workbook, without offering to recover unsaved work |
| `?readonly=1` or `?mode=view` | Opens workbooks view-only |

## Page API

Once a workbook is open, the app puts a programmatic interface on `window.spreadsheetEditor`. A
page that shows the app in a same-origin `<iframe>` reaches it as
`iframe.contentWindow.spreadsheetEditor`.

```js
const ed = window.spreadsheetEditor;
ed.setData([['Item', 'Price'], ['Pen', 1.5]], 'A1');
ed.setValue('C2', '=B2*2');
ed.setStyle('A1:B1', 'font-weight', 'bold');
ed.setFormat('B2', '"$" #,##0.00');
ed.on('change', () => console.log(ed.getValue('C2')));
ed.setLayout('toolbar');   // or 'ribbon'
ed.setTheme('dark');       // or 'light'
ed.download('xlsx');       // also 'xltx', 'xlsm', 'xltm', 'csv'; 'pdf' opens File > Print
await ed.open(blob, 'Budget.xlsx');   // then window.spreadsheetEditor serves the new workbook
```

It also has `getText`, `getData`, `insertRow`, `insertColumn`, `deleteRow`, `deleteColumn`,
`orderBy`, `setFreezeRows`, `setFreezeColumns`, `select`, `getSelected`, `undo`, `redo`,
`worksheets`, `createWorksheet`, `renameWorksheet`, `deleteWorksheet` and `save`. Its events are
`change`, `selection`, `activeSheet`, `sheets`, `file` (the workbook was renamed or saved) and
`open` (another workbook was opened); `on` returns a function that removes the listener.

## Storage

The app keeps unsaved work in IndexedDB (the database `sheetline`) and offers to recover it on the
next visit. It keeps its settings in localStorage under `sheetline-` keys, and its offline copy in
the cache `sheetline-v1`. Copies served from the same origin share them.

## Browsers

Tested in Chrome on desktop and on Android.

## License

MIT. The app bundles third-party packages, icons and fonts: `dist/THIRD-PARTY-LICENSES.md` lists
them with their licences (MIT; Apache-2.0 for the Material Design icons and `bessel`; the SIL Open
Font License 1.1 for the Arimo, Carlito, Cousine and Tinos fonts).
