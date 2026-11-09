# @sheetline/vue

Sheetline is an Excel (XLSX) editor for desktop and mobile that runs in the browser: a ribbon with
contextual tabs, a one-row toolbar layout, light and dark themes, a File backstage, dialogs and
context menus, on its own canvas grid, formula engine and XLSX reader and writer. It also opens XLS,
XLSB and CSV files, prints and writes PDF in the browser, and has a touch layout on phones.

This package is the Vue interface, written in Vue single-file components. `@sheetline/react` is
the same interface written in React, with the same options.

## Install

```bash
npm install @sheetline/vue vue
```

It needs Vue 3.5 or later.

## Use

```vue
<script setup lang="ts">
import { SheetlineEditor, type SpreadsheetApi, type ThemeName } from '@sheetline/vue';
import '@sheetline/vue/style.css';
import { ref } from 'vue';

const theme = ref<ThemeName>('light');
const onReady = (api: SpreadsheetApi) => api.on('change', () => console.log(api.getValue('B10')));
</script>

<template>
  <SheetlineEditor
    v-model:theme="theme"
    style="height: 80vh"
    document="/reports/budget.xlsx"
    author="Jordan Lee"
    @ready="onReady"
  />
</template>
```

Import `style.css` once: it brings the fonts and the editor's styles. The editor fills the height
of its element, so give it (or its container) a height.

`mount()` takes the same options, for pages that are not written in Vue:

```js
import { mount } from '@sheetline/vue';
import '@sheetline/vue/style.css';

const editor = mount(document.querySelector('#editor'), { document: file, theme: 'dark' });
const api = await editor.ready;
editor.update({ readOnly: true });
editor.destroy();
```

## Props

| Prop | |
| --- | --- |
| `document` | A URL, Blob, ArrayBuffer or bytes, or `{ url or data, name }`; opened again when it changes. Without it the editor starts with a blank Book1 |
| `read-only` | Workbooks open view-only |
| `layout` | `'ribbon'` or `'toolbar'` (the one-row toolbar, on desktop); by default the one last chosen, else the ribbon |
| `theme` | `'light'` or `'dark'` |
| `author` | The name on comments and notes, written as the author of new workbooks |
| `storage-key` | Keep the layout, theme, settings, recent colours and cell styles in localStorage under `<storageKey>-…` |
| `autosave`, `offer-recovery` | Keep unsaved changes in an IndexedDB database of this name and offer to recover them when the editor starts without a document |
| `files` | Replace some of the file actions (New, Open, Save, Export, Import, Print, Rename, Share) |
| `global-shortcuts` | Handle the editor's shortcuts (Ctrl+1, Alt+Q, F9…) anywhere on the page, not only inside the editor |

The types are exported as `SheetlineOptions`, `ThemeName`, `Layout`, `FileActions` and so on.

## Events

`@ready` fires once the editor is ready and the initial workbook has opened, with the editor's page
API. `@update:theme` fires with the theme in use when the editor starts and whenever the reader
switches it, and `@update:layout` when the reader switches between the ribbon and the toolbar, so
`v-model:theme` and `v-model:layout` work.

## Page API

The object passed to `@ready` drives the editor from code. It follows the workbook the editor
shows.

```js
api.setData([['Item', 'Price'], ['Pen', 1.5]], 'A1');
api.setValue('C2', '=B2*2');
api.setStyle('A1:B1', 'font-weight', 'bold');
api.setFormat('B2', '"$" #,##0.00');
api.on('change', () => console.log(api.getValue('C2')));
api.setTheme('dark');
api.download('xlsx');   // also 'xltx', 'xlsm', 'xltm', 'csv'; 'pdf' opens the Print backstage
await api.open(blob, 'Budget.xlsx');
```

`installPageApi(api)` from `@sheetline/core` also puts it on the page as
`window.spreadsheetEditor`.

## Page styles

The editor keeps its classes on its own element and on a layer it adds to the body for menus,
dialogs and messages (rendered there with `Teleport`): `sh-` for the interface, `xgrid-` and
`xchart-` for the grid. It changes nothing on `<html>` or `<body>`, and several editors can share a
page. In a window up to 720 pixels wide it uses its phone layout.

## License

MIT
