/**
 * What makes the editor a full-page app: the workbook and mode in the address, stored settings and
 * autosave, the window title, warning before leaving with unsaved changes, files passed by the
 * installed app, and the page API on `window`.
 */
import { type SheetlineOptions, installPageApi, nameOfUrl, type SpreadsheetApi } from '@sheetline/core';

const APP = 'Sheetline';

/** Options from the address: `?src=` opens a workbook, `?new=1` skips recovery, `?readonly=1` or `?mode=view` opens view-only. */
export function shellOptions(): SheetlineOptions {
  const params = new URLSearchParams(location.search);
  const src = params.get('src');
  return {
    storageKey: 'sheetline',
    autosave: 'sheetline',
    offerRecovery: params.get('new') !== '1',
    globalShortcuts: true,
    readOnly: params.get('readonly') === '1' || params.get('mode') === 'view',
    document: src ? { url: src, name: nameOfUrl(src) } : undefined,
    // The page's own background follows the editor's theme.
    onThemeChange: (theme) => {
      document.documentElement.dataset.theme = theme;
    },
  };
}

interface LaunchQueue {
  setConsumer(fn: (p: { files: FileSystemFileHandle[] }) => void): void;
}

/** Connect the ready editor to the page; returns the function that disconnects it. */
export function connectShell(api: SpreadsheetApi): () => void {
  const title = (name: string, modified: boolean) => {
    document.title = `${modified ? '• ' : ''}${name} – ${APP}`;
  };
  title(api.doc.fileName, false);
  const offs = [
    api.on('open', ({ fileName }) => title(fileName, false)),
    api.on('change', () => {
      if (api.doc.modified) title(api.doc.fileName, true);
    }),
    api.on('file', ({ name }) => title(name, api.doc.modified)),
    installPageApi(api),
  ];
  const before = (e: BeforeUnloadEvent) => {
    if (api.doc.modified) {
      e.preventDefault();
      e.returnValue = '';
    }
  };
  window.addEventListener('beforeunload', before);

  // Files opened through the installed app (PWA file handling).
  (window as unknown as { launchQueue?: LaunchQueue }).launchQueue?.setConsumer(async (p) => {
    const handle = p.files?.[0];
    if (handle) void api.open(handle);
  });

  return () => {
    for (const off of offs) off();
    window.removeEventListener('beforeunload', before);
  };
}

/** The installed app works offline (production builds in secure contexts; plain-HTTP LAN addresses skip it). */
export function registerServiceWorker(): void {
  if (window.isSecureContext && 'serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js').catch(() => undefined);
    });
  }
}
