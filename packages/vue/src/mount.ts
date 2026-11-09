import type { SheetlineOptions, SpreadsheetApi } from '@sheetline/core';
import { createApp, h, shallowReactive } from 'vue';
import SheetlineEditor from './SheetlineEditor.vue';

export interface MountedEditor {
  /** Resolves with the page API once the initial workbook has opened. */
  ready: Promise<SpreadsheetApi>;
  /** Change options, as new props would. */
  update(options: SheetlineOptions): void;
  /** Remove the editor from the page. */
  destroy(): void;
}

/** Render the editor into an element, without writing Vue code. */
export function mount(element: HTMLElement, options: SheetlineOptions = {}): MountedEditor {
  let resolve!: (api: SpreadsheetApi) => void;
  const ready = new Promise<SpreadsheetApi>((r) => (resolve = r));
  const current = shallowReactive<SheetlineOptions>({ ...options });
  const app = createApp({
    render() {
      const { onReady, onLayoutChange, onThemeChange, ...props } = current;
      return h(SheetlineEditor, {
        ...props,
        onReady: (api: SpreadsheetApi) => {
          resolve(api);
          onReady?.(api);
        },
        'onUpdate:layout': onLayoutChange,
        'onUpdate:theme': onThemeChange,
      });
    },
  });
  app.mount(element);
  return {
    ready,
    update(next) {
      Object.assign(current, next);
    },
    destroy() {
      app.unmount();
    },
  };
}
