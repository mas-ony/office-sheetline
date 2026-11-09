import type { SheetlineOptions, SpreadsheetApi } from '@sheetline/core';
import { createRoot } from 'react-dom/client';
import { SheetlineEditor } from './SheetlineEditor';

export interface MountedEditor {
  /** Resolves with the page API once the initial workbook has opened. */
  ready: Promise<SpreadsheetApi>;
  /** Change options, as new props would. */
  update(options: SheetlineOptions): void;
  /** Remove the editor from the page. */
  destroy(): void;
}

/** Render the editor into an element, without writing React code. */
export function mount(element: HTMLElement, options: SheetlineOptions = {}): MountedEditor {
  const root = createRoot(element);
  let resolve!: (api: SpreadsheetApi) => void;
  const ready = new Promise<SpreadsheetApi>((r) => (resolve = r));
  let current = options;
  const render = () => {
    const onReady = (api: SpreadsheetApi) => {
      resolve(api);
      current.onReady?.(api);
    };
    root.render(<SheetlineEditor {...current} onReady={onReady} />);
  };
  render();
  return {
    ready,
    update(next) {
      current = { ...current, ...next };
      render();
    },
    destroy() {
      root.unmount();
    },
  };
}
