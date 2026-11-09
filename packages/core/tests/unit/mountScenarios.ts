/**
 * Scenarios for the mount() of the React and Vue packages (run by their own tests): the same
 * library behaviour from both interfaces.
 */
import { afterEach, expect, it, vi } from 'vitest';
import { version } from '../../package.json';
import type { SpreadsheetApi } from '../../src/api/api';
import type { SheetlineOptions } from '../../src/ui/controller';
import { readFixture } from './helpers';

interface Mounted {
  ready: Promise<SpreadsheetApi>;
  update(options: SheetlineOptions): void;
  destroy(): void;
}

const wait = { timeout: 3000 };
/** Opening a workbook in jsdom takes a moment; the first test also loads the modules. */
const slow = 20_000;

/** The element once it exists. */
function find<T extends Element = HTMLElement>(selector: string, root: ParentNode = document): Promise<T> {
  return vi.waitFor(() => {
    const el = root.querySelector<T>(selector);
    if (!el) throw new Error(`${selector} not found`);
    return el;
  }, wait);
}

/** The editor's root element inside the element it was mounted in. */
const editorIn = (el: HTMLElement) => el.querySelector<HTMLElement>('[data-testid="editor"]')!;

export function mountScenarios(mount: (element: HTMLElement, options?: SheetlineOptions) => Mounted) {
  const open: Mounted[] = [];
  const place = () => document.body.appendChild(document.createElement('div'));
  const start = (el: HTMLElement, options: SheetlineOptions) => {
    const m = mount(el, options);
    open.push(m);
    return m;
  };

  afterEach(() => {
    for (const m of open.splice(0)) m.destroy();
    document.body.innerHTML = '';
    localStorage.clear();
  });

  it('opens the workbook, keeps its classes off <html> and removes everything again', async () => {
    const el = place();
    const m = start(el, { document: { data: readFixture('kitchen-sink.xlsx'), name: 'Quarterly report.xlsx' }, theme: 'dark' });
    const api = await m.ready;
    expect(api.doc.fileName).toBe('Quarterly report.xlsx');
    expect(api.version).toBe(version);
    expect(api.getValue('Sales!A1')).toBe('Quarterly Sales Report');
    await vi.waitFor(() => expect(el.contains(api.grid!.root)).toBe(true), wait);
    const editor = editorIn(el);
    const layers = document.querySelectorAll<HTMLElement>('.sh-layer');
    expect(editor.dataset.theme).toBe('dark');
    expect(layers).toHaveLength(1);
    expect(layers[0].dataset.theme).toBe('dark');
    expect(document.documentElement.className).toBe('');
    expect(document.documentElement.hasAttribute('data-theme')).toBe(false);
    expect(document.body.className).toBe('');

    m.update({ theme: 'light' });
    await vi.waitFor(() => expect(editor.dataset.theme).toBe('light'), wait);
    expect(layers[0].dataset.theme).toBe('light');

    m.destroy();
    open.length = 0;
    await vi.waitFor(() => expect(document.querySelector('.sh-layer')).toBeNull(), wait);
    expect(el.childElementCount).toBe(0);
  }, slow);

  it('shows dialogs in its layer and reports the theme and layout the reader picks', async () => {
    const el = place();
    const onThemeChange = vi.fn();
    const onLayoutChange = vi.fn();
    const m = start(el, { theme: 'light', layout: 'ribbon', onThemeChange, onLayoutChange });
    await m.ready;
    expect(onThemeChange).toHaveBeenCalledWith('light');
    const editor = editorIn(el);

    (await find('[data-testid="theme-toggle"]', el)).click();
    await vi.waitFor(() => expect(onThemeChange).toHaveBeenLastCalledWith('dark'), wait);
    await vi.waitFor(() => expect(editor.dataset.theme).toBe('dark'), wait);
    (await find('[data-testid="layout-toggle"]', el)).click();
    await vi.waitFor(() => expect(onLayoutChange).toHaveBeenLastCalledWith('toolbar'), wait);
    await vi.waitFor(() => expect(editor.dataset.layout).toBe('toolbar'), wait);

    (await find('[data-testid="fx-insert"]', el)).click();
    const dialog = await find('[data-testid="insert-function"]');
    expect(dialog.closest('.sh-layer')).not.toBeNull();
    expect(el.contains(dialog)).toBe(false);
    expect(dialog.closest<HTMLElement>('.sh-layer')!.dataset.theme).toBe('dark');
  }, slow);

  it('keeps two editors on one page apart', async () => {
    const a = place();
    const b = place();
    const [ra, rb] = await Promise.all([
      start(a, { document: { data: readFixture('kitchen-sink.xlsx'), name: 'Report.xlsx' }, theme: 'light' }).ready,
      start(b, { document: { data: readFixture('modern.xlsx'), name: 'Stock.xlsx' }, theme: 'light' }).ready,
    ]);
    expect(ra).not.toBe(rb);
    expect(ra.worksheets()).toContain('Sales');
    expect(rb.worksheets()).toEqual(['Stock']);
    await vi.waitFor(() => {
      expect(a.contains(ra.grid!.root)).toBe(true);
      expect(b.contains(rb.grid!.root)).toBe(true);
    }, wait);
    expect(document.querySelectorAll('.sh-layer')).toHaveLength(2);

    rb.setValue('A1', 'Only in B');
    expect(rb.getValue('A1')).toBe('Only in B');
    expect(ra.getValue('Sales!A1')).toBe('Quarterly Sales Report');
    ra.select('C3');
    expect(ra.getSelected()).toBe('C3');
    expect(rb.getSelected()).not.toBe('C3');

    ra.setTheme('dark');
    await vi.waitFor(() => expect(editorIn(a).dataset.theme).toBe('dark'), wait);
    expect(editorIn(b).dataset.theme).toBe('light');
  }, slow);
}
