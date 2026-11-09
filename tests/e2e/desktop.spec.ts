import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { SALES, clickCell, docEval, downloadOf, input, openApp, parts, select, setData, shown, typeIn } from './helpers';

const KITCHEN = '/tests/fixtures/kitchen-sink.xlsx';
const RIBBON_TABS = ['home', 'insert', 'page-layout', 'formulas', 'data', 'review', 'view', 'settings'];

test.describe('desktop editor', () => {
  test('moves, reorders and deletes objects; shapes with text', async ({ page }) => {
    const errors = await openApp(page);
    await clickCell(page, 'B3');
    await page.getByTestId('tab-insert').click();
    await page.getByTestId('rb-insert-shapes').click();
    const gallery = page.getByTestId('shapes-panel');
    for (const group of ['Lines', 'Rectangles', 'Basic Forms', 'Block Arrows', 'Stars', 'Callouts', 'Equation Shapes', 'Flowchart']) await expect(gallery).toContainText(group);
    await gallery.getByTestId('shape-roundRect').click();
    const shape = page.getByTestId('object-shape');
    await expect(shape).toHaveCount(1);
    expect(await docEval(page, '(d) => d.sheet.shapes[0].from')).toMatchObject({ r: 2, c: 1 });
    // A double-click opens the Format Shape pane.
    await shape.dblclick();
    const pane = page.getByTestId('format-pane');
    await expect(pane.getByTestId('shape-fill-color')).toContainText('#4F81BD');
    await pane.getByTestId('shape-text-pane').fill('Hello');
    await pane.getByTestId('shape-text-pane').press('Enter');
    await expect(shape).toContainText('Hello');

    // Drag to move.
    const box = (await shape.boundingBox())!;
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 + 150, box.y + box.height / 2 + 60, { steps: 5 });
    await page.mouse.up();
    expect(await docEval(page, '(d) => d.sheet.shapes[0].from.c')).toBeGreaterThan(1);
    await page.keyboard.press('Control+z');
    expect(await docEval(page, '(d) => d.sheet.shapes[0].from.c')).toBe(1);

    // Object menu: delete.
    await shape.click({ button: 'right' });
    await expect(page.getByTestId('obj-front')).toBeVisible();
    await page.getByTestId('obj-delete').click();
    await expect(shape).toHaveCount(0);
    expect(errors).toEqual([]);
  });

  test('opens a workbook with a chart and a picture and keeps the chart part unchanged', async ({ page }) => {
    const errors = await openApp(page, KITCHEN);
    await expect(page.getByTestId('tab-Sales')).toBeVisible();
    expect(await shown(page, 'D8')).toBe('$6,501.25');
    await expect(page.getByTestId('object-image')).toHaveCount(1);
    const chart = page.getByTestId('object-chart');
    await expect(chart).toHaveCount(1);
    await expect(chart).toContainText('Revenue by region');
    const { bytes } = await downloadOf(page, () => page.keyboard.press('Control+s'));
    const original = parts(new Uint8Array(readFileSync('tests/fixtures/kitchen-sink.xlsx')));
    expect(parts(bytes)['xl/charts/chart1.xml']).toBe(original['xl/charts/chart1.xml']);
    expect(errors).toEqual([]);
  });

  test('data validation, filters, notes and freeze panes', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    // Filters on the data region.
    await clickCell(page, 'B2');
    await page.getByTestId('tab-data').click();
    await page.getByTestId('rb-filter').click();
    expect(await docEval(page, '(d) => !!d.sheet.autoFilter')).toBe(true);
    // Validation: a list.
    await clickCell(page, 'F2');
    await page.getByTestId('rb-data-validation').click();
    const dv = page.getByTestId('validation-dialog');
    await dv.getByTestId('validation-allow').selectOption('list');
    await dv.getByTestId('validation-f1').fill('Yes,No');
    await dv.getByTestId('validation-dialog-ok').click();
    expect(await docEval(page, '(d) => d.sheet.validations.length')).toBe(1);
    // A note from the cell menu.
    await clickCell(page, 'A6', { button: 'right' });
    await page.getByTestId('ctx-note').click();
    await page.getByTestId('note-input').fill('Check with finance');
    await page.getByTestId('note-save').click();
    expect(await docEval(page, '(d) => { const t = d.sheet.commentAt(5, 0); return t && { kind: t.kind, text: t.comments[0].text }; }')).toEqual({ kind: 'note', text: 'Check with finance' });
    // Freeze the top row from the View tab.
    await page.getByTestId('tab-view').click();
    await page.getByTestId('rb-freeze-panes').click();
    await page.getByTestId('freeze-top-row').click();
    expect(await docEval(page, '(d) => d.sheet.freeze.rows')).toBe(1);
    expect(errors).toEqual([]);
  });

  test('sheet tabs: insert, rename, hide and unhide', async ({ page }) => {
    const errors = await openApp(page);
    await page.getByTestId('tab-add').click();
    await expect(page.getByTestId('tab-Sheet2')).toBeVisible();
    await page.getByTestId('tab-Sheet2').dblclick();
    await page.getByTestId('tab-rename-input').fill('Budget');
    await page.getByTestId('tab-rename-input').press('Enter');
    await expect(page.getByTestId('tab-Budget')).toBeVisible();
    await page.getByTestId('tab-Budget').click({ button: 'right' });
    await page.getByTestId('tab-hide').click();
    await expect(page.getByTestId('tab-Budget')).toHaveCount(0);
    await page.getByTestId('tab-Sheet1').click({ button: 'right' });
    await page.getByTestId('tab-unhide').click();
    await page.getByTestId('unhide-sheet-ok').click();
    await expect(page.getByTestId('tab-Budget')).toBeVisible();
    expect(errors).toEqual([]);
  });

  test('opens Book1 with the ribbon; typing values and formulas', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto('/?new=1');
    await expect(page.getByTestId('ribbon')).toBeVisible();
    await expect(page.getByTestId('tab-home')).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('tab-Sheet1')).toBeVisible();
    await expect(page).toHaveTitle(/Book1/);
    // Row-by-row entry: Enter after Tabs returns to the first column.
    await clickCell(page, 'A1');
    for (const [i, v] of ['Item', 'Price', 'Qty'].entries()) {
      await page.keyboard.type(v);
      await page.keyboard.press(i < 2 ? 'Tab' : 'Enter');
    }
    for (const [i, v] of ['Pen', '1.5', '4'].entries()) {
      await page.keyboard.type(v);
      await page.keyboard.press(i < 2 ? 'Tab' : 'Enter');
    }
    expect(await shown(page, 'A2')).toBe('Pen');
    expect(await shown(page, 'C2')).toBe('4');
    await typeIn(page, 'D2', '=B2*C2');
    expect(await shown(page, 'D2')).toBe('6');
    // Function suggestions while typing.
    await clickCell(page, 'D3');
    await page.keyboard.type('=SU');
    await expect(page.getByTestId('formula-suggestions')).toBeVisible();
    await page.keyboard.type('M(D2,10)');
    await page.keyboard.press('Enter');
    expect(await input(page, 'D3')).toBe('=SUM(D2,10)');
    expect(await shown(page, 'D3')).toBe('16');
    // The formula bar shows it: coloured text under a transparent text area.
    await clickCell(page, 'D3');
    await expect(page.getByTestId('formula-bar').locator('div.sh-formula-input')).toHaveText('=SUM(D2,10)');
    expect(await page.getByTestId('formula-input').evaluate((e) => getComputedStyle(e).backgroundColor)).toBe('rgba(0, 0, 0, 0)');
    // The status bar sums the selection.
    await select(page, 'B2:C2');
    await expect(page.getByTestId('status-sum')).toHaveText('Sum: 5.5');
    expect(errors).toEqual([]);
  });

  test('the cell context menu and Paste Special', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await typeIn(page, 'E2', '=B2+C2');
    await clickCell(page, 'E2', { button: 'right' });
    for (const id of ['ctx-cut', 'ctx-copy', 'ctx-insert', 'ctx-delete', 'ctx-sort', 'ctx-note', 'ctx-format-cells']) await expect(page.getByTestId(id)).toBeVisible();
    // Nothing has been copied yet.
    await expect(page.getByTestId('ctx-paste-special')).toBeDisabled();
    await page.getByTestId('ctx-copy').click();
    // Values only: the result of the formula, not the formula.
    await clickCell(page, 'F2', { button: 'right' });
    await page.getByTestId('ctx-paste-special').click();
    const dialog = page.getByTestId('paste-special');
    await dialog.getByTestId('paste-what-values').check();
    await dialog.getByTestId('paste-special-ok').click();
    await expect(dialog).toHaveCount(0);
    expect(await input(page, 'F2')).toBe('270');
    expect(await input(page, 'E2')).toBe('=B2+C2');
    expect(errors).toEqual([]);
  });

  test('every ribbon tab shows its groups; Home formats cells', async ({ page }) => {
    const errors = await openApp(page);
    for (const id of RIBBON_TABS) {
      await page.getByTestId(`tab-${id}`).click();
      await expect(page.getByTestId(`ribbon-${id}`)).toBeVisible();
    }
    await page.getByTestId('tab-home').click();
    await setData(page, SALES);
    await select(page, 'B2:D4');
    await page.getByTestId('rb-bold').click();
    expect(await docEval(page, '(d) => d.styleAt(1, 1).font.bold')).toBe(true);
    // Comma Style is an accounting format: padded for alignment.
    await page.getByTestId('rb-comma-style').click();
    expect(await shown(page, 'B2')).toBe(' 120.00 ');
    await page.getByTestId('rb-decimal-decrease').click();
    expect(await shown(page, 'B2')).toBe(' 120.0 ');
    await page.getByTestId('rb-font-size').fill('14');
    await page.getByTestId('rb-font-size').press('Enter');
    expect(await docEval(page, '(d) => d.styleAt(3, 3).font.size')).toBe(14);
    // Format Cells from the keyboard.
    await select(page, 'B2');
    await page.keyboard.press('Control+1');
    const dialog = page.getByTestId('format-cells');
    await expect(dialog).toBeVisible();
    await dialog.getByRole('option', { name: 'Percentage' }).click();
    await dialog.getByTestId('format-cells-ok').click();
    // The decimal places carry over from the comma style.
    expect(await shown(page, 'B2')).toBe('12000.0%');
    await page.keyboard.press('Control+z');
    expect(await shown(page, 'B2')).toBe(' 120.0 ');
    expect(errors).toEqual([]);
  });

  test('inserts a chart, changes it from Chart Design and saves it as a chart part', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await select(page, 'A1:D4');
    await page.getByTestId('tab-insert').click();
    await page.getByTestId('rb-insert-chart').click();
    const dialog = page.getByTestId('insert-chart');
    await dialog.getByTestId('chart-cat-line').click();
    await dialog.getByTestId('chart-sub-line-markers').click();
    await expect(dialog.getByTestId('chart-preview-0')).toBeVisible();
    await dialog.getByTestId('insert-chart-ok').click();
    const chart = page.getByTestId('object-chart');
    await expect(chart).toHaveCount(1);
    await expect(chart).toContainText('North');
    // The selected chart brings the Chart Design tab.
    await expect(page.getByTestId('tab-chart-design')).toHaveAttribute('aria-selected', 'true');
    await page.getByTestId('rb-switch-row-column').click();
    expect(await docEval(page, '(d) => d.sheet.charts[0].seriesIn')).toBe('columns');
    await expect(chart).toContainText('Q1');
    // The chart follows the data.
    await page.evaluate(() => (window as unknown as { spreadsheetEditor: { setValue(r: string, v: number): void } }).spreadsheetEditor.setValue('B2', 999));
    await expect(chart).toContainText('999');
    const { name, bytes } = await downloadOf(page, () => page.keyboard.press('Control+s'));
    expect(name).toBe('Book1.xlsx');
    const files = parts(bytes);
    expect(files['xl/charts/chart1.xml']).toContain('<c:lineChart>');
    expect(files['xl/drawings/drawing1.xml']).toContain('<c:chart');
    expect(errors).toEqual([]);
  });

  test('Data tab: sort, remove duplicates, subtotals and goal seek', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, [['Region', 'Sales'], ['West', 5], ['East', 10], ['West', 15], ['East', 10]]);
    await select(page, 'A1:B5');
    await page.getByTestId('tab-data').click();
    await page.getByTestId('rb-custom-sort').click();
    const sort = page.getByTestId('sort-dialog');
    await sort.getByTestId('sort-by-0').selectOption({ label: 'Region' });
    await sort.getByTestId('sort-add-level').click();
    await sort.getByTestId('sort-by-1').selectOption({ label: 'Sales' });
    await sort.getByTestId('sort-order-1').selectOption({ label: 'Largest to Smallest' });
    await sort.getByTestId('sort-dialog-ok').click();
    expect(await docEval(page, '(d) => [1,2,3,4].map((r) => d.displayText(r, 0) + d.displayText(r, 1))')).toEqual(['East10', 'East10', 'West15', 'West5']);
    // Remove Duplicates over both columns.
    await page.getByTestId('rb-remove-duplicates').click();
    await page.getByTestId('remove-duplicates-ok').click();
    await expect(page.getByTestId('toast')).toContainText('1 duplicate');
    expect(await shown(page, 'A5')).toBe('');
    // Subtotals at each change in Region.
    await select(page, 'A1:B4');
    await page.getByTestId('rb-subtotal').click();
    await page.getByTestId('subtotal-dialog-ok').click();
    expect(await shown(page, 'A3')).toBe('East Total');
    expect(await shown(page, 'B3')).toBe('10');
    expect(await input(page, 'B7')).toBe('=SUBTOTAL(9,B2:B6)');
    await expect(page.getByTestId('grid')).toBeVisible();
    // Goal Seek: make B7 reach 40 by changing B2.
    await page.getByTestId('rb-what-if').click();
    await page.getByTestId('goal-seek').click();
    const seek = page.getByTestId('goal-seek');
    await seek.getByTestId('goal-seek-set').fill('B7');
    await seek.getByTestId('goal-seek-to').fill('40');
    await seek.getByTestId('goal-seek-by').fill('B2');
    await seek.getByTestId('goal-seek-ok').click();
    await expect(page.getByTestId('goal-seek-status')).toContainText('found a solution');
    await page.getByTestId('goal-seek-status-ok').click();
    expect(await shown(page, 'B2')).toBe('20');
    expect(errors).toEqual([]);
  });

  test('File backstage: export CSV and a template, print preview and PDF', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await page.getByTestId('file-button').click();
    await expect(page.getByTestId('backstage')).toBeVisible();
    await page.getByTestId('backstage-export').click();
    await page.getByTestId('export-csv').click();
    const csv = await downloadOf(page, () => page.getByTestId('export-file').click());
    expect(csv.name).toBe('Book1.csv');
    expect(new TextDecoder().decode(csv.bytes)).toContain('North,120,150,90');
    await page.getByTestId('file-button').click();
    await page.getByTestId('backstage-export').click();
    await page.getByTestId('export-excel').click();
    await page.getByTestId('export-type').selectOption('xltx');
    const template = await downloadOf(page, () => page.getByTestId('export-file').click());
    expect(template.name).toBe('Book1.xltx');
    expect(parts(template.bytes)['[Content_Types].xml']).toContain('spreadsheetml.template.main+xml');
    // Print: a live preview and a PDF.
    await page.getByTestId('file-button').click();
    await page.getByTestId('backstage-print').click();
    await expect(page.getByTestId('print-page').locator('svg text').first()).toBeVisible();
    await expect(page.getByTestId('print-page-count')).toHaveText('of 1');
    const pdf = await downloadOf(page, () => page.getByTestId('print-download').click());
    expect(new TextDecoder('latin1').decode(pdf.bytes.subarray(0, 8))).toBe('%PDF-1.4');
    await page.getByTestId('backstage-back').click();
    await expect(page.getByTestId('backstage')).toHaveCount(0);
    expect(errors).toEqual([]);
  });

  test('the page API opens a workbook and switches the layout and theme', async ({ page }) => {
    const errors = await openApp(page);
    const bytes = [...readFileSync('tests/fixtures/kitchen-sink.xlsx')];
    await page.evaluate(async (data) => {
      const ed = (window as unknown as { spreadsheetEditor: { open(f: Blob, name?: string): Promise<void> } }).spreadsheetEditor;
      await ed.open(new Blob([new Uint8Array(data)]), 'kitchen-sink.xlsx');
    }, bytes);
    await expect(page.getByTestId('tab-Sales')).toBeVisible();
    expect(await shown(page, 'D8')).toBe('$6,501.25');
    await page.evaluate(() => {
      const ed = (window as unknown as { spreadsheetEditor: { setTheme(t: string): void; setLayout(l: string): void } }).spreadsheetEditor;
      ed.setTheme('dark');
      ed.setLayout('toolbar');
    });
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
    await expect(page.getByTestId('editor')).toHaveAttribute('data-layout', 'toolbar');
    expect(errors).toEqual([]);
  });

  test('conditional formatting and format as table', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await select(page, 'B2:D4');
    await page.getByTestId('rb-conditional-format').click();
    await page.getByTestId('cf-highlight').hover();
    await page.getByTestId('cf-greater').click();
    const preset = page.getByTestId('cf-preset');
    await preset.getByTestId('cf-value1').fill('100');
    await preset.getByTestId('cf-preset-ok').click();
    expect(await docEval(page, '(d) => d.sheet.condFormats.length')).toBe(1);
    // Format as Table opens the gallery, then the range dialog.
    await select(page, 'A1:D4');
    await page.getByTestId('rb-format-table').click();
    await page.getByTestId('table-style-TableStyleMedium2').click();
    await expect(page.getByTestId('create-table-range')).toHaveValue('$A$1:$D$4');
    await page.getByTestId('create-table-ok').click();
    expect(await docEval(page, '(d) => d.sheet.tables[0].style.name')).toBe('TableStyleMedium2');
    await expect(page.getByTestId('tab-table-design')).toBeVisible();
    expect(errors).toEqual([]);
  });

  test('a range picked on the sheet leaves the selection to the dialog; Escape cancels Goal Seek', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, [['Pair'], ['a,b'], ['c,d'], ['e,f']]);
    await select(page, 'A2:A4');
    await page.getByTestId('tab-data').click();
    await page.getByTestId('rb-text-to-columns').click();
    await page.getByTestId('text-to-columns-ok').click();
    await page.getByTestId('t2c-comma').check();
    await page.getByTestId('text-to-columns-ok').click();
    await page.getByTestId('t2c-destination-pick').click();
    await clickCell(page, 'F2');
    await page.getByTestId('range-picker-ok').click();
    await expect(page.getByTestId('t2c-destination')).toHaveValue('$F$2');
    await page.getByTestId('text-to-columns-ok').click();
    await expect(page.getByTestId('text-to-columns')).toHaveCount(0);
    expect(await docEval(page, '(d) => [1,2,3].map((r) => d.displayText(r, 5) + d.displayText(r, 6))')).toEqual(['ab', 'cd', 'ef']);
    expect(await docEval(page, '(d) => d.range')).toEqual({ r0: 1, c0: 0, r1: 3, c1: 0 });
    // Escape on Goal Seek Status is Cancel: the changing cell goes back.
    await setData(page, [[5]], 'H1');
    await typeIn(page, 'H2', '=H1*2');
    await page.getByTestId('rb-what-if').click();
    await page.getByTestId('goal-seek').click();
    await page.getByTestId('goal-seek-set').fill('H2');
    await page.getByTestId('goal-seek-to').fill('30');
    await page.getByTestId('goal-seek-by').fill('H1');
    await page.getByTestId('goal-seek-ok').click();
    await expect(page.getByTestId('goal-seek-status')).toContainText('found a solution');
    expect(await shown(page, 'H1')).toBe('15');
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('goal-seek-status')).toHaveCount(0);
    expect(await shown(page, 'H1')).toBe('5');
    expect(errors).toEqual([]);
  });

  test('dark mode, the toolbar layout and the command palette', async ({ page }) => {
    const errors = await openApp(page);
    await page.getByTestId('theme-toggle').click();
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
    await page.getByTestId('layout-toggle').check();
    await expect(page.getByTestId('editor')).toHaveAttribute('data-layout', 'toolbar');
    await expect(page.getByTestId('rb-bold')).toBeVisible();
    await page.getByTestId('layout-toggle').uncheck();
    await page.getByTestId('theme-toggle').click();
    await expect(page.locator('html')).not.toHaveAttribute('data-theme', 'dark');
    // Find a command by name.
    await page.keyboard.press('Alt+q');
    const palette = page.getByTestId('command-palette');
    await palette.getByTestId('palette-search').fill('freeze top row');
    await page.keyboard.press('Enter');
    // The palette closes first and runs the command a moment later.
    await expect.poll(() => docEval(page, '(d) => d.sheet.freeze')).toMatchObject({ rows: 1, cols: 0 });
    expect(errors).toEqual([]);
  });
});
