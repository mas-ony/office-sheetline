import { expect, test } from '@playwright/test';
import { SALES, cellPoint, docEval, downloadOf, input, openApp, parts, select, setData, shown, tapCell } from './helpers';
import { buildXls } from '../../packages/core/tests/unit/binaryWorkbooks';

test.describe('phone layout', () => {
  test('fits the screen, taps select and double taps edit with the key bar', async ({ page }) => {
    const errors = await openApp(page);
    await expect(page.getByTestId('editor')).toHaveClass(/phone/);
    await expect(page.getByTestId('rb-group-font')).toBeVisible();
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1)).toBe(true);
    await tapCell(page, 'B3');
    await expect(page.getByTestId('name-box')).toHaveValue('B3');
    // Double tap opens the editor with the on-screen keyboard and the key bar.
    const p = await cellPoint(page, 'B3');
    await page.touchscreen.tap(p.x, p.y);
    await page.waitForTimeout(60);
    await page.touchscreen.tap(p.x, p.y);
    await expect(page.getByTestId('cell-editor')).toBeFocused();
    await expect(page.getByTestId('keybar')).toBeVisible();
    await page.keyboard.type('42');
    await page.keyboard.press('Enter');
    expect(await shown(page, 'B3')).toBe('42');
    await expect(page.getByTestId('keybar')).toHaveCount(0);
    expect(errors).toEqual([]);
  });

  test('the key bar types formula characters and functions; Done commits and Cancel discards', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, [[1], [2], [3]]);
    const bar = page.getByTestId('keybar');
    const editor = page.getByTestId('cell-editor');
    const doubleTap = async (ref: string) => {
      const p = await cellPoint(page, ref);
      await page.touchscreen.tap(p.x, p.y);
      await page.waitForTimeout(60);
      await page.touchscreen.tap(p.x, p.y);
      await expect(bar).toBeVisible();
    };
    const dialog = page.getByTestId('insert-function');
    const pickFunction = async (name: string) => {
      await bar.getByTestId('key-functions').tap();
      await dialog.getByTestId('insert-function-search').fill(name);
      await dialog.getByRole('option', { name, exact: true }).tap();
      await dialog.getByTestId('insert-function-ok').tap();
      await expect(dialog).toHaveCount(0);
    };
    // Characters and functions from the bar go in at the caret; Done commits and moves down.
    await doubleTap('B1');
    await bar.getByRole('button', { name: 'Insert =' }).tap();
    await page.keyboard.type('A1');
    await bar.getByRole('button', { name: 'Insert +' }).tap();
    await pickFunction('SUM');
    await expect(editor).toHaveValue('=A1+SUM(');
    await expect(editor).toBeFocused();
    await page.keyboard.type('A2');
    await bar.getByRole('button', { name: 'Insert :' }).tap();
    await page.keyboard.type('A3');
    await bar.getByRole('button', { name: 'Insert )' }).tap();
    await expect(editor).toHaveValue('=A1+SUM(A2:A3)');
    await bar.getByTestId('key-done').tap();
    expect(await input(page, 'B1')).toBe('=A1+SUM(A2:A3)');
    expect(await shown(page, 'B1')).toBe('6');
    await expect(page.getByTestId('name-box')).toHaveValue('B2');
    await expect(bar).toHaveCount(0);
    // In a cell without a formula, Functions starts one; closing the list keeps the formula.
    await doubleTap('B2');
    await pickFunction('SUM');
    await expect(editor).toHaveValue('=SUM(');
    await page.keyboard.type('A1');
    await bar.getByRole('button', { name: 'Insert :' }).tap();
    await page.keyboard.type('A2');
    await bar.getByTestId('key-functions').tap();
    await dialog.getByTestId('insert-function-cancel').tap();
    await expect(dialog).toHaveCount(0);
    await expect(editor).toBeFocused();
    await expect(editor).toHaveValue('=SUM(A1:A2');
    await bar.getByRole('button', { name: 'Insert )' }).tap();
    await bar.getByTestId('key-done').tap();
    expect(await input(page, 'B2')).toBe('=SUM(A1:A2)');
    expect(await shown(page, 'B2')).toBe('3');
    // Cancel leaves the cell as it was.
    await doubleTap('A1');
    await page.keyboard.type('99');
    await expect(editor).toHaveValue('199');
    await bar.getByTestId('key-cancel').tap();
    await expect(bar).toHaveCount(0);
    expect(await input(page, 'A1')).toBe('1');
    expect(errors).toEqual([]);
  });

  test('long press opens the cell menu as a bottom sheet', async ({ page }) => {
    await openApp(page);
    const p = await cellPoint(page, 'C4');
    const cdp = await page.context().newCDPSession(page);
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: p.x, y: p.y }] });
    await page.waitForTimeout(700);
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
    const sheet = page.locator('.sh-bottom-sheet');
    await expect(sheet.getByTestId('ctx-format-cells')).toBeVisible();
    await sheet.getByTestId('ctx-format-cells').tap();
    await expect(page.getByTestId('format-cells')).toBeVisible();
  });

  test('opens an Excel 97-2003 file', async ({ page }) => {
    await openApp(page);
    await page.getByTestId('open-input').setInputFiles({ name: 'Budget.xls', mimeType: 'application/vnd.ms-excel', buffer: Buffer.from(buildXls()) });
    await page.waitForFunction(() => (window as unknown as { spreadsheetEditor?: { doc: { fileName: string } } }).spreadsheetEditor?.doc.fileName === 'Budget.xlsx');
    expect(await shown(page, 'A1')).toBe('Region');
    expect(await shown(page, 'B2')).toBe('19.75');
  });

  test('ribbon groups open as bottom sheets', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await select(page, 'B2:D4');
    await page.getByTestId('rb-group-font').tap();
    const panel = page.getByTestId('rb-group-font-panel');
    await expect(panel).toBeVisible();
    const box = (await panel.boundingBox())!;
    const view = page.viewportSize()!;
    expect(box.y + box.height).toBeGreaterThan(view.height - 4);
    expect(box.width).toBeGreaterThan(view.width - 4);
    await panel.getByTestId('rb-bold').tap();
    expect(await docEval(page, '(d) => d.styleAt(1, 1).font.bold')).toBe(true);
    expect(errors).toEqual([]);
  });

  test('inserts a chart under the data, inside the screen width', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await select(page, 'A1:D4');
    await page.getByTestId('tab-insert').tap();
    await page.getByTestId('rb-group-charts').tap();
    await page.getByTestId('rb-group-charts-panel').getByTestId('rb-insert-chart').tap();
    const dialog = page.getByTestId('insert-chart');
    await dialog.getByTestId('chart-cat-pie').tap();
    await dialog.getByTestId('insert-chart-ok').tap();
    const chart = page.getByTestId('object-chart');
    await expect(chart).toBeVisible();
    // One empty row between the data and the chart.
    expect(await docEval(page, '(d) => d.sheet.charts[0].from')).toMatchObject({ c: 0, r: 5 });
    const box = (await chart.boundingBox())!;
    expect(box.x + box.width).toBeLessThanOrEqual(page.viewportSize()!.width + 1);
    expect(errors).toEqual([]);
  });

  test('the File backstage fills the screen and exports the workbook', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await page.getByTestId('file-button').tap();
    const backstage = page.getByTestId('backstage');
    await expect(backstage).toBeVisible();
    expect((await backstage.boundingBox())!.width).toBeGreaterThan(page.viewportSize()!.width - 2);
    await page.getByTestId('backstage-export').tap();
    const { name, bytes } = await downloadOf(page, () => page.getByTestId('export-file').tap());
    expect(name).toBe('Book1.xlsx');
    expect(parts(bytes)['xl/worksheets/sheet1.xml']).toContain('<v>150</v>');
    // Print preview pages fit the screen.
    await page.getByTestId('file-button').tap();
    await page.getByTestId('backstage-print').tap();
    const preview = page.getByTestId('print-page').first();
    await expect(preview).toBeVisible();
    expect((await preview.boundingBox())!.width).toBeLessThanOrEqual(page.viewportSize()!.width);
    expect(errors).toEqual([]);
  });
});
