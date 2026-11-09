import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import '../../src/core/doc';
import { SpreadsheetDocument } from '../../src/core/document';
import { chartFromRange, readChartData } from '../../src/core/chartData';
import type { ChartType, SheetChart } from '../../src/core/model/drawing';
import { parseRange } from '../../src/core/refs';
import { readXlsx } from '../../src/core/xlsx/import';
import { parseXml } from '../../src/core/xml/xml';
import { COLOR_SCHEMES, FONT_SCHEMES } from '../../src/core/themes';
import { chartSvg } from '../../src/grid/chartRender';

const rect = { x: 500, y: 0, width: 400, height: 300 };

function roundTrip(d: SpreadsheetDocument, opts?: { template?: boolean; macros?: boolean }) {
  const bytes = d.save(true, opts);
  return { bytes, files: unzipSync(bytes), wb: readXlsx(bytes) };
}

/** The file without the editor's own chart settings, as a chart that Excel wrote reads. */
function withoutEditorSettings(bytes: Uint8Array): Uint8Array {
  const files = unzipSync(bytes);
  for (const name of Object.keys(files)) {
    if (/^xl\/charts\/chart\d+\.xml$/.test(name)) files[name] = strToU8(strFromU8(files[name]).replace(/<c:ext uri="\{5D0E8E86[^"]*"[\s\S]*?<\/c:ext>/, ''));
  }
  return zipSync(files);
}

function addChart(d: SpreadsheetDocument, type: ChartType, ref: string, patch: Partial<SheetChart> = {}) {
  const def = chartFromRange(d, d.sheet, parseRange(ref)!, type);
  return d.addChart({ ...def, name: 'Chart 1', ...patch, series: patch.series ?? def.series }, rect)!;
}

describe('sparklines', () => {
  it('creates one sparkline per data row and saves the group as an x14 extension', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[1, 5, 3, 8], [4, -2, 6, 1]]);
    expect(d.addSparklines('line', { sheet: d.sheet, range: parseRange('A1:D2')! }, parseRange('E1:E2')!)).toBeNull();
    const group = d.sheet.sparklineGroups[0];
    expect(group.sparklines.map((s) => [s.r, s.c, s.ref])).toEqual([[0, 4, 'Sheet1!$A$1:$D$1'], [1, 4, 'Sheet1!$A$2:$D$2']]);
    d.updateSparklineGroup(group.id, { markers: true, high: true, lineWeight: 1.5, displayEmptyCellsAs: 'zero' });
    const { files, wb } = roundTrip(d);
    const xml = strFromU8(files['xl/worksheets/sheet1.xml']);
    expect(xml).toContain('<x14:sparklineGroups');
    expect(xml).toContain('<xm:f>Sheet1!$A$1:$D$1</xm:f>');
    expect(xml).toContain('<xm:sqref>E1</xm:sqref>');
    expect(() => parseXml(xml)).not.toThrow();
    const back = wb.sheets[0].sparklineGroups;
    expect(back).toHaveLength(1);
    expect(back[0]).toMatchObject({ type: 'line', markers: true, high: true, lineWeight: 1.5, displayEmptyCellsAs: 'zero' });
    expect(back[0].sparklines.map((s) => [s.r, s.c])).toEqual([[0, 4], [1, 4]]);
  });

  it('checks the location against the data and replaces older sparklines in a cell', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[1, 2], [3, 4]]);
    expect(d.addSparklines('column', { sheet: d.sheet, range: parseRange('A1:B2')! }, parseRange('C1:D2')!)).toMatch(/single column or row|same column or row/);
    expect(d.addSparklines('column', { sheet: d.sheet, range: parseRange('A1:B2')! }, parseRange('C1')!)).toMatch(/single row or a single column/);
    d.addSparklines('column', { sheet: d.sheet, range: parseRange('A1:B2')! }, parseRange('C1:C2')!);
    d.addSparklines('stacked', { sheet: d.sheet, range: parseRange('A1:B1')! }, parseRange('C1')!);
    expect(d.sparklineAt(0, 2)?.group.type).toBe('stacked');
    expect(d.sparklineAt(1, 2)?.group.type).toBe('column');
    expect(d.sheet.sparklineGroups).toHaveLength(2);
  });

  it('groups, ungroups and clears sparklines', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[1, 2], [3, 4], [5, 6]]);
    d.addSparklines('line', { sheet: d.sheet, range: parseRange('A1:B3')! }, parseRange('C1:C3')!);
    d.ungroupSparklines([parseRange('C1:C3')!]);
    expect(d.sheet.sparklineGroups).toHaveLength(3);
    d.groupSparklines([parseRange('C1:C2')!]);
    expect(d.sheet.sparklineGroups.map((g) => g.sparklines.length).sort()).toEqual([1, 2]);
    d.clearSparklines([parseRange('C1')!], true);
    expect(d.sheet.sparklineGroups.map((g) => g.sparklines.map((s) => s.r))).toEqual([[2]]);
    d.undo();
    d.clearSparklines([parseRange('C1')!], false);
    expect(d.sheet.sparklineGroups.flatMap((g) => g.sparklines.map((s) => s.r)).sort()).toEqual([1, 2]);
  });

  it('moves with inserted rows and follows its data', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[1, 2, 3]]);
    d.addSparklines('line', { sheet: d.sheet, range: parseRange('A1:C1')! }, parseRange('D1')!);
    d.insertRows(0, 2);
    const s = d.sheet.sparklineGroups[0].sparklines[0];
    expect([s.r, s.c, s.ref]).toEqual([2, 3, 'Sheet1!$A$3:$C$3']);
    d.deleteRows(2, 1);
    expect(d.sheet.sparklineGroups).toHaveLength(0);
  });
});

describe('themes', () => {
  it('writes the chosen colours and fonts and reads them back', () => {
    const d = SpreadsheetDocument.blank();
    const scheme = COLOR_SCHEMES.find((c) => c.name !== 'Office')!;
    const fonts = FONT_SCHEMES.find((f) => f.name === 'Office 2007')!;
    d.applyTheme({ name: scheme.name, colors: { ...scheme.colors }, majorFont: fonts.major, minorFont: fonts.minor });
    const { files, wb } = roundTrip(d);
    const xml = strFromU8(files['xl/theme/theme1.xml']);
    expect(xml).toContain(`<a:accent1><a:srgbClr val="${scheme.colors.accent1}"/></a:accent1>`);
    expect(xml).toContain('<a:latin typeface="Cambria"/>');
    expect(wb.theme.colors.accent1).toBe(scheme.colors.accent1);
    expect([wb.theme.majorFont, wb.theme.minorFont]).toEqual(['Cambria', 'Calibri']);
    expect(wb.themeName).toBe(scheme.name);
    d.undo();
    expect(d.wb.themeName).toBe('Office');
    expect(readXlsx(d.save()).themeName).toBe('Office');
  });
});

describe('templates and macros', () => {
  it('saves a template with the template content type', () => {
    const d = SpreadsheetDocument.blank();
    const types = strFromU8(roundTrip(d, { template: true }).files['[Content_Types].xml']);
    expect(types).toContain('application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml');
    const macro = strFromU8(roundTrip(d, { template: true, macros: true }).files['[Content_Types].xml']);
    expect(macro).toContain('application/vnd.ms-excel.template.macroEnabled.main+xml');
    const plain = strFromU8(roundTrip(d).files['[Content_Types].xml']);
    expect(plain).toContain('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
  });
});

function stockDoc(volume: boolean) {
  const d = SpreadsheetDocument.blank();
  const rows: (string | number)[][] = [
    ['Date', 'Volume', 'Open', 'High', 'Low', 'Close'],
    ['2024-01-02', 1200, 10, 12, 9, 11],
    ['2024-01-03', 900, 11, 13, 10, 12],
    ['2024-01-04', 1500, 12, 12.5, 10.5, 11],
  ];
  d.setValues(0, 0, volume ? rows : rows.map((r) => [r[0], ...r.slice(2)]));
  return d;
}

describe('combo, stock, sunburst and scatter charts in files', () => {
  it('writes a column-line combo with a secondary axis', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['Month', 'Sales', 'Cost', 'Margin'], ['Jan', 10, 6, 0.4], ['Feb', 12, 7, 0.42], ['Mar', 9, 6, 0.33]]);
    const def = chartFromRange(d, d.sheet, parseRange('A1:D4')!, 'combo');
    const series = def.series.map((s, i) => ({ ...s, kind: i === 2 ? ('line' as const) : ('column' as const), secondaryAxis: i === 2 || undefined }));
    addChart(d, 'combo', 'A1:D4', { series });
    const { bytes, files } = roundTrip(d);
    const xml = strFromU8(files['xl/charts/chart1.xml']);
    expect(xml).toContain('<c:barChart>');
    expect(xml).toContain('<c:lineChart>');
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2);
    // Read like a file from Excel: the chart groups say what it is.
    const back = readXlsx(withoutEditorSettings(bytes)).sheets[0].charts[0];
    expect(back.type).toBe('combo');
    expect(back.series.map((s) => [s.kind, !!s.secondaryAxis])).toEqual([['column', false], ['column', false], ['line', true]]);
    const svg = chartSvg(back, readChartData(d, back), 400, 300);
    expect(svg).not.toContain('NaN');
  });

  it('writes stock charts with and without volume', () => {
    for (const kind of ['ohlc', 'vohlc'] as const) {
      const volume = kind === 'vohlc';
      const d = stockDoc(volume);
      const chart = addChart(d, 'stock', volume ? 'A1:F4' : 'A1:E4', { stockKind: kind });
      expect(chart.series, kind).toHaveLength(volume ? 5 : 4);
      const { bytes, files } = roundTrip(d);
      const xml = strFromU8(files['xl/charts/chart1.xml']);
      expect(xml, kind).toContain('<c:stockChart>');
      expect(xml, kind).toContain('<c:upDownBars>');
      expect(xml.includes('<c:barChart>'), kind).toBe(volume);
      const back = readXlsx(withoutEditorSettings(bytes)).sheets[0].charts[0];
      expect([back.type, back.stockKind], kind).toEqual(['stock', kind]);
      expect(back.series, kind).toHaveLength(volume ? 5 : 4);
      expect(chartSvg(back, readChartData(d, back), 400, 300), kind).not.toContain('NaN');
    }
  });

  it('writes hierarchical sunburst labels, innermost level first', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['Region', 'Rep', 'Sales'], ['East', 'Ann', 10], [null, 'Bob', 20], ['West', 'Cid', 5]]);
    const chart = addChart(d, 'sunburst', 'A1:C4');
    expect(chart.categories).toBe('Sheet1!$A$2:$B$4');
    expect(readChartData(d, chart).levels).toEqual([['East', 'Ann'], ['East', 'Bob'], ['West', 'Cid']]);
    const { files, wb } = roundTrip(d);
    const part = Object.keys(files).find((f) => /^xl\/charts\/chartEx\d+\.xml$/.test(f))!;
    const xml = strFromU8(files[part]);
    expect(xml).toContain('<cx:f>Sheet1!$A$2:$B$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">Ann</cx:pt>');
    expect(xml).toContain('<cx:lvl ptCount="3"><cx:pt idx="0">East</cx:pt><cx:pt idx="2">West</cx:pt></cx:lvl>');
    expect(wb.sheets[0].charts[0].type).toBe('sunburst');
  });

  it('keeps the scatter style and error bars', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['x', 'y'], [1, 3], [2, 5], [3, 4]]);
    addChart(d, 'scatter', 'A1:B4', { scatterStyle: 'smooth', errorBars: 'percent' });
    const { bytes, files } = roundTrip(d);
    const xml = strFromU8(files['xl/charts/chart1.xml']);
    expect(xml).toContain('<c:smooth val="1"/>');
    expect(xml).toContain('<c:errValType val="percentage"/>');
    const back = readXlsx(withoutEditorSettings(bytes)).sheets[0].charts[0];
    expect([back.type, back.scatterStyle, back.errorBars]).toEqual(['scatter', 'smooth', 'percent']);
  });
});

describe('object names, alternative text and visibility', () => {
  it('round-trips alt text and hidden objects and keeps them out of print', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['a', 'b'], [1, 2]]);
    const chart = addChart(d, 'column', 'A1:B2');
    const box = d.addShape({ name: 'Note', preset: 'rect', fill: '#FFFFFF', line: { color: '#000000', width: 1 }, fontColor: '#000000', text: 'Hi', textBox: true, bold: true, fontSize: 14 }, { x: 10, y: 100, width: 120, height: 40 })!;
    d.setAltText(chart.id, 'Sales chart', 'Sales by month');
    d.setObjectProps(box.id, { name: 'Callout', hidden: true });
    const { files, wb } = roundTrip(d);
    const drawing = strFromU8(files['xl/drawings/drawing1.xml']);
    expect(drawing).toContain('descr="Sales by month"');
    expect(drawing).toContain('title="Sales chart"');
    expect(drawing).toContain('hidden="1"');
    expect(drawing).toContain('txBox="1"');
    const back = wb.sheets[0];
    expect(back.charts[0]).toMatchObject({ altTitle: 'Sales chart', descr: 'Sales by month' });
    expect(back.shapes[0]).toMatchObject({ name: 'Callout', hidden: true, textBox: true, bold: true, fontSize: 14, text: 'Hi' });
  });
});

describe('defined names', () => {
  it('saves comments and sheet scope', () => {
    const d = SpreadsheetDocument.blank();
    d.addSheet('Two');
    expect(d.defineName({ name: 'Rate', formula: '0.05', comment: 'Interest' })).toBeNull();
    expect(d.defineName({ name: 'Local', formula: 'Two!$A$1', localSheet: 1 })).toBeNull();
    expect(d.defineName({ name: 'A1', formula: '1' })).not.toBeNull();
    const names = readXlsx(d.save()).names;
    expect(names.find((n) => n.name === 'Rate')).toMatchObject({ formula: '0.05', comment: 'Interest' });
    expect(names.find((n) => n.name === 'Local')).toMatchObject({ localSheet: 1 });
    d.removeName('Rate');
    expect(readXlsx(d.save()).names.some((n) => n.name === 'Rate')).toBe(false);
  });
});
