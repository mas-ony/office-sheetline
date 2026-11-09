/** Sparkline groups in the worksheet extension list (x14:sparklineGroups, Excel 2010). */
import { cellName, parseCellRef, parseSqref } from '../refs';
import { attr, childrenNamed, el, findAll, numAttr, textContent, txt, type XmlElement } from '../xml/xml';
import { colorAttrs, readColor } from './styles';
import { defaultSparklineColors, newSparklineGroup, type SparklineColors, type SparklineGroup } from '../model/sparkline';

export const SPARKLINE_EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';

const COLOR_KEYS: SparklineColors[] = ['colorSeries', 'colorNegative', 'colorAxis', 'colorMarkers', 'colorFirst', 'colorLast', 'colorHigh', 'colorLow'];

const flag = (e: XmlElement, name: string) => attr(e, name) === '1' || attr(e, name) === 'true';

/** Sparkline groups of a worksheet's extension list. */
export function readSparklineGroups(extLst: XmlElement | undefined): SparklineGroup[] {
  if (!extLst) return [];
  const out: SparklineGroup[] = [];
  for (const g of findAll(extLst, 'x14:sparklineGroup')) {
    const type = attr(g, 'type');
    const sparklines = findAll(g, 'x14:sparkline').flatMap((s) => {
      const f = findAll(s, 'xm:f')[0];
      const sq = findAll(s, 'xm:sqref')[0];
      const at = sq ? parseCellRef(textContent(sq).trim()) ?? parseSqref(textContent(sq).trim())[0] : null;
      if (!at) return [];
      const r = 'r' in at ? at.r : at.r0;
      const c = 'c' in at ? at.c : at.c0;
      return [{ r, c, ref: f ? textContent(f).trim() : '' }];
    });
    if (!sparklines.length) continue;
    const group = newSparklineGroup(type === 'column' ? 'column' : type === 'stacked' ? 'stacked' : 'line', sparklines);
    const colors = defaultSparklineColors();
    for (const k of COLOR_KEYS) {
      const spec = readColor(childrenNamed(g, `x14:${k}`)[0]);
      if (spec) colors[k] = spec;
    }
    const empty = attr(g, 'displayEmptyCellsAs');
    const axisType = (v: string | undefined) => (v === 'group' || v === 'custom' ? v : 'individual');
    out.push({
      ...group,
      colors,
      lineWeight: numAttr(g, 'lineWeight') ?? 0.75,
      markers: flag(g, 'markers'),
      high: flag(g, 'high'),
      low: flag(g, 'low'),
      first: flag(g, 'first'),
      last: flag(g, 'last'),
      negative: flag(g, 'negative'),
      displayXAxis: flag(g, 'displayXAxis'),
      displayHidden: flag(g, 'displayHidden'),
      displayEmptyCellsAs: empty === 'gap' || empty === 'span' ? empty : 'zero',
      minAxisType: axisType(attr(g, 'minAxisType')),
      maxAxisType: axisType(attr(g, 'maxAxisType')),
      manualMin: numAttr(g, 'manualMin'),
      manualMax: numAttr(g, 'manualMax'),
      rightToLeft: flag(g, 'rightToLeft'),
    });
  }
  return out;
}

/** The extension element holding a sheet's sparkline groups. */
export function sparklineExt(groups: SparklineGroup[]): XmlElement {
  const bool = (v: boolean) => (v ? 1 : undefined);
  return el('ext', { uri: SPARKLINE_EXT_URI, 'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main' }, [
    el('x14:sparklineGroups', { 'xmlns:xm': 'http://schemas.microsoft.com/office/excel/2006/main' }, groups.map((g) =>
      el('x14:sparklineGroup', {
        manualMax: g.maxAxisType === 'custom' ? g.manualMax : undefined,
        manualMin: g.minAxisType === 'custom' ? g.manualMin : undefined,
        lineWeight: g.lineWeight !== 0.75 ? g.lineWeight : undefined,
        type: g.type === 'line' ? undefined : g.type,
        displayEmptyCellsAs: g.displayEmptyCellsAs === 'zero' ? undefined : g.displayEmptyCellsAs,
        markers: bool(g.markers),
        high: bool(g.high),
        low: bool(g.low),
        first: bool(g.first),
        last: bool(g.last),
        negative: bool(g.negative),
        displayXAxis: bool(g.displayXAxis),
        displayHidden: bool(g.displayHidden),
        minAxisType: g.minAxisType === 'individual' ? undefined : g.minAxisType,
        maxAxisType: g.maxAxisType === 'individual' ? undefined : g.maxAxisType,
        rightToLeft: bool(g.rightToLeft),
      }, [
        ...COLOR_KEYS.map((k) => el(`x14:${k}`, colorAttrs(g.colors[k]))),
        el('x14:sparklines', {}, g.sparklines.map((s) => el('x14:sparkline', {}, [el('xm:f', {}, [txt(s.ref)]), el('xm:sqref', {}, [txt(cellName(s.r, s.c))])]))),
      ]),
    )),
  ]);
}
