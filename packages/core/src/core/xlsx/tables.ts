/** Table parts (xl/tables/tableN.xml): reading into SheetTable and writing them back. */
import { parseRange, rangeName } from '../refs';
import type { Sheet } from '../model/sheet';
import type { SheetTable, TableColumn, TotalsFunction } from '../model/table';
import { attr, boolAttr, child, childrenNamed, el, numAttr, textContent, txt, serializeXml, type XmlElement } from '../xml/xml';
import { MAIN_NS } from '../xml/namespaces';
import { readFilterColumns } from './import';
import { autoFilterXml } from './export';
import { formulaFromFile } from './formulaText';
import { formulaToFile } from './formulaText';
import { tokenize } from '../formula/tokenizer';
import { rewriteStructured } from '../formula/structured';
import type { OoxmlPackage, Relationship } from './package';
import { REL } from './package';

export const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

const TOTALS: TotalsFunction[] = ['none', 'sum', 'average', 'count', 'countNums', 'max', 'min', 'stdDev', 'var', 'custom'];

function decode(s: string): string {
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
}

function encode(s: string): string {
  // eslint-disable-next-line no-control-regex
  return s.replace(/[\u0000-\u001F]/g, (ch) => `_x${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);
}

/** Reads a <table> element. */
export function readTable(root: XmlElement, path: string): SheetTable | null {
  const ref = parseRange(attr(root, 'ref') ?? '');
  if (!ref) return null;
  const headerRowCount = numAttr(root, 'headerRowCount') ?? 1;
  const totalsRowCount = numAttr(root, 'totalsRowCount') ?? 0;
  const columns: TableColumn[] = childrenNamed(child(root, 'tableColumns'), 'tableColumn').map((tc, i): TableColumn => {
    const fn = attr(tc, 'totalsRowFunction') as TotalsFunction | undefined;
    const calc = child(tc, 'calculatedColumnFormula');
    const tot = child(tc, 'totalsRowFormula');
    return {
      id: numAttr(tc, 'id') ?? i + 1,
      name: decode(attr(tc, 'name') ?? `Column${i + 1}`),
      totalsFunction: fn && TOTALS.includes(fn) ? fn : undefined,
      totalsLabel: attr(tc, 'totalsRowLabel'),
      totalsFormula: tot ? formulaFromFile(textContent(tot)) : undefined,
      calculatedFormula: calc ? formulaFromFile(textContent(calc)) : undefined,
      raw: tc,
    };
  });
  const width = ref.c1 - ref.c0 + 1;
  while (columns.length < width) columns.push({ id: columns.length + 1, name: `Column${columns.length + 1}` });
  columns.length = width;
  const af = child(root, 'autoFilter');
  const info = child(root, 'tableStyleInfo');
  const name = attr(root, 'displayName') ?? attr(root, 'name') ?? 'Table1';
  return {
    id: numAttr(root, 'id') ?? 1,
    name,
    ref,
    headerRow: headerRowCount > 0,
    totalsRow: totalsRowCount > 0,
    columns,
    style: {
      name: info ? attr(info, 'name') ?? '' : '',
      showFirstColumn: boolAttr(info, 'showFirstColumn'),
      showLastColumn: boolAttr(info, 'showLastColumn'),
      showRowStripes: info ? boolAttr(info, 'showRowStripes') : false,
      showColumnStripes: boolAttr(info, 'showColumnStripes'),
    },
    autoFilter: !!af && headerRowCount > 0,
    filters: af ? readFilterColumns(af, ref.c0) : new Map(),
    autoExpand: true,
    comment: attr(root, 'comment'),
    raw: root,
    sourcePath: path,
  };
}

/** Structured references in the form stored in files: [@Col] becomes Table1[[#This Row],[Col]]. */
export function structuredToFile(formula: string, table: string | null): string {
  if (!formula.includes('[')) return formula;
  try {
    return rewriteStructured(formula, tokenize(formula), (text, t) => {
      if (text === t || !text.includes('[')) return null;
      const name = t || table;
      if (!name) return null;
      let body = text.slice(t.length);
      // [@Col] / [@[Col]] / [@]
      const at = /^\[@(.*)\]$/s.exec(body);
      if (at) {
        const col = at[1].trim();
        if (!col) return `${name}[#This Row]`;
        const inner = col.startsWith('[') ? col : `[${col}]`;
        return `${name}[[#This Row],${inner}]`;
      }
      if (!t) body = body.trim();
      return t ? null : `${name}${body}`;
    });
  } catch {
    return formula;
  }
}

function tableXml(t: SheetTable): string {
  const kids: XmlElement[] = [];
  const filterRef = { ...t.ref, r1: t.ref.r1 - (t.totalsRow ? 1 : 0) };
  if (t.autoFilter && t.headerRow) {
    const af = autoFilterXml({ ref: filterRef, columns: t.filters });
    kids.push(af);
  }
  kids.push(el('tableColumns', { count: t.columns.length }, t.columns.map((col) => {
    const inner: XmlElement[] = [];
    if (col.calculatedFormula) inner.push(el('calculatedColumnFormula', {}, [txt(formulaToFile(structuredToFile(col.calculatedFormula, t.name)))]));
    if (col.totalsFunction === 'custom' && col.totalsFormula) inner.push(el('totalsRowFormula', {}, [txt(formulaToFile(structuredToFile(col.totalsFormula, t.name)))]));
    return el('tableColumn', {
      id: col.id,
      name: encode(col.name),
      totalsRowFunction: t.totalsRow && col.totalsFunction && col.totalsFunction !== 'none' ? col.totalsFunction : undefined,
      totalsRowLabel: t.totalsRow && col.totalsLabel !== undefined && (!col.totalsFunction || col.totalsFunction === 'none') ? col.totalsLabel : undefined,
    }, inner);
  })));
  if (t.style.name || t.style.showRowStripes || t.style.showColumnStripes) {
    kids.push(el('tableStyleInfo', {
      name: t.style.name || undefined,
      showFirstColumn: t.style.showFirstColumn ? 1 : 0,
      showLastColumn: t.style.showLastColumn ? 1 : 0,
      showRowStripes: t.style.showRowStripes ? 1 : 0,
      showColumnStripes: t.style.showColumnStripes ? 1 : 0,
    }));
  }
  const root = el('table', {
    id: t.id,
    name: t.name,
    displayName: t.name,
    ref: rangeName(t.ref),
    headerRowCount: t.headerRow ? undefined : 0,
    totalsRowCount: t.totalsRow ? 1 : undefined,
    totalsRowShown: t.totalsRow ? undefined : 0,
    comment: t.comment,
  }, kids);
  return serializeXml(root, { namespaces: { '': MAIN_NS } });
}

/**
 * Writes the tables of a sheet as table parts and returns the <tableParts> element (or null).
 * Relationships are added to `rels`; parts of tables that no longer exist are removed by the caller.
 */
export function writeTables(sheet: Sheet, sheetPath: string, rels: Relationship[], pkg: OoxmlPackage, used: Set<string>, nextRelId: () => string): XmlElement | null {
  if (!sheet.tables.length) return null;
  const parts: XmlElement[] = [];
  for (const t of sheet.tables) {
    let path = t.sourcePath && !used.has(t.sourcePath) ? t.sourcePath : undefined;
    if (!path) {
      let n = 1;
      while (pkg.has(`xl/tables/table${n}.xml`) || used.has(`xl/tables/table${n}.xml`)) n++;
      path = `xl/tables/table${n}.xml`;
    }
    used.add(path);
    pkg.setText(path, tableXml(t), CT_TABLE);
    const rid = nextRelId();
    rels.push({ id: rid, type: REL.table, target: pkg.relativeTarget(sheetPath, path) });
    parts.push(el('tablePart', { 'r:id': rid }));
  }
  return el('tableParts', { count: parts.length }, parts);
}
