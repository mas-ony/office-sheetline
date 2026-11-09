/**
 * Conditional formatting commands: rules are added with the highest priority (like Excel), edited,
 * reordered, deleted and cleared from the selection, the sheet or a table. Each rule is kept in its
 * own <conditionalFormatting> block so it can have its own "Applies to" ranges.
 */
import { SpreadsheetDocument } from '../document';
import type { Sheet, ConditionalFormat } from '../model/sheet';
import { cloneXml, el, numAttr, type XmlElement } from '../xml/xml';
import { formatToDxf, newCfId, normalizeCondFormats, ruleFromXml, ruleToXml, type CfRule, type RuleEntry } from '../condFormats';
import { intersect, rangesIntersect, sqrefString, type Range } from '../refs';

declare module '../document' {
  interface SpreadsheetDocument {
    /** Rules of a sheet in priority order, optionally only those touching the given ranges. */
    cfRules(sheet?: Sheet, touching?: Range[]): RuleEntry[];
    addCondFormat(rule: CfRule, ranges?: Range[], opts?: { stopIfTrue?: boolean }, sheet?: Sheet): string;
    /**
     * Applies the Rules Manager's edits in one step: the listed rules in their new order (top =
     * highest priority); rules of the sheet not listed are deleted when `replaceAll` is set.
     */
    applyRuleList(list: { id?: string; rule: CfRule | null; raw?: XmlElement; ranges: Range[]; stopIfTrue: boolean }[], scope: Range[] | null, sheet?: Sheet): void;
    deleteCondFormat(id: string, sheet?: Sheet): void;
    clearCondFormats(scope: 'selection' | 'sheet' | 'table', sheet?: Sheet): void;
  }
}

const proto = SpreadsheetDocument.prototype;

function ruleElement(cf: ConditionalFormat): XmlElement | undefined {
  return cf.raw.children.find((k): k is XmlElement => k.type === 'element' && k.name === 'cfRule');
}

function addDxf(doc: SpreadsheetDocument, rule: CfRule): number | undefined {
  if (!('format' in rule)) return undefined;
  const id = doc.wb.dxfs.length;
  doc.wb.dxfs = [...doc.wb.dxfs, formatToDxf(rule.format)];
  return id;
}

proto.cfRules = function (this: SpreadsheetDocument, sheet = this.sheet, touching) {
  normalizeCondFormats(sheet);
  const entries: RuleEntry[] = [];
  for (const cf of sheet.condFormats) {
    const raw = ruleElement(cf);
    if (!raw) continue;
    if (touching && !cf.sqref.some((r) => touching.some((t) => rangesIntersect(r, t)))) continue;
    entries.push({
      id: cf.id!,
      rule: ruleFromXml(raw, this.wb.dxfs, this.wb.theme, sheet.x14Rules),
      raw,
      ranges: cf.sqref.map((r) => ({ ...r })),
      priority: numAttr(raw, 'priority') ?? 1000,
      stopIfTrue: raw.attrs.stopIfTrue === '1',
    });
  }
  return entries.sort((a, b) => a.priority - b.priority);
};

function renumber(sheet: Sheet, order: string[]): void {
  // Priorities follow the given order, then the remaining blocks in their current order.
  const all = [...sheet.condFormats].sort((a, b) => (numAttr(ruleElement(a), 'priority') ?? 0) - (numAttr(ruleElement(b), 'priority') ?? 0));
  const ids = [...order, ...all.map((cf) => cf.id!).filter((id) => !order.includes(id))];
  sheet.condFormats = sheet.condFormats.map((cf) => {
    const p = ids.indexOf(cf.id!) + 1;
    const raw = cloneXml(cf.raw);
    const rule = ruleElement({ ...cf, raw });
    if (rule) rule.attrs.priority = String(p);
    return { ...cf, raw };
  });
}

proto.addCondFormat = function (this: SpreadsheetDocument, rule, ranges = this.selection.ranges, opts = {}, sheet = this.sheet) {
  const id = newCfId();
  const bounded = ranges.map((r) => this.bounded(r, sheet) ?? r);
  this.transact('Conditional formatting', (tx) => {
    tx.snapshotSheet(sheet);
    normalizeCondFormats(sheet);
    const dxfId = addDxf(this, rule);
    const ruleEl = ruleToXml(rule, 1, !!opts.stopIfTrue, dxfId, bounded);
    const raw = el('conditionalFormatting', { sqref: sqrefString(bounded) }, [ruleEl]);
    const order = this.cfRules(sheet).map((e) => e.id);
    sheet.condFormats = [...sheet.condFormats, { sqref: bounded.map((r) => ({ ...r })), raw, id }];
    renumber(sheet, [id, ...order]);
  });
  return id;
};

proto.applyRuleList = function (this: SpreadsheetDocument, list, scope, sheet = this.sheet) {
  this.transact('Conditional formatting', (tx) => {
    tx.snapshotSheet(sheet);
    normalizeCondFormats(sheet);
    const before = this.cfRules(sheet);
    const inScope = (e: RuleEntry) => !scope || e.ranges.some((r) => scope.some((s) => rangesIntersect(r, s)));
    // Rules shown in the manager but missing from the list were deleted.
    const listed = new Set(list.map((x) => x.id).filter(Boolean));
    const removed = new Set(before.filter((e) => inScope(e) && !listed.has(e.id)).map((e) => e.id));
    const kept = sheet.condFormats.filter((cf) => !removed.has(cf.id!));
    const byId = new Map(kept.map((cf) => [cf.id!, cf]));
    const order: string[] = [];
    const next: ConditionalFormat[] = kept.filter((cf) => !listed.has(cf.id!));
    for (const item of list) {
      const old = item.id ? byId.get(item.id) : undefined;
      const oldEntry = before.find((e) => e.id === item.id);
      const id = old?.id ?? newCfId();
      order.push(id);
      const sameRule = oldEntry && JSON.stringify(oldEntry.rule) === JSON.stringify(item.rule);
      let ruleEl: XmlElement;
      if (old && sameRule) {
        ruleEl = cloneXml(ruleElement(old)!);
        if (item.stopIfTrue) ruleEl.attrs.stopIfTrue = '1';
        else delete ruleEl.attrs.stopIfTrue;
      } else if (item.rule) ruleEl = ruleToXml(item.rule, 1, item.stopIfTrue, addDxf(this, item.rule), item.ranges);
      else if (item.raw) ruleEl = cloneXml(item.raw);
      else continue;
      const attrs = old ? { ...old.raw.attrs } : {};
      attrs.sqref = sqrefString(item.ranges);
      next.push({ sqref: item.ranges.map((r) => ({ ...r })), raw: el('conditionalFormatting', attrs, [ruleEl]), id });
    }
    sheet.condFormats = next;
    // Unlisted rules (outside the manager's scope) keep their relative order after the listed ones.
    const others = before.filter((e) => !listed.has(e.id) && !removed.has(e.id)).map((e) => e.id);
    const merged: string[] = [];
    // Keep out-of-scope rules where they were relative to the edited ones as far as possible.
    const positions = before.map((e) => e.id);
    for (const id of positions) {
      if (others.includes(id)) merged.push(id);
      else if (listed.has(id) || removed.has(id)) {
        if (!merged.includes('*')) merged.push('*');
      }
    }
    const finalOrder = merged.flatMap((x) => (x === '*' ? order : [x]));
    for (const id of order) if (!finalOrder.includes(id)) finalOrder.push(id);
    renumber(sheet, finalOrder);
  });
};

proto.deleteCondFormat = function (this: SpreadsheetDocument, id, sheet = this.sheet) {
  this.transact('Delete rule', (tx) => {
    tx.snapshotSheet(sheet);
    sheet.condFormats = sheet.condFormats.filter((cf) => cf.id !== id);
  });
};

proto.clearCondFormats = function (this: SpreadsheetDocument, scope, sheet = this.sheet) {
  const sel = this.selectionOf(sheet).ranges;
  let area: Range[] | null = null;
  if (scope === 'selection') area = sel;
  else if (scope === 'table') {
    const a = this.selectionOf(sheet).active;
    const t = sheet.tables.find((x) => a.r >= x.ref.r0 && a.r <= x.ref.r1 && a.c >= x.ref.c0 && a.c <= x.ref.c1);
    if (!t) return;
    area = [t.ref];
  }
  if (!sheet.condFormats.length) return;
  this.transact('Clear rules', (tx) => {
    tx.snapshotSheet(sheet);
    if (!area) {
      sheet.condFormats = [];
      return;
    }
    // Rules lose the cleared cells; rules left without cells are removed.
    sheet.condFormats = sheet.condFormats.flatMap((cf) => {
      let pieces = cf.sqref.map((r) => ({ ...r }));
      for (const a of area!) pieces = pieces.flatMap((p) => subtract(p, a));
      if (!pieces.length) return [];
      const raw = cloneXml(cf.raw);
      raw.attrs.sqref = sqrefString(pieces);
      return [{ ...cf, sqref: pieces, raw }];
    });
  });
};

/** p minus a, as up to four rectangles. */
function subtract(p: Range, a: Range): Range[] {
  const i = intersect(p, a);
  if (!i) return [p];
  const out: Range[] = [];
  if (p.r0 < i.r0) out.push({ r0: p.r0, c0: p.c0, r1: i.r0 - 1, c1: p.c1 });
  if (p.r1 > i.r1) out.push({ r0: i.r1 + 1, c0: p.c0, r1: p.r1, c1: p.c1 });
  if (p.c0 < i.c0) out.push({ r0: i.r0, c0: p.c0, r1: i.r1, c1: i.c0 - 1 });
  if (p.c1 > i.c1) out.push({ r0: i.r0, c0: i.c1 + 1, r1: i.r1, c1: p.c1 });
  return out;
}

export {};
