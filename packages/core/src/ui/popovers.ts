/** Helpers of the cell pop-ups: comment cards, the column filter drop-down and pop-ups next to a cell. */
import type { FilterOperator } from '../core/model/sheet';

/** The time stamp on a comment card, such as "10/21/2026 7:12:27". */
export function commentStamp(iso?: string): string {
  if (!iso) return '';
  const t = Date.parse(iso.endsWith('Z') || /[+-]\d\d:\d\d$/.test(iso) ? iso : `${iso}Z`);
  if (!Number.isFinite(t)) return '';
  const d = new Date(t);
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()} ${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}`;
}

/** The initial and colour of an author's avatar. */
export function avatarOf(name: string): { initial: string; color: string } {
  const initial = (name || '?').trim().slice(0, 1).toUpperCase();
  let h = 0;
  for (const ch of name) h = (h * 31 + ch.charCodeAt(0)) % 360;
  return { initial, color: `hsl(${h}, 45%, 38%)` };
}

/** Text Filters of the filter drop-down. */
export const TEXT_CONDITIONS: { op: FilterOperator; label: string }[] = [
  { op: 'equal', label: 'Equals...' },
  { op: 'notEqual', label: 'Does Not Equal...' },
  { op: 'beginsWith', label: 'Begins With...' },
  { op: 'endsWith', label: 'Ends With...' },
  { op: 'contains', label: 'Contains...' },
  { op: 'notContains', label: 'Does Not Contain...' },
];

/** Number Filters of the filter drop-down. */
export const NUMBER_CONDITIONS: { op: FilterOperator; label: string }[] = [
  { op: 'equal', label: 'Equals...' },
  { op: 'notEqual', label: 'Does Not Equal...' },
  { op: 'greaterThan', label: 'Greater Than...' },
  { op: 'greaterThanOrEqual', label: 'Greater Than Or Equal To...' },
  { op: 'lessThan', label: 'Less Than...' },
  { op: 'lessThanOrEqual', label: 'Less Than Or Equal To...' },
];

/** The operators of the Custom AutoFilter dialog. */
export const FILTER_OPERATORS: { op: FilterOperator; label: string }[] = [
  { op: 'equal', label: 'equals' },
  { op: 'notEqual', label: 'does not equal' },
  { op: 'greaterThan', label: 'is greater than' },
  { op: 'greaterThanOrEqual', label: 'is greater than or equal to' },
  { op: 'lessThan', label: 'is less than' },
  { op: 'lessThanOrEqual', label: 'is less than or equal to' },
  { op: 'beginsWith', label: 'begins with' },
  { op: 'endsWith', label: 'ends with' },
  { op: 'contains', label: 'contains' },
  { op: 'notContains', label: 'does not contain' },
];

/**
 * Places a fixed pop-up right of a rectangle (or left of it, or under it when neither fits), or
 * under it with `below`, inside the visible viewport.
 */
export function placeNextTo(el: HTMLElement, rect: DOMRect, below?: boolean): void {
  const w = el.offsetWidth;
  const h = el.offsetHeight;
  const vw = window.innerWidth;
  const vh = window.visualViewport?.height ?? window.innerHeight;
  let x = below ? rect.left : rect.right + 4;
  let y = below ? rect.bottom : rect.top;
  if (x + w > vw - 4) x = below ? Math.max(4, vw - w - 4) : rect.left - w - 4;
  if (x < 4) {
    x = Math.max(4, Math.min(rect.left, vw - w - 4));
    y = rect.bottom + 4;
  }
  if (y + h > vh - 4) y = Math.max(4, vh - h - 4);
  el.style.left = `${x}px`;
  el.style.top = `${y}px`;
}
