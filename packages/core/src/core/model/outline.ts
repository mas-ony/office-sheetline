/** Outline levels of rows and columns (grouping): the groups and their summary lines. */
import type { Sheet } from './sheet';
import { MAX_COLS, MAX_ROWS } from '../refs';

export type OutlineAxis = 'row' | 'col';

export const MAX_OUTLINE_LEVEL = 7;

/** A group of consecutive lines at one outline level, with its summary line (button). */
export interface OutlineGroup {
  axis: OutlineAxis;
  level: number;
  start: number;
  end: number;
  /** The summary line: below/right of the group, or above/left of it. */
  summary: number;
  collapsed: boolean;
}

export function levelOf(sheet: Sheet, axis: OutlineAxis, i: number): number {
  return (axis === 'row' ? sheet.rows.get(i)?.outline : sheet.cols.get(i)?.outline) ?? 0;
}

/** Outlined lines of an axis in order. */
export function outlined(sheet: Sheet, axis: OutlineAxis): number[] {
  const map = axis === 'row' ? sheet.rows : sheet.cols;
  const out: number[] = [];
  for (const [i, info] of map) if (info.outline) out.push(i);
  return out.sort((a, b) => a - b);
}

/** Highest outline level of an axis. */
export function maxOutlineLevel(sheet: Sheet, axis: OutlineAxis): number {
  const map = axis === 'row' ? sheet.rows : sheet.cols;
  let max = 0;
  for (const info of map.values()) if (info.outline && info.outline > max) max = info.outline;
  return Math.min(MAX_OUTLINE_LEVEL, max);
}

/** Every group of an axis, outer levels first. */
export function outlineGroups(sheet: Sheet, axis: OutlineAxis): OutlineGroup[] {
  const lines = outlined(sheet, axis);
  if (!lines.length) return [];
  const below = axis === 'row' ? sheet.outlineBelow : sheet.outlineRight;
  const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
  const max = maxOutlineLevel(sheet, axis);
  const out: OutlineGroup[] = [];
  for (let level = 1; level <= max; level++) {
    let start = -1;
    let prev = -2;
    const flush = () => {
      if (start < 0) return;
      const summary = below ? Math.min(limit - 1, prev + 1) : Math.max(0, start - 1);
      const info = axis === 'row' ? sheet.rows.get(summary) : sheet.cols.get(summary);
      let hidden = true;
      for (let i = start; i <= prev && hidden; i++) hidden = axis === 'row' ? !!sheet.rows.get(i)?.hidden : !!sheet.cols.get(i)?.hidden;
      out.push({ axis, level, start, end: prev, summary, collapsed: !!info?.collapsed || hidden });
    };
    for (const i of lines) {
      if (levelOf(sheet, axis, i) < level) continue;
      if (i !== prev + 1) {
        flush();
        start = i;
      }
      prev = i;
    }
    flush();
  }
  return out;
}

