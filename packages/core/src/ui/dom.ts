/** Small browser helpers of the interface: media queries, touch input and pop-up placement. */
import type { MenuAnchor } from './menuItems';

/** Phones: narrow screens (the touch layout, bottom sheets and full-screen dialogs). */
export const PHONE_QUERY = '(max-width: 720px)';

export function matchesMedia(query: string): boolean {
  return typeof window !== 'undefined' && !!window.matchMedia && window.matchMedia(query).matches;
}

/** Calls `fn` when the query starts or stops matching; returns the function that stops watching. */
export function watchMedia(query: string, fn: (matches: boolean) => void): () => void {
  const mq = typeof window !== 'undefined' ? window.matchMedia?.(query) : undefined;
  if (!mq) return () => undefined;
  const on = () => fn(mq.matches);
  mq.addEventListener?.('change', on);
  return () => mq.removeEventListener?.('change', on);
}

/** Whether the primary input is touch (phones and tablets). */
export function isTouchDevice(): boolean {
  if (typeof window === 'undefined') return false;
  return window.matchMedia?.('(pointer: coarse)').matches || 'ontouchstart' in window;
}

function isRect(at: MenuAnchor): at is DOMRect {
  return 'width' in at;
}

/**
 * Places a fixed-position menu at its anchor, inside the viewport: below a rectangle (or above it
 * when there is no room), right of it for a submenu, or at a point.
 */
export function placeMenu(el: HTMLElement, at: MenuAnchor, side: 'below' | 'right'): void {
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  const w = el.offsetWidth;
  const h = el.offsetHeight;
  let x: number;
  let y: number;
  if (isRect(at)) {
    if (side === 'right') {
      x = at.right - 2;
      y = at.top - 6;
      if (x + w > vw - 4) x = Math.max(4, at.left - w + 2);
    } else {
      x = at.left;
      y = at.bottom + 2;
      if (y + h > vh - 4 && at.top - h - 2 > 4) y = at.top - h - 2;
    }
  } else {
    x = at.x;
    y = at.y;
    if (x + w > vw - 4) x = Math.max(4, vw - w - 4);
    if (y + h > vh - 4) y = Math.max(4, at.y - h);
  }
  if (x + w > vw - 4) x = Math.max(4, vw - w - 4);
  if (y + h > vh - 4) y = Math.max(4, vh - h - 4);
  el.style.left = `${Math.round(x)}px`;
  el.style.top = `${Math.round(y)}px`;
}
