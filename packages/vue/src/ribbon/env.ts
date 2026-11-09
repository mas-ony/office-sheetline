/** What the ribbon's controls share: the context and the open drop-down. */
import type { Ctx } from '@sheetline/core/views';

export type Open = { id: string; rect: DOMRect; kind: 'menu' | 'panel' | 'group' | 'color' | 'more' } | null;

export interface ControlEnv {
  ctx: Ctx;
  open: Open;
  setOpen: (o: Open) => void;
  /** Toolbar layout: large buttons render as small ones. */
  compact: boolean;
  phone: boolean;
  /** Closes the popup holding the control once a command ran (toggles keep it open). */
  done?: () => void;
}
