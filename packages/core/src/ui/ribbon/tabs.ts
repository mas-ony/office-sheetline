/** The ribbon tabs in order, and every command they offer (command palette, shortcuts). */
import { commandsOf, type Command, type Ctx, type Tab } from './types';
import { homeTab } from './homeTab';
import { insertTab } from './insertTab';
import { pageLayoutTab } from './pageLayoutTab';
import { formulasTab } from './formulasTab';
import { dataTab } from './dataTab';
import { reviewTab } from './reviewTab';
import { viewTab } from './viewTab';
import { settingsTab } from './settingsTab';
import { tableDesignTab, chartDesignTab, shapeFormatTab, pictureFormatTab, sparklineTab } from './contextTabs';

export const RIBBON_TABS: Tab[] = [
  homeTab,
  insertTab,
  pageLayoutTab,
  formulasTab,
  dataTab,
  reviewTab,
  viewTab,
  settingsTab,
  tableDesignTab,
  sparklineTab,
  chartDesignTab,
  shapeFormatTab,
  pictureFormatTab,
];

/** Every command of every tab (contextual tabs included), without duplicates. */
export function allCommands(ctx: Ctx): Command[] {
  const seen = new Set<string>();
  const out: Command[] = [];
  for (const tab of RIBBON_TABS) {
    for (const g of tab.groups(ctx)) {
      for (const c of [...commandsOf(g.controls), ...(g.launcher ? [g.launcher] : [])]) {
        if (seen.has(c.id)) continue;
        seen.add(c.id);
        out.push(c);
      }
    }
  }
  return out;
}
