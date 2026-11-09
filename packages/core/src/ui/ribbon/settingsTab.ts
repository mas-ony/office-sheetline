/** Settings tab: Spread Settings (workbook behaviour) and Sheet Settings (sheet view and behaviour). */
import { TINT, type Command, type Tab } from './types';

const spread = (tab: string, label: string, icon: Command['icon']): Command => ({
  id: `spread-settings-${tab}`,
  label,
  tip: `Spread Settings: ${label}`,
  icon,
  tint: TINT.blue,
  run: (ctx) => ctx.openDialog({ kind: 'spread-settings', tab }),
});

const sheet = (tab: string, label: string, icon: Command['icon']): Command => ({
  id: `sheet-settings-${tab}`,
  label,
  tip: `Sheet Settings: ${label}`,
  icon,
  tint: TINT.green,
  run: (ctx) => ctx.openDialog({ kind: 'sheet-settings', tab }),
});

export const S = {
  general: spread('general', 'General', 'settings'),
  scrollbars: spread('scrollbars', 'Scroll Bars', 'swap_vert'),
  calculation: spread('calculation', 'Calculation', 'calculate'),
  tabStrip: spread('tabstrip', 'Tab Strip', 'tab'),
  sheetGeneral: sheet('general', 'General', 'tune'),
  gridlines: sheet('gridlines', 'Gridlines', 'grid_on'),
  headers: sheet('headers', 'Headers', 'view_column'),
  sheetTab: sheet('sheettab', 'Sheet Tab', 'sheet_tab'),
};

export const settingsTab: Tab = {
  id: 'settings',
  label: 'Settings',
  groups: () => [
    {
      id: 'spread-settings',
      label: 'Spread Settings',
      icon: 'settings',
      controls: [{ type: 'large', cmd: S.general }, { type: 'large', cmd: S.scrollbars }, { type: 'large', cmd: S.calculation }, { type: 'large', cmd: S.tabStrip }],
    },
    {
      id: 'sheet-settings',
      label: 'Sheet Settings',
      icon: 'tune',
      controls: [{ type: 'large', cmd: S.sheetGeneral }, { type: 'large', cmd: S.gridlines }, { type: 'large', cmd: S.headers }, { type: 'large', cmd: S.sheetTab }],
    },
  ],
};
