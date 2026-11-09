/** Page Layout > Themes: the workbook theme's colours and fonts. */
import { SpreadsheetDocument } from '../document';
import type { ThemeSlot } from '../model/color';

export interface ThemePatch {
  name?: string;
  colors?: Record<ThemeSlot, string>;
  majorFont?: string;
  minorFont?: string;
}

declare module '../document' {
  interface SpreadsheetDocument {
    /** Replaces theme colours and/or fonts (cells with theme colours and fonts follow). */
    applyTheme(patch: ThemePatch): void;
  }
}

SpreadsheetDocument.prototype.applyTheme = function (this: SpreadsheetDocument, patch) {
  const wb = this.wb;
  this.transact('Theme', (tx) => {
    tx.snapshotWorkbook(wb);
    // A new object: render caches are keyed by the theme.
    wb.theme = {
      ...wb.theme,
      colors: { ...(patch.colors ?? wb.theme.colors) },
      majorFont: patch.majorFont ?? wb.theme.majorFont,
      minorFont: patch.minorFont ?? wb.theme.minorFont,
    };
    if (patch.name) wb.themeName = patch.name;
    wb.themeChanged = true;
  });
};
