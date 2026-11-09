/** Formulas > Calculation: automatic or manual calculation, Calculate Now and Calculate Sheet. */
import { SpreadsheetDocument } from '../document';

declare module '../document' {
  interface SpreadsheetDocument {
    setCalcMode(mode: 'auto' | 'autoNoTable' | 'manual'): void;
    /** F9: recalculates what changed (all formulas when `full`). */
    calculateNow(full?: boolean): void;
    /** Shift+F9 (the engine recalculates the workbook's dirty formulas; sheets share dependencies). */
    calculateSheet(): void;
  }
}

const proto = SpreadsheetDocument.prototype;

proto.setCalcMode = function (this: SpreadsheetDocument, mode) {
  if (this.wb.calcMode === mode) return;
  this.wb.calcMode = mode;
  this.modified = true;
  if (mode !== 'manual') this.calculateNow();
  else this.emit('change', { structure: false, layout: false, ranges: [] });
};

proto.calculateNow = function (this: SpreadsheetDocument, full = false) {
  if (full) this.engine.invalidateAll();
  this.engine.recalc();
  this.calcPending = false;
  this.emit('change', { structure: false, layout: false, ranges: [] });
};

proto.calculateSheet = function (this: SpreadsheetDocument) {
  this.calculateNow(false);
};

export {};
