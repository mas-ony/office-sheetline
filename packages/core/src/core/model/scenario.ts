/** Scenarios (Data > What-If Analysis > Scenario Manager), stored per sheet. */

export interface Scenario {
  name: string;
  /** Changing cells (on the scenario's sheet). */
  cells: { r: number; c: number }[];
  /** Values as entered, one per changing cell. */
  values: string[];
  comment?: string;
  locked?: boolean;
  hidden?: boolean;
  user?: string;
}

export interface SheetScenarios {
  list: Scenario[];
  /** Index of the scenario shown last. */
  current?: number;
}
