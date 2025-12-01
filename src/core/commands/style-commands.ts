/**
 * Style Commands - Commands that modify conditional and direct cell styles
 *
 * These commands do NOT require re-evaluation after execution.
 * They only affect styling, not cell values or formulas.
 */

import type { StyleManager } from "../managers/style-manager";
import type { ConditionalStyle, DirectCellStyle, RangeAddress } from "../types";
import type { EngineCommand, EngineAction } from "./types";
import { ActionTypes } from "./types";

/**
 * Command to add a conditional style.
 */
export class AddConditionalStyleCommand implements EngineCommand {
  readonly requiresReevaluation = false;
  private addedIndex: number = -1;

  constructor(
    private styleManager: StyleManager,
    private style: ConditionalStyle
  ) {}

  execute(): void {
    // Track the index where it's added
    const allStyles = this.styleManager.getAllConditionalStyles();
    this.addedIndex = allStyles.length;
    this.styleManager.addConditionalStyle(this.style);
  }

  undo(): void {
    // Remove by finding and removing the style
    // Since we know the style object, we can find it
    const allStyles = this.styleManager.getAllConditionalStyles();
    const workbookName = this.style.areas[0]?.workbookName;
    if (workbookName) {
      // Find the index of this style in the workbook's styles
      let workbookStyleIndex = 0;
      for (let i = 0; i < allStyles.length; i++) {
        const s = allStyles[i];
        if (s?.areas.some((a) => a.workbookName === workbookName)) {
          if (s === this.style || i === this.addedIndex) {
            this.styleManager.removeConditionalStyle(workbookName, workbookStyleIndex);
            return;
          }
          workbookStyleIndex++;
        }
      }
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.ADD_CONDITIONAL_STYLE,
      payload: { style: this.style },
    };
  }
}

/**
 * Command to remove a conditional style.
 */
export class RemoveConditionalStyleCommand implements EngineCommand {
  readonly requiresReevaluation = false;
  private removedStyle: ConditionalStyle | undefined;

  constructor(
    private styleManager: StyleManager,
    private workbookName: string,
    private index: number
  ) {}

  execute(): void {
    // Capture the style before removal
    const allStyles = this.styleManager.getAllConditionalStyles();
    const workbookStyles = allStyles.filter((s) =>
      s.areas.some((a) => a.workbookName === this.workbookName)
    );
    this.removedStyle = workbookStyles[this.index];

    this.styleManager.removeConditionalStyle(this.workbookName, this.index);
  }

  undo(): void {
    if (!this.removedStyle) return;
    this.styleManager.addConditionalStyle(this.removedStyle);
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.REMOVE_CONDITIONAL_STYLE,
      payload: {
        workbookName: this.workbookName,
        index: this.index,
      },
    };
  }
}

/**
 * Command to add a direct cell style.
 */
export class AddCellStyleCommand implements EngineCommand {
  readonly requiresReevaluation = false;
  private addedIndex: number = -1;

  constructor(
    private styleManager: StyleManager,
    private style: DirectCellStyle
  ) {}

  execute(): void {
    // Track the index where it's added
    const allStyles = this.styleManager.getAllCellStyles();
    this.addedIndex = allStyles.length;
    this.styleManager.addCellStyle(this.style);
  }

  undo(): void {
    // Remove by finding and removing the style
    const allStyles = this.styleManager.getAllCellStyles();
    const workbookName = this.style.areas[0]?.workbookName;
    if (workbookName) {
      // Find the index of this style in the workbook's styles
      let workbookStyleIndex = 0;
      for (let i = 0; i < allStyles.length; i++) {
        const s = allStyles[i];
        if (s?.areas.some((a) => a.workbookName === workbookName)) {
          if (s === this.style || i === this.addedIndex) {
            this.styleManager.removeCellStyle(workbookName, workbookStyleIndex);
            return;
          }
          workbookStyleIndex++;
        }
      }
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.ADD_CELL_STYLE,
      payload: { style: this.style },
    };
  }
}

/**
 * Command to remove a direct cell style.
 */
export class RemoveCellStyleCommand implements EngineCommand {
  readonly requiresReevaluation = false;
  private removedStyle: DirectCellStyle | undefined;

  constructor(
    private styleManager: StyleManager,
    private workbookName: string,
    private index: number
  ) {}

  execute(): void {
    // Capture the style before removal
    const allStyles = this.styleManager.getAllCellStyles();
    const workbookStyles = allStyles.filter((s) =>
      s.areas.some((a) => a.workbookName === this.workbookName)
    );
    this.removedStyle = workbookStyles[this.index];

    this.styleManager.removeCellStyle(this.workbookName, this.index);
  }

  undo(): void {
    if (!this.removedStyle) return;
    this.styleManager.addCellStyle(this.removedStyle);
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.REMOVE_CELL_STYLE,
      payload: {
        workbookName: this.workbookName,
        index: this.index,
      },
    };
  }
}

/**
 * Captured styles for undo purposes.
 */
interface ClearedStylesSnapshot {
  conditionalStyles: ConditionalStyle[];
  cellStyles: DirectCellStyle[];
}

/**
 * Command to clear cell styles in a range.
 */
export class ClearCellStylesCommand implements EngineCommand {
  readonly requiresReevaluation = false;
  private snapshot: ClearedStylesSnapshot | undefined;

  constructor(
    private styleManager: StyleManager,
    private range: RangeAddress
  ) {}

  execute(): void {
    // Capture affected styles before clearing
    this.snapshot = {
      conditionalStyles: this.styleManager.getConditionalStylesIntersectingWithRange(
        this.range
      ),
      cellStyles: this.styleManager.getStylesIntersectingWithRange(this.range),
    };

    this.styleManager.clearCellStyles(this.range);
  }

  undo(): void {
    if (!this.snapshot) return;

    // Note: This is a simplified undo - it re-adds the styles
    // A more sophisticated implementation would restore the exact state
    // including any range adjustments that clearCellStyles made
    for (const style of this.snapshot.conditionalStyles) {
      this.styleManager.addConditionalStyle(style);
    }
    for (const style of this.snapshot.cellStyles) {
      this.styleManager.addCellStyle(style);
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.CLEAR_CELL_STYLES,
      payload: { range: this.range },
    };
  }
}

