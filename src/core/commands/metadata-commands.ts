/**
 * Metadata Commands - Commands that modify cell/sheet/workbook metadata
 *
 * These commands do NOT require re-evaluation after execution.
 * They only affect metadata, not cell values or formulas.
 */

import type { WorkbookManager } from "../managers/workbook-manager";
import type { CellAddress } from "../types";
import type { EngineCommand, EngineAction } from "./types";
import { ActionTypes } from "./types";
import { getCellReference } from "../utils";

/**
 * Command to set cell metadata.
 */
export class SetCellMetadataCommand<TMetadata = unknown> implements EngineCommand {
  readonly requiresReevaluation = false;
  private previousMetadata: TMetadata | undefined;
  private hadPreviousMetadata = false;

  constructor(
    private workbookManager: WorkbookManager,
    private address: CellAddress,
    private newMetadata: TMetadata | undefined
  ) {}

  execute(): void {
    // Capture previous metadata
    const sheet = this.workbookManager.getSheet({
      workbookName: this.address.workbookName,
      sheetName: this.address.sheetName,
    });

    if (sheet) {
      const key = getCellReference(this.address);
      this.hadPreviousMetadata = sheet.metadata.has(key);
      this.previousMetadata = sheet.metadata.get(key) as TMetadata | undefined;
    }

    this.workbookManager.setCellMetadata(this.address, this.newMetadata);
  }

  undo(): void {
    if (this.hadPreviousMetadata) {
      this.workbookManager.setCellMetadata(this.address, this.previousMetadata);
    } else {
      this.workbookManager.setCellMetadata(this.address, undefined);
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.SET_CELL_METADATA,
      payload: {
        address: this.address,
        metadata: this.newMetadata,
      },
    };
  }
}

/**
 * Command to set sheet metadata.
 */
export class SetSheetMetadataCommand<TMetadata = unknown> implements EngineCommand {
  readonly requiresReevaluation = false;
  private previousMetadata: TMetadata | undefined;

  constructor(
    private workbookManager: WorkbookManager,
    private opts: { workbookName: string; sheetName: string },
    private newMetadata: TMetadata
  ) {}

  execute(): void {
    // Capture previous metadata
    const sheet = this.workbookManager.getSheet(this.opts);
    if (sheet) {
      this.previousMetadata = sheet.sheetMetadata as TMetadata | undefined;
    }

    this.workbookManager.setSheetMetadata(this.opts, this.newMetadata);
  }

  undo(): void {
    if (this.previousMetadata !== undefined) {
      this.workbookManager.setSheetMetadata(this.opts, this.previousMetadata);
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.SET_SHEET_METADATA,
      payload: {
        opts: this.opts,
        metadata: this.newMetadata,
      },
    };
  }
}

/**
 * Command to set workbook metadata.
 */
export class SetWorkbookMetadataCommand<TMetadata = unknown> implements EngineCommand {
  readonly requiresReevaluation = false;
  private previousMetadata: TMetadata | undefined;

  constructor(
    private workbookManager: WorkbookManager,
    private workbookName: string,
    private newMetadata: TMetadata
  ) {}

  execute(): void {
    // Capture previous metadata
    const workbook = this.workbookManager.getWorkbooks().get(this.workbookName);
    if (workbook) {
      this.previousMetadata = workbook.workbookMetadata as TMetadata | undefined;
    }

    this.workbookManager.setWorkbookMetadata(this.workbookName, this.newMetadata);
  }

  undo(): void {
    if (this.previousMetadata !== undefined) {
      this.workbookManager.setWorkbookMetadata(this.workbookName, this.previousMetadata);
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.SET_WORKBOOK_METADATA,
      payload: {
        workbookName: this.workbookName,
        metadata: this.newMetadata,
      },
    };
  }
}

