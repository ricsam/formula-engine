/**
 * Content Commands - Commands that modify cell/sheet content
 *
 * These commands all require re-evaluation after execution.
 */

import type { WorkbookManager } from "../managers/workbook-manager";
import type { CopyManager } from "../managers/copy-manager";
import type {
  CellAddress,
  CopyCellsOptions,
  RangeAddress,
  SerializedCellValue,
} from "../types";
import type { EngineCommand, EngineAction } from "./types";
import { ActionTypes } from "./types";
import { getCellReference, parseCellReference } from "../utils";

/**
 * Command to set a single cell's content.
 */
export class SetCellContentCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private previousValue: SerializedCellValue | undefined;
  private hadPreviousValue = false;

  constructor(
    private workbookManager: WorkbookManager,
    private address: CellAddress,
    private newContent: SerializedCellValue
  ) {}

  execute(): void {
    // Capture previous value
    const sheet = this.workbookManager.getSheet({
      workbookName: this.address.workbookName,
      sheetName: this.address.sheetName,
    });
    if (sheet) {
      const key = getCellReference(this.address);
      this.hadPreviousValue = sheet.content.has(key);
      this.previousValue = sheet.content.get(key);
    }

    this.workbookManager.setCellContent(this.address, this.newContent);
  }

  undo(): void {
    if (this.hadPreviousValue) {
      this.workbookManager.setCellContent(this.address, this.previousValue);
    } else {
      this.workbookManager.setCellContent(this.address, undefined);
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.SET_CELL_CONTENT,
      payload: {
        address: this.address,
        content: this.newContent,
      },
    };
  }
}

/**
 * Command to set an entire sheet's content.
 */
export class SetSheetContentCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private previousContent: Map<string, SerializedCellValue> | undefined;

  constructor(
    private workbookManager: WorkbookManager,
    private opts: { workbookName: string; sheetName: string },
    private newContent: Map<string, SerializedCellValue>
  ) {}

  execute(): void {
    // Capture previous content
    const sheet = this.workbookManager.getSheet(this.opts);
    if (sheet) {
      this.previousContent = new Map(sheet.content);
    }

    this.workbookManager.setSheetContent(this.opts, this.newContent);
  }

  undo(): void {
    if (this.previousContent) {
      this.workbookManager.setSheetContent(this.opts, this.previousContent);
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.SET_SHEET_CONTENT,
      payload: {
        opts: this.opts,
        content: Array.from(this.newContent.entries()),
      },
    };
  }
}

/**
 * Command to clear a range of cells.
 */
export class ClearRangeCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private clearedCells: Map<string, SerializedCellValue> = new Map();

  constructor(
    private workbookManager: WorkbookManager,
    private address: RangeAddress
  ) {}

  execute(): void {
    // Capture cells before clearing using the optimized iterator
    // This handles infinite ranges by only iterating over cells that actually exist
    try {
      for (const cellAddress of this.workbookManager.iterateCellsInRange(this.address)) {
        const key = getCellReference(cellAddress);
        const value = this.workbookManager.getCellContent(cellAddress);
        if (value !== undefined) {
          this.clearedCells.set(key, value);
        }
      }
    } catch {
      // Sheet doesn't exist, nothing to capture
    }

    this.workbookManager.clearSpreadsheetRange(this.address);
  }

  undo(): void {
    // Restore cleared cells
    for (const [key, value] of this.clearedCells) {
      const { colIndex, rowIndex } = parseCellReference(key);
      this.workbookManager.setCellContent(
        {
          workbookName: this.address.workbookName,
          sheetName: this.address.sheetName,
          colIndex,
          rowIndex,
        },
        value
      );
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.CLEAR_RANGE,
      payload: {
        address: this.address,
      },
    };
  }
}

/**
 * Snapshot of a cell for undo purposes.
 */
interface CellSnapshot {
  address: CellAddress;
  content: SerializedCellValue | undefined;
  metadata: unknown | undefined;
}

/**
 * Command to paste cells.
 */
export class PasteCellsCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private targetSnapshots: CellSnapshot[] = [];
  private sourceSnapshots: CellSnapshot[] = [];

  constructor(
    private workbookManager: WorkbookManager,
    private copyManager: CopyManager,
    private source: CellAddress[],
    private target: CellAddress,
    private options: CopyCellsOptions
  ) {}

  execute(): void {
    // Calculate target cells and capture their state
    if (this.source.length > 0) {
      const firstSource = this.source[0]!;
      const colOffset = this.target.colIndex - firstSource.colIndex;
      const rowOffset = this.target.rowIndex - firstSource.rowIndex;

      // Capture target cells before paste
      for (const sourceCell of this.source) {
        const targetCell: CellAddress = {
          workbookName: this.target.workbookName,
          sheetName: this.target.sheetName,
          colIndex: sourceCell.colIndex + colOffset,
          rowIndex: sourceCell.rowIndex + rowOffset,
        };

        const sheet = this.workbookManager.getSheet({
          workbookName: targetCell.workbookName,
          sheetName: targetCell.sheetName,
        });

        if (sheet) {
          const key = getCellReference(targetCell);
          this.targetSnapshots.push({
            address: targetCell,
            content: sheet.content.get(key),
            metadata: sheet.metadata.get(key),
          });
        }
      }

      // If cut operation, capture source cells too
      if (this.options.cut) {
        for (const sourceCell of this.source) {
          const sheet = this.workbookManager.getSheet({
            workbookName: sourceCell.workbookName,
            sheetName: sourceCell.sheetName,
          });

          if (sheet) {
            const key = getCellReference(sourceCell);
            this.sourceSnapshots.push({
              address: sourceCell,
              content: sheet.content.get(key),
              metadata: sheet.metadata.get(key),
            });
          }
        }
      }
    }

    this.copyManager.pasteCells(this.source, this.target, this.options);
  }

  undo(): void {
    // Restore target cells
    for (const snapshot of this.targetSnapshots) {
      this.workbookManager.setCellContent(snapshot.address, snapshot.content);
      if (snapshot.metadata !== undefined) {
        this.workbookManager.setCellMetadata(snapshot.address, snapshot.metadata);
      }
    }

    // Restore source cells if it was a cut operation
    if (this.options.cut) {
      for (const snapshot of this.sourceSnapshots) {
        this.workbookManager.setCellContent(snapshot.address, snapshot.content);
        if (snapshot.metadata !== undefined) {
          this.workbookManager.setCellMetadata(snapshot.address, snapshot.metadata);
        }
      }
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.PASTE_CELLS,
      payload: {
        source: this.source,
        target: this.target,
        options: this.options,
      },
    };
  }
}

/**
 * Command to fill areas with a seed range.
 */
export class FillAreasCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private targetSnapshots: Map<string, CellSnapshot> = new Map();

  constructor(
    private workbookManager: WorkbookManager,
    private copyManager: CopyManager,
    private seedRange: RangeAddress,
    private targetRanges: RangeAddress[],
    private options: CopyCellsOptions
  ) {}

  execute(): void {
    // Capture all target cells before filling
    for (const targetRange of this.targetRanges) {
      const sheet = this.workbookManager.getSheet({
        workbookName: targetRange.workbookName,
        sheetName: targetRange.sheetName,
      });

      if (sheet) {
        const { start, end } = targetRange.range;
        const endCol = end.col.type === "number" ? end.col.value : start.col + 100;
        const endRow = end.row.type === "number" ? end.row.value : start.row + 100;

        for (let col = start.col; col <= endCol; col++) {
          for (let row = start.row; row <= endRow; row++) {
            const address: CellAddress = {
              workbookName: targetRange.workbookName,
              sheetName: targetRange.sheetName,
              colIndex: col,
              rowIndex: row,
            };
            const key = getCellReference(address);
            const fullKey = `${targetRange.workbookName}:${targetRange.sheetName}:${key}`;

            if (!this.targetSnapshots.has(fullKey)) {
              this.targetSnapshots.set(fullKey, {
                address,
                content: sheet.content.get(key),
                metadata: sheet.metadata.get(key),
              });
            }
          }
        }
      }
    }

    this.copyManager.fillAreas(this.seedRange, this.targetRanges, this.options);
  }

  undo(): void {
    // Restore all target cells
    for (const snapshot of this.targetSnapshots.values()) {
      this.workbookManager.setCellContent(snapshot.address, snapshot.content);
      if (snapshot.metadata !== undefined) {
        this.workbookManager.setCellMetadata(snapshot.address, snapshot.metadata);
      }
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.FILL_AREAS,
      payload: {
        seedRange: this.seedRange,
        targetRanges: this.targetRanges,
        options: this.options,
      },
    };
  }
}

/**
 * Command to move a single cell.
 */
export class MoveCellCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private sourceSnapshot: CellSnapshot | undefined;
  private targetSnapshot: CellSnapshot | undefined;

  constructor(
    private workbookManager: WorkbookManager,
    private copyManager: CopyManager,
    private source: CellAddress,
    private target: CellAddress
  ) {}

  execute(): void {
    // Capture source cell
    const sourceSheet = this.workbookManager.getSheet({
      workbookName: this.source.workbookName,
      sheetName: this.source.sheetName,
    });

    if (sourceSheet) {
      const key = getCellReference(this.source);
      this.sourceSnapshot = {
        address: this.source,
        content: sourceSheet.content.get(key),
        metadata: sourceSheet.metadata.get(key),
      };
    }

    // Capture target cell
    const targetSheet = this.workbookManager.getSheet({
      workbookName: this.target.workbookName,
      sheetName: this.target.sheetName,
    });

    if (targetSheet) {
      const key = getCellReference(this.target);
      this.targetSnapshot = {
        address: this.target,
        content: targetSheet.content.get(key),
        metadata: targetSheet.metadata.get(key),
      };
    }

    // Execute the move via paste with cut option
    this.copyManager.pasteCells([this.source], this.target, {
      cut: true,
      type: "formula",
      include: "all",
    });
  }

  undo(): void {
    // Restore source cell
    if (this.sourceSnapshot) {
      this.workbookManager.setCellContent(
        this.sourceSnapshot.address,
        this.sourceSnapshot.content
      );
      if (this.sourceSnapshot.metadata !== undefined) {
        this.workbookManager.setCellMetadata(
          this.sourceSnapshot.address,
          this.sourceSnapshot.metadata
        );
      }
    }

    // Restore target cell
    if (this.targetSnapshot) {
      this.workbookManager.setCellContent(
        this.targetSnapshot.address,
        this.targetSnapshot.content
      );
      if (this.targetSnapshot.metadata !== undefined) {
        this.workbookManager.setCellMetadata(
          this.targetSnapshot.address,
          this.targetSnapshot.metadata
        );
      }
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.MOVE_CELL,
      payload: {
        source: this.source,
        target: this.target,
      },
    };
  }
}

/**
 * Command to move a range of cells.
 */
export class MoveRangeCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private sourceSnapshots: CellSnapshot[] = [];
  private targetSnapshots: CellSnapshot[] = [];

  constructor(
    private workbookManager: WorkbookManager,
    private copyManager: CopyManager,
    private sourceRange: RangeAddress,
    private target: CellAddress
  ) {}

  execute(): void {
    // Expand source range to cells
    const sourceCells = this.copyManager.expandRangeToCells(this.sourceRange);

    // Capture source cells
    for (const cell of sourceCells) {
      const sheet = this.workbookManager.getSheet({
        workbookName: cell.workbookName,
        sheetName: cell.sheetName,
      });

      if (sheet) {
        const key = getCellReference(cell);
        this.sourceSnapshots.push({
          address: cell,
          content: sheet.content.get(key),
          metadata: sheet.metadata.get(key),
        });
      }
    }

    // Calculate and capture target cells
    if (sourceCells.length > 0) {
      const firstSource = sourceCells[0]!;
      const colOffset = this.target.colIndex - firstSource.colIndex;
      const rowOffset = this.target.rowIndex - firstSource.rowIndex;

      for (const sourceCell of sourceCells) {
        const targetCell: CellAddress = {
          workbookName: this.target.workbookName,
          sheetName: this.target.sheetName,
          colIndex: sourceCell.colIndex + colOffset,
          rowIndex: sourceCell.rowIndex + rowOffset,
        };

        const sheet = this.workbookManager.getSheet({
          workbookName: targetCell.workbookName,
          sheetName: targetCell.sheetName,
        });

        if (sheet) {
          const key = getCellReference(targetCell);
          this.targetSnapshots.push({
            address: targetCell,
            content: sheet.content.get(key),
            metadata: sheet.metadata.get(key),
          });
        }
      }
    }

    // Execute the move
    this.copyManager.pasteCells(sourceCells, this.target, {
      cut: true,
      type: "formula",
      include: "all",
    });
  }

  undo(): void {
    // Restore source cells
    for (const snapshot of this.sourceSnapshots) {
      this.workbookManager.setCellContent(snapshot.address, snapshot.content);
      if (snapshot.metadata !== undefined) {
        this.workbookManager.setCellMetadata(snapshot.address, snapshot.metadata);
      }
    }

    // Restore target cells
    for (const snapshot of this.targetSnapshots) {
      this.workbookManager.setCellContent(snapshot.address, snapshot.content);
      if (snapshot.metadata !== undefined) {
        this.workbookManager.setCellMetadata(snapshot.address, snapshot.metadata);
      }
    }
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.MOVE_RANGE,
      payload: {
        sourceRange: this.sourceRange,
        target: this.target,
      },
    };
  }
}

/**
 * Command to auto-fill ranges based on a seed pattern.
 * Captures the entire affected area for undo.
 */
export class AutoFillCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private previousContent: Map<string, SerializedCellValue> = new Map();
  private previousMetadata: Map<string, unknown> = new Map();

  constructor(
    private workbookManager: WorkbookManager,
    private styleManager: import("../managers/style-manager").StyleManager,
    private autoFillManager: import("../autofill-utils").AutoFill,
    private opts: { sheetName: string; workbookName: string },
    private seedRange: import("../types").SpreadsheetRange,
    private fillRanges: import("../types").SpreadsheetRange[],
    private direction: import("@ricsam/selection-manager").FillDirection
  ) {}

  execute(): void {
    // Capture current content and metadata in fill ranges before modification
    for (const fillRange of this.fillRanges) {
      if (fillRange.end.col.type === "infinity" || fillRange.end.row.type === "infinity") {
        continue; // Skip infinite ranges
      }
      
      const startCol = fillRange.start.col;
      const startRow = fillRange.start.row;
      const endCol = fillRange.end.col.value;
      const endRow = fillRange.end.row.value;

      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          const key = getCellReference({ colIndex: col, rowIndex: row });
          const address = {
            workbookName: this.opts.workbookName,
            sheetName: this.opts.sheetName,
            colIndex: col,
            rowIndex: row,
          };
          
          const content = this.workbookManager.getCellContent(address);
          if (content !== undefined) {
            this.previousContent.set(key, content);
          }
          
          const metadata = this.workbookManager.getCellMetadata(address);
          if (metadata !== undefined) {
            this.previousMetadata.set(key, metadata);
          }
        }
      }
    }

    // Execute the auto-fill operation
    this.autoFillManager.fill(this.opts, this.seedRange, this.fillRanges, this.direction);
  }

  undo(): void {
    // Get current sheet content
    const currentContent = this.workbookManager.getSheetSerialized(this.opts);
    const newContent = new Map(currentContent);

    // Restore all fill range cells to their previous state
    for (const fillRange of this.fillRanges) {
      if (fillRange.end.col.type === "infinity" || fillRange.end.row.type === "infinity") {
        continue;
      }
      
      const startCol = fillRange.start.col;
      const startRow = fillRange.start.row;
      const endCol = fillRange.end.col.value;
      const endRow = fillRange.end.row.value;

      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          const key = getCellReference({ colIndex: col, rowIndex: row });
          const address = {
            workbookName: this.opts.workbookName,
            sheetName: this.opts.sheetName,
            colIndex: col,
            rowIndex: row,
          };
          
          // Restore content
          if (this.previousContent.has(key)) {
            newContent.set(key, this.previousContent.get(key)!);
          } else {
            newContent.delete(key);
          }
          
          // Restore metadata
          if (this.previousMetadata.has(key)) {
            this.workbookManager.setCellMetadata(address, this.previousMetadata.get(key));
          } else {
            this.workbookManager.setCellMetadata(address, undefined);
          }
        }
      }
    }

    // Apply restored content
    this.workbookManager.setSheetContent(this.opts, newContent);
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.AUTO_FILL,
      payload: {
        opts: this.opts,
        seedRange: this.seedRange,
        fillRanges: this.fillRanges,
        direction: this.direction,
      },
    };
  }
}

