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
import type {
  EngineCommand,
  EngineAction,
  MutationInvalidation,
} from "./types";
import {
  ActionTypes,
  emptyMutationInvalidation,
  getSerializedCellValueKind,
} from "./types";
import { getCellReference, parseCellReference } from "../utils";

function buildTouchedCells(
  cells: Array<{
    address: CellAddress;
    before: SerializedCellValue | undefined;
    after: SerializedCellValue | undefined;
  }>
): MutationInvalidation["touchedCells"] {
  const deduped = new Map<
    string,
    {
      address: CellAddress;
      beforeKind: ReturnType<typeof getSerializedCellValueKind>;
      afterKind: ReturnType<typeof getSerializedCellValueKind>;
    }
  >();

  for (const cell of cells) {
    deduped.set(`${cell.address.workbookName}:${cell.address.sheetName}:${getCellReference(cell.address)}`, {
      address: cell.address,
      beforeKind: getSerializedCellValueKind(cell.before),
      afterKind: getSerializedCellValueKind(cell.after),
    });
  }

  return Array.from(deduped.values());
}

function getAddressKey(address: CellAddress): string {
  return `${address.workbookName}:${address.sheetName}:${getCellReference(address)}`;
}

function captureCellContents(
  workbookManager: WorkbookManager,
  addresses: CellAddress[]
): Map<string, SerializedCellValue | undefined> {
  const contents = new Map<string, SerializedCellValue | undefined>();
  for (const address of addresses) {
    contents.set(getAddressKey(address), workbookManager.getCellContent(address));
  }
  return contents;
}

/**
 * Command to set a single cell's content.
 */
export class SetCellContentCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private previousValue: SerializedCellValue | undefined;
  private hadPreviousValue = false;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

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
    this.executeFootprint = {
      touchedCells: buildTouchedCells([
        {
          address: this.address,
          before: this.previousValue,
          after: this.newContent,
        },
      ]),
      resourceKeys: [],
    };
    this.undoFootprint = {
      touchedCells: buildTouchedCells([
        {
          address: this.address,
          before: this.newContent,
          after: this.hadPreviousValue ? this.previousValue : undefined,
        },
      ]),
      resourceKeys: [],
    };
  }

  undo(): void {
    if (this.hadPreviousValue) {
      this.workbookManager.setCellContent(this.address, this.previousValue);
    } else {
      this.workbookManager.setCellContent(this.address, undefined);
    }
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

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
    const touchedKeys = new Set<string>([
      ...Array.from(this.previousContent?.keys() ?? []),
      ...Array.from(this.newContent.keys()),
    ]);
    const touchedCells = buildTouchedCells(
      Array.from(touchedKeys, (key) => ({
        address: {
          workbookName: this.opts.workbookName,
          sheetName: this.opts.sheetName,
          ...parseCellReference(key),
        },
        before: this.previousContent?.get(key),
        after: this.newContent.get(key),
      }))
    );
    this.executeFootprint = {
      touchedCells,
      resourceKeys: [],
    };
    this.undoFootprint = {
      touchedCells: buildTouchedCells(
        Array.from(touchedKeys, (key) => ({
          address: {
            workbookName: this.opts.workbookName,
            sheetName: this.opts.sheetName,
            ...parseCellReference(key),
          },
          before: this.newContent.get(key),
          after: this.previousContent?.get(key),
        }))
      ),
      resourceKeys: [],
    };
  }

  undo(): void {
    if (this.previousContent) {
      this.workbookManager.setSheetContent(this.opts, this.previousContent);
    }
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private workbookManager: WorkbookManager,
    private address: RangeAddress
  ) {}

  execute(): void {
    this.clearedCells.clear();
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
    this.executeFootprint = {
      touchedCells: buildTouchedCells(
        Array.from(this.clearedCells.entries(), ([key, before]) => ({
          address: {
            workbookName: this.address.workbookName,
            sheetName: this.address.sheetName,
            ...parseCellReference(key),
          },
          before,
          after: undefined,
        }))
      ),
      resourceKeys: [],
    };
    this.undoFootprint = {
      touchedCells: buildTouchedCells(
        Array.from(this.clearedCells.entries(), ([key, after]) => ({
          address: {
            workbookName: this.address.workbookName,
            sheetName: this.address.sheetName,
            ...parseCellReference(key),
          },
          before: undefined,
          after,
        }))
      ),
      resourceKeys: [],
    };
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

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();
  private executedContents: Map<string, SerializedCellValue | undefined> =
    new Map();

  constructor(
    private workbookManager: WorkbookManager,
    private copyManager: CopyManager,
    private source: CellAddress[],
    private target: CellAddress,
    private options: CopyCellsOptions
  ) {}

  execute(): void {
    this.targetSnapshots = [];
    this.sourceSnapshots = [];
    this.executedContents.clear();
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
    const touchedAddresses = [
      ...this.targetSnapshots.map((snapshot) => snapshot.address),
      ...this.sourceSnapshots.map((snapshot) => snapshot.address),
    ];
    this.executedContents = captureCellContents(this.workbookManager, touchedAddresses);
    this.executeFootprint = {
      touchedCells: buildTouchedCells(
        touchedAddresses.map((address) => {
          const beforeSnapshot =
            this.targetSnapshots.find(
              (snapshot) =>
                snapshot.address.workbookName === address.workbookName &&
                snapshot.address.sheetName === address.sheetName &&
                snapshot.address.colIndex === address.colIndex &&
                snapshot.address.rowIndex === address.rowIndex
            ) ??
            this.sourceSnapshots.find(
              (snapshot) =>
                snapshot.address.workbookName === address.workbookName &&
                snapshot.address.sheetName === address.sheetName &&
                snapshot.address.colIndex === address.colIndex &&
                snapshot.address.rowIndex === address.rowIndex
            );
          return {
            address,
            before: beforeSnapshot?.content,
            after: this.executedContents.get(getAddressKey(address)),
          };
        })
      ),
      resourceKeys: [],
    };
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

    const touchedAddresses = [
      ...this.targetSnapshots.map((snapshot) => snapshot.address),
      ...this.sourceSnapshots.map((snapshot) => snapshot.address),
    ];
    this.undoFootprint = {
      touchedCells: buildTouchedCells(
        touchedAddresses.map((address) => {
          return {
            address,
            before: this.executedContents.get(getAddressKey(address)),
            after: this.workbookManager.getCellContent(address),
          };
        })
      ),
      resourceKeys: [],
    };
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();
  private executedContents: Map<string, SerializedCellValue | undefined> =
    new Map();

  constructor(
    private workbookManager: WorkbookManager,
    private copyManager: CopyManager,
    private seedRange: RangeAddress,
    private targetRanges: RangeAddress[],
    private options: CopyCellsOptions
  ) {}

  execute(): void {
    this.targetSnapshots.clear();
    this.executedContents.clear();
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
    const touchedAddresses = Array.from(this.targetSnapshots.values()).map(
      (snapshot) => snapshot.address
    );
    this.executedContents = captureCellContents(this.workbookManager, touchedAddresses);
    this.executeFootprint = {
      touchedCells: buildTouchedCells(
        touchedAddresses.map((address) => ({
          address,
          before: this.targetSnapshots.get(getAddressKey(address))?.content,
          after: this.executedContents.get(getAddressKey(address)),
        }))
      ),
      resourceKeys: [],
    };
  }

  undo(): void {
    // Restore all target cells
    for (const snapshot of this.targetSnapshots.values()) {
      this.workbookManager.setCellContent(snapshot.address, snapshot.content);
      if (snapshot.metadata !== undefined) {
        this.workbookManager.setCellMetadata(snapshot.address, snapshot.metadata);
      }
    }

    const touchedAddresses = Array.from(this.targetSnapshots.values()).map(
      (snapshot) => snapshot.address
    );
    this.undoFootprint = {
      touchedCells: buildTouchedCells(
        touchedAddresses.map((address) => ({
          address,
          before: this.executedContents.get(getAddressKey(address)),
          after: this.workbookManager.getCellContent(address),
        }))
      ),
      resourceKeys: [],
    };
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();
  private executedContents: Map<string, SerializedCellValue | undefined> =
    new Map();

  constructor(
    private workbookManager: WorkbookManager,
    private copyManager: CopyManager,
    private source: CellAddress,
    private target: CellAddress
  ) {}

  execute(): void {
    this.sourceSnapshot = undefined;
    this.targetSnapshot = undefined;
    this.executedContents.clear();
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
    const touchedAddresses = [this.source, this.target];
    this.executedContents = captureCellContents(this.workbookManager, touchedAddresses);
    this.executeFootprint = {
      touchedCells: buildTouchedCells(
        touchedAddresses.map((address) => ({
          address,
          before:
            this.sourceSnapshot?.address === address
              ? this.sourceSnapshot.content
              : this.targetSnapshot?.content,
          after: this.executedContents.get(getAddressKey(address)),
        }))
      ),
      resourceKeys: [],
    };
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

    const touchedAddresses = [this.source, this.target];
    this.undoFootprint = {
      touchedCells: buildTouchedCells(
        touchedAddresses.map((address) => ({
          address,
          before: this.executedContents.get(getAddressKey(address)),
          after: this.workbookManager.getCellContent(address),
        }))
      ),
      resourceKeys: [],
    };
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();
  private executedContents: Map<string, SerializedCellValue | undefined> =
    new Map();

  constructor(
    private workbookManager: WorkbookManager,
    private copyManager: CopyManager,
    private sourceRange: RangeAddress,
    private target: CellAddress
  ) {}

  execute(): void {
    this.sourceSnapshots = [];
    this.targetSnapshots = [];
    this.executedContents.clear();
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
    const touchedAddresses = [
      ...this.sourceSnapshots.map((snapshot) => snapshot.address),
      ...this.targetSnapshots.map((snapshot) => snapshot.address),
    ];
    this.executedContents = captureCellContents(this.workbookManager, touchedAddresses);
    this.executeFootprint = {
      touchedCells: buildTouchedCells(
        touchedAddresses.map((address) => {
          const snapshot =
            this.sourceSnapshots.find(
              (candidate) => getAddressKey(candidate.address) === getAddressKey(address)
            ) ??
            this.targetSnapshots.find(
              (candidate) => getAddressKey(candidate.address) === getAddressKey(address)
            );
          return {
            address,
            before: snapshot?.content,
            after: this.executedContents.get(getAddressKey(address)),
          };
        })
      ),
      resourceKeys: [],
    };
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

    const touchedAddresses = [
      ...this.sourceSnapshots.map((snapshot) => snapshot.address),
      ...this.targetSnapshots.map((snapshot) => snapshot.address),
    ];
    this.undoFootprint = {
      touchedCells: buildTouchedCells(
        touchedAddresses.map((address) => ({
          address,
          before: this.executedContents.get(getAddressKey(address)),
          after: this.workbookManager.getCellContent(address),
        }))
      ),
      resourceKeys: [],
    };
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();
  private touchedAddresses: CellAddress[] = [];
  private executedContents: Map<string, SerializedCellValue | undefined> =
    new Map();

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
    this.previousContent.clear();
    this.previousMetadata.clear();
    this.touchedAddresses = [];
    this.executedContents.clear();
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
          this.touchedAddresses.push(address);
          
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
    this.executedContents = captureCellContents(this.workbookManager, this.touchedAddresses);
    this.executeFootprint = {
      touchedCells: buildTouchedCells(
        this.touchedAddresses.map((address) => {
          const key = getCellReference(address);
          return {
            address,
            before: this.previousContent.get(key),
            after: this.executedContents.get(getAddressKey(address)),
          };
        })
      ),
      resourceKeys: [],
    };
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
    this.undoFootprint = {
      touchedCells: buildTouchedCells(
        this.touchedAddresses.map((address) => ({
          address,
          before: this.executedContents.get(getAddressKey(address)),
          after: this.workbookManager.getCellContent(address),
        }))
      ),
      resourceKeys: [],
    };
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
