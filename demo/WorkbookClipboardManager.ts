import { SelectionManager, type SMArea } from "@ricsam/selection-manager";
import { FormulaEngine } from "../src/core/engine";
import type {
  CellAddress,
  SerializedCellValue,
  SpreadsheetRange,
} from "../src/core/types";
import { parseCellReference } from "../src/core/utils";
import { indexToColumn } from "../src/core/utils";
import { getCellReference } from "../src/core/utils";

export type CellDataUpdate = {
  rowIndex: number;
  colIndex: number;
  value: string;
};

export class ClipboardUtils {
  public extractCellsFromSelection(
    selectionManager: SelectionManager,
    cellData: Map<string, SerializedCellValue>
  ) {
    const boundingRect = selectionManager.getSelectionsBoundingRect();
    if (!boundingRect) return;

    const endRow = boundingRect.end.row;
    const endCol = boundingRect.end.col;

    // 🧮 Create a proper grid for export
    let height: number;
    let width: number;
    const startRow = boundingRect.start.row;
    const startCol = boundingRect.start.col;

    // Handle infinity cases - read all data and determine bounds
    if (endRow.type === "infinity" || endCol.type === "infinity") {
      // Find the actual bounds from the cellData
      let maxRow = 0;
      let maxCol = 0;

      cellData.forEach((value, key) => {
        const { rowIndex, colIndex } = parseCellReference(key);

        // Only consider cells within the starting bounds
        if (rowIndex >= startRow && colIndex >= startCol) {
          // Apply infinity constraints
          const withinRowBounds =
            endRow.type === "infinity" || rowIndex <= endRow.value;
          const withinColBounds =
            endCol.type === "infinity" || colIndex <= endCol.value;

          if (withinRowBounds && withinColBounds) {
            maxRow = Math.max(maxRow, rowIndex);
            maxCol = Math.max(maxCol, colIndex);
          }
        }
      });

      // If no data found, create minimal grid
      if (maxRow === 0 && maxCol === 0) {
        height = 1;
        width = 1;
      } else {
        height = maxRow - startRow + 1;
        width = maxCol - startCol + 1;
      }
    } else {
      // Both finite - use original logic
      height = endRow.value - startRow + 1;
      width = endCol.value - startCol + 1;
    }

    const cells: {
      relative: { rowIndex: number; columnIndex: number };
      absolute: { rowIndex: number; columnIndex: number };
      value: SerializedCellValue;
      key: string;
    }[] = [];

    // Fill the grid with data
    if (endRow.type === "infinity" || endCol.type === "infinity") {
      // For infinity cases, iterate through cellData
      cellData.forEach((value, key) => {
        const { rowIndex, colIndex } = parseCellReference(key);

        // Check if this cell should be included
        if (rowIndex >= startRow && colIndex >= startCol) {
          const withinRowBounds =
            endRow.type === "infinity" || rowIndex <= endRow.value;
          const withinColBounds =
            endCol.type === "infinity" || colIndex <= endCol.value;

          if (withinRowBounds && withinColBounds) {
            const gridRow = rowIndex - startRow;
            const gridCol = colIndex - startCol;

            if (gridRow < height && gridCol < width) {
              cells.push({
                relative: { rowIndex: gridRow, columnIndex: gridCol },
                absolute: { rowIndex: rowIndex, columnIndex: colIndex },
                value,
                key,
              });
            }
          }
        }
      });
    } else {
      // For finite cases, use the existing forEachSelectedCell logic
      selectionManager.forEachSelectedCell(({ absolute, relative }) => {
        const key = `${indexToColumn(absolute.col)}${absolute.row + 1}`;
        const value = cellData.get(key) ?? "";
        cells.push({
          relative: { rowIndex: relative.row, columnIndex: relative.col },
          absolute: { rowIndex: absolute.row, columnIndex: absolute.col },
          value,
          key,
        });
      });
    }

    return { width, height, cells };
  }

  public writeToOsClipboard(grid: string[][]) {
    const tsvString = grid.map((row) => row.join("\t")).join("\n");
    navigator.clipboard.writeText(tsvString);
    return tsvString;
  }

  public createExportGrid(width: number, height: number) {
    return Array(height)
      .fill(null)
      .map(() => Array(width).fill(""));
  }
}

export class WorkbookClipboardManager extends ClipboardUtils {
  constructor(private engine: FormulaEngine) {
    super();
  }
  copiedCells: CellAddress[] = [];
  signature: string = "";

  public triggerCopy(context: {
    workbookName: string;
    sheetName: string;
    selectionManager: SelectionManager;
    copyType: "value" | "formula";
  }): void {
    const cellData = this.engine.getSheet({
      workbookName: context.workbookName,
      sheetName: context.sheetName,
    })?.content;
    if (!cellData) return;
    const extractedCells = this.extractCellsFromSelection(
      context.selectionManager,
      cellData
    );
    if (!extractedCells) return;
    const { width, height, cells } = extractedCells;
    const exportGrid = this.createExportGrid(width, height);
    this.copiedCells = [];
    const updates: CellDataUpdate[] = [];
    cells.forEach(({ relative, absolute }) => {
      const cellAddress: CellAddress = {
        workbookName: context.workbookName,
        sheetName: context.sheetName,
        colIndex: absolute.columnIndex,
        rowIndex: absolute.rowIndex,
      };
      this.copiedCells.push(cellAddress);
      const value =
        context.copyType === "value"
          ? this.engine.getCellValue(cellAddress, false)
          : cellData.get(getCellReference(cellAddress));
      exportGrid[relative.rowIndex]![relative.columnIndex] = value;
      updates.push({
        rowIndex: relative.rowIndex,
        colIndex: relative.columnIndex,
        value: value?.toString() ?? "",
      });
    });
    this.signature = this.writeToOsClipboard(exportGrid);
  }
  public triggerPaste(context: {
    workbookName: string;
    sheetName: string;
    selectionManager: SelectionManager;
    updates: CellDataUpdate[];
    rawString: string;
    pasteType: "value" | "formula";
  }): void {
    if (context.rawString === this.signature) {
      // Internal paste operation - use smartPaste to handle both copy and fill
      const selections = context.selectionManager.selections;
      if (!selections || selections.length === 0) return;

      // Convert each SMArea to SpreadsheetRange
      const convertSMAreaToSpreadsheetRange = (
        area: SMArea
      ): SpreadsheetRange => {
        return {
          start: {
            col: area.start.col,
            row: area.start.row,
          },
          end: {
            col:
              area.end.col.type === "infinity"
                ? { type: "infinity" as const, sign: "positive" as const }
                : { type: "number" as const, value: area.end.col.value },
            row:
              area.end.row.type === "infinity"
                ? { type: "infinity" as const, sign: "positive" as const }
                : { type: "number" as const, value: area.end.row.value },
          },
        };
      };

      const areas = selections.map(convertSMAreaToSpreadsheetRange);

      this.engine.smartPaste(
        this.copiedCells,
        {
          workbookName: context.workbookName,
          sheetName: context.sheetName,
          areas,
        },
        {
          cut: false,
          type: context.pasteType,
          target: "all",
        }
      );
    } else {
      // External paste operation
      context.selectionManager.saveCellValues(context.updates);
    }
  }
}
