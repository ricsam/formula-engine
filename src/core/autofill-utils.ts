/**
 * AutoFill class for handling spreadsheet autofill functionality
 */

import type {
  CellAddress,
  SerializedCellValue,
  SpreadsheetRange,
  FiniteSpreadsheetRange,
} from "./types";
import type { FillDirection } from "@ricsam/selection-manager";
import { parseFormula } from "../parser/parser";
import { astToString } from "../parser/formatter";
import { transformAST } from "./ast-traverser";
import type { ReferenceNode, RangeNode } from "../parser/ast";
import { getCellReference } from "./utils";
import type { SheetManager } from "./managers/sheet-manager";

export class AutoFill {
  constructor(
    private sheetManager: SheetManager,
    private engine: {
      setCellContent: (
        address: CellAddress,
        content: SerializedCellValue
      ) => void;
      setSheetContent: (
        sheetName: string,
        content: Map<string, SerializedCellValue>
      ) => void;
    }
  ) {}

  /**
   * Converts a SpreadsheetRange to FiniteSpreadsheetRange, throwing an error if infinite
   */
  private toFiniteRange(range: SpreadsheetRange): FiniteSpreadsheetRange {
    if (range.end.col.type === "infinity" || range.end.row.type === "infinity") {
      throw new Error("AutoFill with infinite ranges is not supported");
    }
    
    return {
      start: range.start,
      end: {
        col: range.end.col.value,
        row: range.end.row.value,
      },
    };
  }

  private getCellContent(address: CellAddress): SerializedCellValue {
    const sheet = this.sheetManager.getSheets().get(address.sheetName);
    return sheet?.content.get(getCellReference(address)) || undefined;
  }

  fill(
    sheetName: string,
    seedRange: SpreadsheetRange,
    fillRange: SpreadsheetRange,
    direction: FillDirection
  ) {
    // Convert to finite ranges (throws error if infinite)
    const finiteSeedRange = this.toFiniteRange(seedRange);
    const finiteFillRange = this.toFiniteRange(fillRange);

    // Get seed cells data
    const seedCells = this.getSeedCells(sheetName, finiteSeedRange);

    // Collect all changes to apply in a single batch
    const changes = new Map<string, SerializedCellValue>();

    // Determine if we have a single cell or multi-cell seed
    const isSingleCell = seedCells.length === 1;

    if (isSingleCell) {
      const seedCell = seedCells[0];
      if (seedCell) {
        this.collectSingleCellPattern(seedCell, finiteFillRange, direction, changes);
      }
    } else {
      this.collectMultiCellPattern(
        seedCells,
        finiteSeedRange,
        finiteFillRange,
        direction,
        changes
      );
    }

    // Get current sheet content and merge with changes
    const currentContent = this.sheetManager.getSheetSerialized(sheetName);
    const newContent = new Map(currentContent);

    // Apply all changes
    changes.forEach((value, key) => {
      if (value === undefined) {
        newContent.delete(key);
      } else {
        newContent.set(key, value);
      }
    });

    // Update sheet content in a single operation
    this.engine.setSheetContent(sheetName, newContent);
  }

  private getSeedCells(sheetName: string, seedRange: FiniteSpreadsheetRange) {
    const cells: Array<{
      address: CellAddress;
      content: SerializedCellValue;
      rowOffset: number;
      colOffset: number;
    }> = [];

    const startCol = seedRange.start.col;
    const startRow = seedRange.start.row;
    const endCol = seedRange.end.col;
    const endRow = seedRange.end.row;

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const address: CellAddress = {
          sheetName,
          colIndex: col,
          rowIndex: row,
        };
        const content = this.getCellContent(address);
        cells.push({
          address,
          content,
          rowOffset: row - startRow,
          colOffset: col - startCol,
        });
      }
    }

    return cells;
  }

  private collectSingleCellPattern(
    seedCell: { address: CellAddress; content: SerializedCellValue },
    fillRange: FiniteSpreadsheetRange,
    direction: FillDirection,
    changes: Map<string, SerializedCellValue>
  ) {
    const { content } = seedCell;

    const startCol = fillRange.start.col;
    const startRow = fillRange.start.row;
    const endCol = fillRange.end.col;
    const endRow = fillRange.end.row;

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const targetAddress: CellAddress = {
          sheetName: seedCell.address.sheetName,
          colIndex: col,
          rowIndex: row,
        };

        let newContent: SerializedCellValue;

        if (content === undefined || content === "") {
          // Blank cell - fills blanks (effectively clears targets)
          newContent = undefined;
        } else if (typeof content === "string" && content.startsWith("=")) {
          // Formula - adjust relative references
          newContent = this.adjustFormulaReferences(
            content,
            seedCell.address,
            targetAddress
          );
        } else {
          // Number or text - copy by default
          newContent = content;
        }

        changes.set(getCellReference(targetAddress), newContent);
      }
    }
  }

  private collectMultiCellPattern(
    seedCells: Array<{
      address: CellAddress;
      content: SerializedCellValue;
      rowOffset: number;
      colOffset: number;
    }>,
    seedRange: FiniteSpreadsheetRange,
    fillRange: FiniteSpreadsheetRange,
    direction: FillDirection,
    changes: Map<string, SerializedCellValue>
  ) {
    // Try to infer linear step from numbers
    const step = this.inferLinearStep(seedCells, direction);

    const startCol = fillRange.start.col;
    const startRow = fillRange.start.row;
    const endCol = fillRange.end.col;
    const endRow = fillRange.end.row;

    const seedWidth = seedRange.end.col - seedRange.start.col + 1;
    const seedHeight = seedRange.end.row - seedRange.start.row + 1;

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const targetAddress: CellAddress = {
          sheetName: seedCells[0]?.address.sheetName || "",
          colIndex: col,
          rowIndex: row,
        };

        let newContent: SerializedCellValue;

        if (step !== null) {
          // Use linear progression
          newContent = this.applyLinearStep(
            seedCells,
            targetAddress,
            seedRange,
            step,
            direction
          );
        } else {
          // Repeat pattern (cycle through seed block)
          const seedColOffset = (col - startCol) % seedWidth;
          const seedRowOffset = (row - startRow) % seedHeight;

          const seedCell = seedCells.find(
            (cell) =>
              cell.colOffset === seedColOffset &&
              cell.rowOffset === seedRowOffset
          );

          if (seedCell) {
            if (
              typeof seedCell.content === "string" &&
              seedCell.content.startsWith("=")
            ) {
              // Formula - adjust relative references
              newContent = this.adjustFormulaReferences(
                seedCell.content,
                seedCell.address,
                targetAddress
              );
            } else {
              newContent = seedCell.content;
            }
          } else {
            newContent = undefined;
          }
        }

        changes.set(getCellReference(targetAddress), newContent);
      }
    }
  }

  private inferLinearStep(
    seedCells: Array<{
      content: SerializedCellValue;
      rowOffset: number;
      colOffset: number;
    }>,
    direction: FillDirection
  ): number | null {
    // Extract numeric values in the direction of fill
    const values: number[] = [];

    if (direction === "down" || direction === "up") {
      // Look at first column, different rows
      const firstColCells = seedCells
        .filter((cell) => cell.colOffset === 0)
        .sort((a, b) => a.rowOffset - b.rowOffset);

      for (const cell of firstColCells) {
        if (typeof cell.content === "number") {
          values.push(cell.content);
        } else if (typeof cell.content === "string") {
          const num = parseFloat(cell.content);
          if (!isNaN(num)) {
            values.push(num);
          } else {
            return null; // Non-numeric content, can't infer step
          }
        } else {
          return null;
        }
      }
    } else {
      // Look at first row, different columns
      const firstRowCells = seedCells
        .filter((cell) => cell.rowOffset === 0)
        .sort((a, b) => a.colOffset - b.colOffset);

      for (const cell of firstRowCells) {
        if (typeof cell.content === "number") {
          values.push(cell.content);
        } else if (typeof cell.content === "string") {
          const num = parseFloat(cell.content);
          if (!isNaN(num)) {
            values.push(num);
          } else {
            return null; // Non-numeric content, can't infer step
          }
        } else {
          return null;
        }
      }
    }

    if (values.length < 2) {
      return null;
    }

    // Check if there's a consistent step
    const step = values[1]! - values[0]!;
    for (let i = 2; i < values.length; i++) {
      if (Math.abs(values[i]! - values[i - 1]! - step) > 1e-10) {
        return null; // Inconsistent step
      }
    }

    return step;
  }

  private applyLinearStep(
    seedCells: Array<{
      content: SerializedCellValue;
      rowOffset: number;
      colOffset: number;
    }>,
    targetAddress: CellAddress,
    seedRange: FiniteSpreadsheetRange,
    step: number,
    direction: FillDirection
  ): SerializedCellValue {
    // Calculate the position in the overall sequence
    let positionInSequence: number;
    let baseValue: number;

    if (direction === "down") {
      positionInSequence = targetAddress.rowIndex - seedRange.start.row;

      // Find the base value from the first row of the seed
      const baseCell = seedCells.find(
        (cell) =>
          cell.colOffset === targetAddress.colIndex - seedRange.start.col &&
          cell.rowOffset === 0
      );

      if (!baseCell) return undefined;
      baseValue =
        typeof baseCell.content === "number"
          ? baseCell.content
          : parseFloat(String(baseCell.content || "0"));
    } else if (direction === "up") {
      positionInSequence = targetAddress.rowIndex - seedRange.start.row;

      // Find the base value from the first row of the seed (same as down direction)
      const baseCell = seedCells.find(
        (cell) =>
          cell.colOffset === targetAddress.colIndex - seedRange.start.col &&
          cell.rowOffset === 0
      );

      if (!baseCell) return undefined;
      baseValue =
        typeof baseCell.content === "number"
          ? baseCell.content
          : parseFloat(String(baseCell.content || "0"));
    } else if (direction === "right") {
      positionInSequence = targetAddress.colIndex - seedRange.start.col;

      // Find the base value from the first column of the seed
      const baseCell = seedCells.find(
        (cell) =>
          cell.rowOffset === targetAddress.rowIndex - seedRange.start.row &&
          cell.colOffset === 0
      );

      if (!baseCell) return undefined;
      baseValue =
        typeof baseCell.content === "number"
          ? baseCell.content
          : parseFloat(String(baseCell.content || "0"));
    } else {
      // left
      positionInSequence = targetAddress.colIndex - seedRange.start.col;

      // Find the base value from the first column of the seed (same as right direction)
      const baseCell = seedCells.find(
        (cell) =>
          cell.rowOffset === targetAddress.rowIndex - seedRange.start.row &&
          cell.colOffset === 0
      );

      if (!baseCell) return undefined;
      baseValue =
        typeof baseCell.content === "number"
          ? baseCell.content
          : parseFloat(String(baseCell.content || "0"));
    }

    if (isNaN(baseValue)) {
      return undefined;
    }

    const result = baseValue + step * positionInSequence;
    return typeof seedCells[0]?.content === "string"
      ? result.toString()
      : result;
  }

  private adjustFormulaReferences(
    formula: string,
    sourceAddress: CellAddress,
    targetAddress: CellAddress
  ): string {
    try {
      const ast = parseFormula(formula.slice(1)); // Remove the "=" sign

      const rowDelta = targetAddress.rowIndex - sourceAddress.rowIndex;
      const colDelta = targetAddress.colIndex - sourceAddress.colIndex;

      const adjustedAst = transformAST(ast, (node) => {
        if (node.type === "reference") {
          const refNode = node as ReferenceNode;
          return {
            ...refNode,
            address: {
              colIndex: refNode.isAbsolute.col
                ? refNode.address.colIndex
                : refNode.address.colIndex + colDelta,
              rowIndex: refNode.isAbsolute.row
                ? refNode.address.rowIndex
                : refNode.address.rowIndex + rowDelta,
            },
          };
        } else if (node.type === "range") {
          const rangeNode = node as RangeNode;
          return {
            ...rangeNode,
            range: {
              start: {
                col: rangeNode.isAbsolute.start.col
                  ? rangeNode.range.start.col
                  : rangeNode.range.start.col + colDelta,
                row: rangeNode.isAbsolute.start.row
                  ? rangeNode.range.start.row
                  : rangeNode.range.start.row + rowDelta,
              },
              end: {
                col:
                  rangeNode.range.end.col.type === "number"
                    ? rangeNode.isAbsolute.end.col
                      ? rangeNode.range.end.col
                      : {
                          type: "number" as const,
                          value: rangeNode.range.end.col.value + colDelta,
                        }
                    : rangeNode.range.end.col,
                row:
                  rangeNode.range.end.row.type === "number"
                    ? rangeNode.isAbsolute.end.row
                      ? rangeNode.range.end.row
                      : {
                          type: "number" as const,
                          value: rangeNode.range.end.row.value + rowDelta,
                        }
                    : rangeNode.range.end.row,
              },
            },
          };
        }
        return node;
      });

      return `=${astToString(adjustedAst)}`;
    } catch (error) {
      // If parsing fails, return the original formula
      console.warn("Failed to adjust formula references:", error);
      return formula;
    }
  }
}
