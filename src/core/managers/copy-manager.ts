/**
 * CopyManager - Manages cell copy/paste operations
 */

import type {
  CellAddress,
  CopyCellsOptions,
  ConditionalStyle,
  DirectCellStyle,
  LocalCellAddress,
  SerializedCellValue,
  SpreadsheetRange,
} from "../types";
import type { WorkbookManager } from "./workbook-manager";
import type { EvaluationManager } from "./evaluation-manager";
import type { StyleManager } from "./style-manager";
import { parseFormula } from "../../parser/parser";
import { astToString } from "../../parser/formatter";
import { transformAST } from "../ast-traverser";
import type { ReferenceNode, RangeNode } from "../../parser/ast";
import { isCellInRange } from "../utils";
import { intersectRanges } from "../utils/range-utils";

export class CopyManager {
  constructor(
    private workbookManager: WorkbookManager,
    private evaluationManager: EvaluationManager,
    private styleManager: StyleManager
  ) {}

  /**
   * Copy cells from source to target
   */
  copyCells(
    source: CellAddress[],
    target: CellAddress,
    options: CopyCellsOptions
  ): void {
    if (source.length === 0) {
      return;
    }

    // Find top-left cell of source (minimum row/col indices)
    const topLeft = this.findTopLeft(source);

    // Calculate offset from source top-left to target
    const rowOffset = target.rowIndex - topLeft.rowIndex;
    const colOffset = target.colIndex - topLeft.colIndex;

    // Copy cell contents
    for (const sourceCell of source) {
      const targetCell: CellAddress = {
        workbookName: target.workbookName,
        sheetName: target.sheetName,
        colIndex: sourceCell.colIndex + colOffset,
        rowIndex: sourceCell.rowIndex + rowOffset,
      };

      this.copyCellContent(sourceCell, targetCell, topLeft, options);
    }

    // Copy formatting if requested
    if (options.formatting) {
      this.copyFormatting(source, topLeft, target, rowOffset, colOffset);
    }

    // Clear source cells if cut
    if (options.cut) {
      this.clearSourceCells(source);
    }
  }

  /**
   * Find the top-left cell (minimum row/col indices)
   */
  private findTopLeft(cells: CellAddress[]): CellAddress {
    let minRow = Infinity;
    let minCol = Infinity;
    let topLeftCell = cells[0]!;

    for (const cell of cells) {
      if (
        cell.rowIndex < minRow ||
        (cell.rowIndex === minRow && cell.colIndex < minCol)
      ) {
        minRow = cell.rowIndex;
        minCol = cell.colIndex;
        topLeftCell = cell;
      }
    }

    return topLeftCell;
  }

  /**
   * Copy content from one cell to another
   */
  private copyCellContent(
    sourceCell: CellAddress,
    targetCell: CellAddress,
    sourceTopLeft: CellAddress,
    options: CopyCellsOptions
  ): void {
    const sheet = this.workbookManager.getSheet({
      workbookName: sourceCell.workbookName,
      sheetName: sourceCell.sheetName,
    });

    if (!sheet) {
      return;
    }

    const key = `${String.fromCharCode(65 + sourceCell.colIndex)}${
      sourceCell.rowIndex + 1
    }`;
    const cellContent = sheet.content.get(key);

    if (!cellContent) {
      // Source cell is empty
      return;
    }

    let targetContent: SerializedCellValue;

    if (options.type === "value") {
      // Copy evaluated value
      const evalResult = this.evaluationManager.getCellEvaluationResult(sourceCell);
      
      if (!evalResult || evalResult.type !== "value") {
        // If evaluation failed or is not a value, copy as-is
        targetContent = cellContent;
      } else {
        // Convert to literal value
        const result = evalResult.result;
        if (result.type === "number") {
          targetContent = result.value;
        } else if (result.type === "string") {
          targetContent = result.value;
        } else if (result.type === "boolean") {
          targetContent = result.value;
        } else {
          // Error or other type, copy as-is
          targetContent = cellContent;
        }
      }
    } else {
      // Copy formula
      if (typeof cellContent === "string" && cellContent.startsWith("=")) {
        // Adjust formula references
        targetContent = this.adjustFormulaReferences(
          cellContent,
          {
            colIndex: sourceCell.colIndex,
            rowIndex: sourceCell.rowIndex,
          },
          {
            colIndex: targetCell.colIndex,
            rowIndex: targetCell.rowIndex,
          }
        );
      } else {
        // Not a formula, copy as-is
        targetContent = cellContent;
      }
    }

    // Set target cell content (using the engine's method through workbook manager)
    const targetSheet = this.workbookManager.getSheet({
      workbookName: targetCell.workbookName,
      sheetName: targetCell.sheetName,
    });

    if (targetSheet) {
      const targetKey = `${String.fromCharCode(65 + targetCell.colIndex)}${
        targetCell.rowIndex + 1
      }`;
      targetSheet.content.set(targetKey, targetContent);
    }
  }

  /**
   * Adjust formula references when copying
   * Based on autofill-utils.ts adjustFormulaReferences
   */
  private adjustFormulaReferences(
    formula: string,
    sourceAddress: LocalCellAddress,
    targetAddress: LocalCellAddress
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

  /**
   * Copy formatting (cellStyles and conditionalStyles) from source to target
   */
  private copyFormatting(
    sourceCells: CellAddress[],
    sourceTopLeft: CellAddress,
    target: CellAddress,
    rowOffset: number,
    colOffset: number
  ): void {
    // Get all styles for the source workbook
    const allConditionalStyles = this.styleManager.getAllConditionalStyles();
    const allCellStyles = this.styleManager.getAllCellStyles();

    // Find styles that intersect with source cells
    const sourceRange = this.getBoundingBox(sourceCells);

    // Copy conditional styles
    for (const style of allConditionalStyles) {
      if (
        style.area.workbookName === sourceTopLeft.workbookName &&
        style.area.sheetName === sourceTopLeft.sheetName
      ) {
        // Calculate intersection of style range with source bounding box
        const intersection = intersectRanges(style.area.range, sourceRange);
        if (intersection) {
          // Copy only the intersection, offset to target
          const newStyle: ConditionalStyle = {
            area: {
              workbookName: target.workbookName,
              sheetName: target.sheetName,
              range: this.adjustRange(intersection, rowOffset, colOffset),
            },
            condition: style.condition,
          };
          this.styleManager.addConditionalStyle(newStyle);
        }
      }
    }

    // Copy cell styles
    for (const style of allCellStyles) {
      if (
        style.area.workbookName === sourceTopLeft.workbookName &&
        style.area.sheetName === sourceTopLeft.sheetName
      ) {
        // Calculate intersection of style range with source bounding box
        const intersection = intersectRanges(style.area.range, sourceRange);
        if (intersection) {
          // Copy only the intersection, offset to target
          const newStyle: DirectCellStyle = {
            area: {
              workbookName: target.workbookName,
              sheetName: target.sheetName,
              range: this.adjustRange(intersection, rowOffset, colOffset),
            },
            style: style.style,
          };
          this.styleManager.addCellStyle(newStyle);
        }
      }
    }
  }

  /**
   * Get bounding box for a set of cells
   */
  private getBoundingBox(cells: CellAddress[]): SpreadsheetRange {
    let minRow = Infinity;
    let maxRow = -Infinity;
    let minCol = Infinity;
    let maxCol = -Infinity;

    for (const cell of cells) {
      minRow = Math.min(minRow, cell.rowIndex);
      maxRow = Math.max(maxRow, cell.rowIndex);
      minCol = Math.min(minCol, cell.colIndex);
      maxCol = Math.max(maxCol, cell.colIndex);
    }

    return {
      start: { col: minCol, row: minRow },
      end: {
        col: { type: "number", value: maxCol },
        row: { type: "number", value: maxRow },
      },
    };
  }


  /**
   * Adjust a range by row and column offsets
   */
  private adjustRange(
    range: SpreadsheetRange,
    rowOffset: number,
    colOffset: number
  ): SpreadsheetRange {
    return {
      start: {
        col: range.start.col + colOffset,
        row: range.start.row + rowOffset,
      },
      end: {
        col:
          range.end.col.type === "number"
            ? { type: "number", value: range.end.col.value + colOffset }
            : range.end.col,
        row:
          range.end.row.type === "number"
            ? { type: "number", value: range.end.row.value + rowOffset }
            : range.end.row,
      },
    };
  }

  /**
   * Clear source cells (for cut operation)
   */
  private clearSourceCells(cells: CellAddress[]): void {
    for (const cell of cells) {
      const sheet = this.workbookManager.getSheet({
        workbookName: cell.workbookName,
        sheetName: cell.sheetName,
      });

      if (sheet) {
        const key = `${String.fromCharCode(65 + cell.colIndex)}${
          cell.rowIndex + 1
        }`;
        sheet.content.delete(key);
      }
    }
  }
}

