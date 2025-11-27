/**
 * CopyManager - Manages cell copy/paste operations
 */

import type {
  CellAddress,
  CopyCellsOptions,
  ConditionalStyle,
  DirectCellStyle,
  LocalCellAddress,
  RangeAddress,
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
   * Paste cells from source to target
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

    // Copy cell contents (skip if only copying style)
    if (options.target !== 'style') {
      for (const sourceCell of source) {
        const targetCell: CellAddress = {
          workbookName: target.workbookName,
          sheetName: target.sheetName,
          colIndex: sourceCell.colIndex + colOffset,
          rowIndex: sourceCell.rowIndex + rowOffset,
        };

        this.copyCellContent(sourceCell, targetCell, topLeft, options);
      }
    }

    // Copy formatting if requested
    if (options.target === 'all' || options.target === 'style') {
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
   * Clears existing cell styles in target range first (Excel behavior)
   */
  private copyFormatting(
    sourceCells: CellAddress[],
    sourceTopLeft: CellAddress,
    target: CellAddress,
    rowOffset: number,
    colOffset: number
  ): void {
    // STEP 1: Clear existing cell styles in target range (Excel-like replacement)
    const sourceRange = this.getBoundingBox(sourceCells);
    const targetRange: RangeAddress = {
      workbookName: target.workbookName,
      sheetName: target.sheetName,
      range: this.adjustRange(sourceRange, rowOffset, colOffset),
    };
    
    this.styleManager.clearCellStylesInRange(targetRange);

    // Get all styles for the source workbook
    const allConditionalStyles = this.styleManager.getAllConditionalStyles();
    const allCellStyles = this.styleManager.getAllCellStyles();

    // STEP 2: Copy conditional styles
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

    // STEP 3: Copy cell styles
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

  /**
   * Fill one or more areas with a seed range's content/style
   * Uses column-first strategy: fills down, then replicates right
   * Formulas are adjusted based on each target cell's offset from the seed
   */
  fillAreas(
    seedRange: RangeAddress,
    targetRanges: RangeAddress[],
    options: CopyCellsOptions
  ): void {
    for (const targetRange of targetRanges) {
      this.fillRangeWithSeed(seedRange, targetRange, {
        copyContent: options.target !== 'style',
        copyStyles: options.target === 'all' || options.target === 'style',
        contentType: options.type,
        adjustFormulas: true,
      });
    }

    // Clear seed range if cut operation
    if (options.cut) {
      const seedCells = this.expandRangeToCells(seedRange);
      this.clearSourceCells(seedCells);
    }
  }

  /**
   * Fill a target range with a seed range using column-first strategy
   * Step 0: Clear existing cell styles in target (Excel behavior)
   * Step 1: Fill down - extend seed pattern vertically to match target height
   * Step 2: Replicate right - copy filled columns horizontally
   */
  private fillRangeWithSeed(
    seedRange: RangeAddress,
    targetRange: RangeAddress,
    options: {
      copyContent: boolean;
      copyStyles: boolean;
      contentType: "value" | "formula";
      adjustFormulas: boolean;
    }
  ): void {
    // Step 0: Clear existing cell styles in target range (Excel-like replacement)
    if (options.copyStyles) {
      this.styleManager.clearCellStylesInRange(targetRange);
    }

    const seedCells = this.expandRangeToCells(seedRange);
    const seedWidth = this.getRangeWidth(seedRange);
    const seedHeight = this.getRangeHeight(seedRange);
    const targetWidth = this.getRangeWidth(targetRange);
    const targetHeight = this.getRangeHeight(targetRange);

    // Step 1: Fill down - replicate seed pattern vertically
    const filledColumns: Map<string, { cell: CellAddress; content: SerializedCellValue }> = new Map();
    
    for (let col = 0; col < seedWidth; col++) {
      for (let row = 0; row < targetHeight; row++) {
        const seedRow = row % seedHeight;
        const seedCell = seedCells.find(
          (c) => 
            c.colIndex === seedRange.range.start.col + col &&
            c.rowIndex === seedRange.range.start.row + seedRow
        );

        if (seedCell) {
          const targetCell: CellAddress = {
            workbookName: targetRange.workbookName,
            sheetName: targetRange.sheetName,
            colIndex: targetRange.range.start.col + col,
            rowIndex: targetRange.range.start.row + row,
          };

          const rowDelta = targetCell.rowIndex - seedCell.rowIndex;
          const colDelta = targetCell.colIndex - seedCell.colIndex;

          if (options.copyContent) {
            this.copyCellContentWithOffset(seedCell, targetCell, rowDelta, colDelta, {
              type: options.contentType,
              cut: false,
              target: 'content',
            });
          }

          if (options.copyStyles) {
            this.copyCellFormatting(seedCell, targetCell);
          }

          // Store filled column for horizontal replication
          const key = `${targetCell.colIndex}-${targetCell.rowIndex}`;
          const sheet = this.workbookManager.getSheet({
            workbookName: targetCell.workbookName,
            sheetName: targetCell.sheetName,
          });
          const cellKey = `${String.fromCharCode(65 + targetCell.colIndex)}${targetCell.rowIndex + 1}`;
          const content = sheet?.content.get(cellKey) || "";
          filledColumns.set(key, { cell: targetCell, content });
        }
      }
    }

    // Step 2: Replicate right - copy filled columns horizontally
    if (targetWidth > seedWidth) {
      for (let col = seedWidth; col < targetWidth; col++) {
        const sourceCol = col % seedWidth;
        
        for (let row = 0; row < targetHeight; row++) {
          const sourceCell: CellAddress = {
            workbookName: targetRange.workbookName,
            sheetName: targetRange.sheetName,
            colIndex: targetRange.range.start.col + sourceCol,
            rowIndex: targetRange.range.start.row + row,
          };

          const targetCell: CellAddress = {
            workbookName: targetRange.workbookName,
            sheetName: targetRange.sheetName,
            colIndex: targetRange.range.start.col + col,
            rowIndex: targetRange.range.start.row + row,
          };

          const colDelta = targetCell.colIndex - sourceCell.colIndex;

          if (options.copyContent) {
            this.copyCellContentWithOffset(sourceCell, targetCell, 0, colDelta, {
              type: options.contentType,
              cut: false,
              target: 'content',
            });
          }

          if (options.copyStyles) {
            this.copyCellFormatting(sourceCell, targetCell);
          }
        }
      }
    }
  }

  /**
   * Get the width of a range (number of columns)
   */
  private getRangeWidth(range: RangeAddress): number {
    if (range.range.end.col.type === "infinity") {
      return 100; // Default for infinite
    }
    return range.range.end.col.value - range.range.start.col + 1;
  }

  /**
   * Get the height of a range (number of rows)
   */
  private getRangeHeight(range: RangeAddress): number {
    if (range.range.end.row.type === "infinity") {
      return 100; // Default for infinite
    }
    return range.range.end.row.value - range.range.start.row + 1;
  }

  /**
   * Expand a RangeAddress into an array of CellAddress
   * Handles finite ranges, row-bounded, and column-bounded ranges
   */
  private expandRangeToCells(rangeAddress: RangeAddress): CellAddress[] {
    const { workbookName, sheetName, range } = rangeAddress;
    const cells: CellAddress[] = [];

    const startCol = range.start.col;
    const startRow = range.start.row;

    // Determine end column
    let endCol: number;
    if (range.end.col.type === "infinity") {
      // Limit infinite column ranges to a reasonable size (e.g., 100 columns)
      endCol = startCol + 99;
    } else {
      endCol = range.end.col.value;
    }

    // Determine end row
    let endRow: number;
    if (range.end.row.type === "infinity") {
      // Limit infinite row ranges to a reasonable size (e.g., 100 rows)
      endRow = startRow + 99;
    } else {
      endRow = range.end.row.value;
    }

    // Generate all cells in the range
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        cells.push({
          workbookName,
          sheetName,
          colIndex: col,
          rowIndex: row,
        });
      }
    }

    return cells;
  }

  /**
   * Copy cell content with explicit row/column offset for fill operations
   */
  private copyCellContentWithOffset(
    sourceCell: CellAddress,
    targetCell: CellAddress,
    rowDelta: number,
    colDelta: number,
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
      // Source cell is empty - clear target
      const targetSheet = this.workbookManager.getSheet({
        workbookName: targetCell.workbookName,
        sheetName: targetCell.sheetName,
      });
      if (targetSheet) {
        const targetKey = `${String.fromCharCode(65 + targetCell.colIndex)}${
          targetCell.rowIndex + 1
        }`;
        targetSheet.content.delete(targetKey);
      }
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
      // Copy formula with offset adjustment
      if (typeof cellContent === "string" && cellContent.startsWith("=")) {
        // Adjust formula references based on offset
        targetContent = this.adjustFormulaWithOffset(
          cellContent,
          rowDelta,
          colDelta
        );
      } else {
        // Not a formula, copy as-is
        targetContent = cellContent;
      }
    }

    // Set target cell content
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
   * Adjust formula references by a specific row/column offset
   */
  private adjustFormulaWithOffset(
    formula: string,
    rowDelta: number,
    colDelta: number
  ): string {
    try {
      const ast = parseFormula(formula.slice(1)); // Remove the "=" sign

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
      console.warn("Failed to adjust formula with offset:", error);
      return formula;
    }
  }

  /**
   * Copy formatting from one cell to another
   * Clears existing cell styles at target (Excel behavior) before copying new ones
   */
  private copyCellFormatting(
    sourceCell: CellAddress,
    targetCell: CellAddress
  ): void {
    // STEP 1: Clear existing cell styles at target cell (Excel-like replacement)
    const targetCellRange: RangeAddress = {
      workbookName: targetCell.workbookName,
      sheetName: targetCell.sheetName,
      range: {
        start: { col: targetCell.colIndex, row: targetCell.rowIndex },
        end: {
          col: { type: "number", value: targetCell.colIndex },
          row: { type: "number", value: targetCell.rowIndex },
        },
      },
    };
    
    this.styleManager.clearCellStylesInRange(targetCellRange);

    // Get all styles that intersect with the source cell
    const allConditionalStyles = this.styleManager.getAllConditionalStyles();
    const allCellStyles = this.styleManager.getAllCellStyles();

    const sourceCellRange: SpreadsheetRange = {
      start: { col: sourceCell.colIndex, row: sourceCell.rowIndex },
      end: {
        col: { type: "number", value: sourceCell.colIndex },
        row: { type: "number", value: sourceCell.rowIndex },
      },
    };

    // STEP 2: Copy conditional styles that apply to source cell
    for (const style of allConditionalStyles) {
      if (
        style.area.workbookName === sourceCell.workbookName &&
        style.area.sheetName === sourceCell.sheetName
      ) {
        const intersection = intersectRanges(style.area.range, sourceCellRange);
        if (intersection) {
          // Apply style to target cell
          const newStyle: ConditionalStyle = {
            area: {
              workbookName: targetCell.workbookName,
              sheetName: targetCell.sheetName,
              range: {
                start: { col: targetCell.colIndex, row: targetCell.rowIndex },
                end: {
                  col: { type: "number", value: targetCell.colIndex },
                  row: { type: "number", value: targetCell.rowIndex },
                },
              },
            },
            condition: style.condition,
          };
          this.styleManager.addConditionalStyle(newStyle);
        }
      }
    }

    // STEP 3: Copy cell styles that apply to source cell
    for (const style of allCellStyles) {
      if (
        style.area.workbookName === sourceCell.workbookName &&
        style.area.sheetName === sourceCell.sheetName
      ) {
        const intersection = intersectRanges(style.area.range, sourceCellRange);
        if (intersection) {
          // Apply style to target cell
          const newStyle: DirectCellStyle = {
            area: {
              workbookName: targetCell.workbookName,
              sheetName: targetCell.sheetName,
              range: {
                start: { col: targetCell.colIndex, row: targetCell.rowIndex },
                end: {
                  col: { type: "number", value: targetCell.colIndex },
                  row: { type: "number", value: targetCell.rowIndex },
                },
              },
            },
            style: style.style,
          };
          this.styleManager.addCellStyle(newStyle);
        }
      }
    }
  }
}

