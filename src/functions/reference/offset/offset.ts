import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
  type SpreadsheetRangeEnd,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { ReferenceNode, RangeNode } from "src/parser/ast";
import { isCellAddress, isRangeAddress } from "src/core/utils";

/**
 * OFFSET function - Returns a reference offset from a given reference
 * 
 * Usage: OFFSET(reference, rows, cols, [height], [width])
 * 
 * reference: The reference from which you want to base the offset
 * rows: The number of rows, up or down, that you want the upper-left cell to refer to
 * cols: The number of columns, left or right, that you want the upper-left cell to refer to
 * height: (optional) The height, in number of rows, that you want the returned reference to be
 * width: (optional) The width, in number of columns, that you want the returned reference to be
 * 
 * Examples:
 * - OFFSET(A1, 2, 3) returns the value in cell D3 (2 rows down, 3 columns right from A1)
 * - OFFSET(A1, 0, 0, 3, 1) returns the range A1:A3
 * - OFFSET(A1, 1, 1, 2, 2) returns the range B2:C3
 * 
 * Notes:
 * - If rows or cols offset reference from the edge of the worksheet, returns #REF! error
 * - Height and width must be positive
 * - If height or width is omitted, the same height or width as reference is used
 */
export const OFFSET: FunctionDefinition = {
  name: "OFFSET",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 3 || node.args.length > 5) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "OFFSET function requires 3 to 5 arguments",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate reference
    const refResult = this.evaluateNode(node.args[0]!, context);
    if (
      refResult.type === "error" ||
      refResult.type === "awaiting-evaluation"
    ) {
      return refResult;
    }

    // Evaluate rows offset
    const rowsResult = this.evaluateNode(node.args[1]!, context);
    if (
      rowsResult.type === "error" ||
      rowsResult.type === "awaiting-evaluation"
    ) {
      return rowsResult;
    }

    if (rowsResult.type !== "value" || rowsResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "OFFSET function rows must be a number",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate cols offset
    const colsResult = this.evaluateNode(node.args[2]!, context);
    if (
      colsResult.type === "error" ||
      colsResult.type === "awaiting-evaluation"
    ) {
      return colsResult;
    }

    if (colsResult.type !== "value" || colsResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "OFFSET function cols must be a number",
        errAddress: context.dependencyNode,
      };
    }

    const rowsOffset = Math.floor(rowsResult.result.value);
    const colsOffset = Math.floor(colsResult.result.value);

    // Optional height parameter
    let height: number | undefined;
    if (node.args.length >= 4) {
      const heightResult = this.evaluateNode(node.args[3]!, context);
      if (
        heightResult.type === "error" ||
        heightResult.type === "awaiting-evaluation"
      ) {
        return heightResult;
      }

      if (heightResult.type === "value" && heightResult.result.type === "number") {
        height = Math.floor(heightResult.result.value);
        if (height < 1) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "OFFSET function height must be positive",
            errAddress: context.dependencyNode,
          };
        }
      }
    }

    // Optional width parameter
    let width: number | undefined;
    if (node.args.length === 5) {
      const widthResult = this.evaluateNode(node.args[4]!, context);
      if (
        widthResult.type === "error" ||
        widthResult.type === "awaiting-evaluation"
      ) {
        return widthResult;
      }

      if (widthResult.type === "value" && widthResult.result.type === "number") {
        width = Math.floor(widthResult.result.value);
        if (width < 1) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "OFFSET function width must be positive",
            errAddress: context.dependencyNode,
          };
        }
      }
    }

    // Get the base cell/range from the evaluated reference
    let baseRange: SpreadsheetRange;
    let baseSheetName: string = context.cellAddress.sheetName;
    let baseWorkbookName: string = context.cellAddress.workbookName;

    // Check if we have a cell or range reference
    const cellOrRange =
      refResult.type === "spilled-values"
        ? (refResult.sourceCell ?? refResult.sourceRange)
        : refResult.sourceCell;

    if (!cellOrRange) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "OFFSET function requires a cell or range reference",
        errAddress: context.dependencyNode,
      };
    }

    // Extract the base range from sourceCell or sourceRange
    if (isCellAddress(cellOrRange)) {
      // It's a CellAddress
      baseSheetName = cellOrRange.sheetName;
      baseWorkbookName = cellOrRange.workbookName;
      baseRange = {
        start: {
          col: cellOrRange.colIndex,
          row: cellOrRange.rowIndex,
        },
        end: {
          col: { type: "number", value: cellOrRange.colIndex },
          row: { type: "number", value: cellOrRange.rowIndex },
        },
      };
    } else if (isRangeAddress(cellOrRange)) {
      // It's a RangeAddress
      baseSheetName = cellOrRange.sheetName;
      baseWorkbookName = cellOrRange.workbookName;
      baseRange = cellOrRange.range;
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "OFFSET function requires a cell or range reference",
        errAddress: context.dependencyNode,
      };
    }

    // Calculate the offset range
    const offsetStartCol = baseRange.start.col + colsOffset;
    const offsetStartRow = baseRange.start.row + rowsOffset;

    // Validate bounds (assuming reasonable spreadsheet size)
    if (offsetStartCol < 0 || offsetStartRow < 0) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: "OFFSET function resulted in invalid cell reference",
        errAddress: context.dependencyNode,
      };
    }

    // Calculate end position based on height/width or original range size
    let offsetEndCol: number;
    let offsetEndRow: number;

    if (width !== undefined) {
      offsetEndCol = offsetStartCol + width - 1;
    } else if (baseRange.end.col.type === "number") {
      offsetEndCol = offsetStartCol + (baseRange.end.col.value - baseRange.start.col);
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "OFFSET function requires finite reference when width is omitted",
        errAddress: context.dependencyNode,
      };
    }

    if (height !== undefined) {
      offsetEndRow = offsetStartRow + height - 1;
    } else if (baseRange.end.row.type === "number") {
      offsetEndRow = offsetStartRow + (baseRange.end.row.value - baseRange.start.row);
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "OFFSET function requires finite reference when height is omitted",
        errAddress: context.dependencyNode,
      };
    }

    // Validate the end bounds
    if (offsetEndCol < 0 || offsetEndRow < 0) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: "OFFSET function resulted in invalid cell reference",
        errAddress: context.dependencyNode,
      };
    }

    // Create the offset range
    const offsetRange: SpreadsheetRange = {
      start: {
        col: offsetStartCol,
        row: offsetStartRow,
      },
      end: {
        col: { type: "number", value: offsetEndCol },
        row: { type: "number", value: offsetEndRow },
      },
    };

    // If it's a single cell, return its value
    if (offsetStartCol === offsetEndCol && offsetStartRow === offsetEndRow) {
      return this.evaluateReference(
        {
          type: "reference",
          address: {
            colIndex: offsetStartCol,
            rowIndex: offsetStartRow,
          },
          sheetName: baseSheetName,
          workbookName: baseWorkbookName,
          isAbsolute: {
            col: true,
            row: true,
          },
        },
        context
      );
    }

    // Otherwise, return a range reference (spilled values)
    return this.evaluateRange(
      {
        type: "range",
        range: offsetRange,
        sheetName: baseSheetName,
        workbookName: baseWorkbookName,
        isAbsolute: {
          start: { col: true, row: true },
          end: { col: true, row: true },
        },
      },
      context
    );
  },
};
