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
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { ReferenceNode } from "src/parser/ast";
import { isCellAddress, isRangeAddress } from "src/core/utils";

/**
 * CELL function - Returns information about the formatting, location, or contents of a cell
 * 
 * Usage: CELL(info_type, [reference])
 * 
 * info_type: Text value that specifies what type of cell information you want
 * reference: The cell or range to get information about (defaults to current cell if omitted)
 * 
 * Supported info_types:
 * - "address" - Returns the reference of the first cell in reference, as text
 * - "col" - Returns the column number of the cell
 * - "row" - Returns the row number of the cell
 * - "contents" - Returns the value of the upper-left cell
 * - "type" - Returns "b" for blank, "l" for label (text), "v" for value (number)
 * - "width" - Returns the column width
 * - "filename" - Returns the filename (including full path) of the file
 * 
 * Examples:
 * - CELL("address", A1) returns "$A$1"
 * - CELL("row", B5) returns 5
 * - CELL("col", C3) returns 3
 */
export const CELL: FunctionDefinition = {
  name: "CELL",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length === 0 || node.args.length > 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "CELL function requires 1 or 2 arguments",
        errAddress: context.originCell.cellAddress,
      };
    }

    // Evaluate the info_type argument
    const infoTypeResult = this.evaluateNode(node.args[0]!, context);
    if (
      infoTypeResult.type === "error" ||
      infoTypeResult.type === "awaiting-evaluation"
    ) {
      return infoTypeResult;
    }

    if (infoTypeResult.type !== "value" || infoTypeResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "CELL function info_type must be a string",
        errAddress: context.originCell.cellAddress,
      };
    }

    const infoType = infoTypeResult.result.value.toLowerCase();

    // Determine the reference cell
    let referenceCell: CellAddress = context.originCell.cellAddress;
    let referenceValue: CellValue | null = null;
    
    if (node.args.length === 2) {
      const refArg = node.args[1]!;
      
      // Evaluate the reference argument
      const refResult = this.evaluateNode(refArg, context);
      if (
        refResult.type === "error" ||
        refResult.type === "awaiting-evaluation"
      ) {
        return refResult;
      }
      
      // Check if we have a cell or range reference
      const cellOrRange =
        refResult.type === "spilled-values"
          ? (refResult.sourceCell ?? refResult.sourceRange)
          : refResult.sourceCell;
      
      if (!cellOrRange) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "CELL function requires a cell or range reference",
          errAddress: context.originCell.cellAddress,
        };
      }
      
      // Extract the cell address
      if (isCellAddress(cellOrRange)) {
        // It's a CellAddress
        referenceCell = cellOrRange;
      } else if (isRangeAddress(cellOrRange)) {
        // It's a RangeAddress - use the upper-left cell
        referenceCell = {
          sheetName: cellOrRange.sheetName,
          workbookName: cellOrRange.workbookName,
          colIndex: cellOrRange.range.start.col,
          rowIndex: cellOrRange.range.start.row,
        };
      } else {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "CELL function requires a cell or range reference",
          errAddress: context.originCell.cellAddress,
        };
      }
      
      // For "contents" and "type", we need to get the cell value
      if (infoType === "contents" || infoType === "type") {
        if (refResult.type === "value") {
          referenceValue = refResult.result;
        } else if (refResult.type === "spilled-values") {
          const cellResult = refResult.evaluate({ x: 0, y: 0 }, context);
          if (cellResult && cellResult.type === "value") {
            referenceValue = cellResult.result;
          }
        }
      }
    }

    // Handle different info types (return 1-based row/column numbers)
    switch (infoType) {
      case "row":
        return {
          type: "value",
          result: { type: "number", value: referenceCell.rowIndex + 1 },
        };
      
      case "col":
        return {
          type: "value",
          result: { type: "number", value: referenceCell.colIndex + 1 },
        };
      
      case "address": {
        // Return the cell address as text in A1 format
        const col = columnNumberToLetter(referenceCell.colIndex + 1);
        const row = referenceCell.rowIndex + 1;
        let address = `$${col}$${row}`;
        
        // Include sheet name if it's different from current sheet
        if (referenceCell.sheetName !== context.originCell.cellAddress.sheetName) {
          address = `${referenceCell.sheetName}!${address}`;
        }
        
        return {
          type: "value",
          result: { type: "string", value: address },
        };
      }
      
      case "contents":
        // Return the value of the cell
        if (referenceValue !== null) {
          // Empty string counts as empty cell in Excel
          if (referenceValue.type === "string" && referenceValue.value === "") {
            return {
              type: "value",
              result: { type: "number", value: 0 },
            };
          }
          return {
            type: "value",
            result: referenceValue,
          };
        }
        // If we couldn't get the value, return 0 (empty cell)
        return {
          type: "value",
          result: { type: "number", value: 0 },
        };
      
      case "type": {
        // Return "b" for blank, "l" for label (text), "v" for value (number)
        if (referenceValue === null) {
          return {
            type: "value",
            result: { type: "string", value: "b" },
          };
        }
        // Empty string counts as blank cell in Excel
        if (referenceValue.type === "string" && referenceValue.value === "") {
          return {
            type: "value",
            result: { type: "string", value: "b" },
          };
        }
        if (referenceValue.type === "string") {
          return {
            type: "value",
            result: { type: "string", value: "l" },
          };
        }
        if (referenceValue.type === "number") {
          return {
            type: "value",
            result: { type: "string", value: "v" },
          };
        }
        // Boolean or other types default to "v"
        return {
          type: "value",
          result: { type: "string", value: "v" },
        };
      }
      
      default:
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: `CELL function unknown info_type: "${infoType}"`,
          errAddress: context.originCell.cellAddress,
        };
    }
  },
};

/**
 * Helper function to convert column number (1-based) to letter(s)
 */
function columnNumberToLetter(num: number): string {
  let letter = "";
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}
