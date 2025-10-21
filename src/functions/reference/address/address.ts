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

/**
 * ADDRESS function - Creates a cell reference as text given row and column numbers
 * 
 * Usage: ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])
 * 
 * row_num: The row number to use in the cell reference
 * column_num: The column number to use in the cell reference
 * abs_num: (optional) Type of reference to return (1=absolute, 2=absolute row/relative col, 3=relative row/absolute col, 4=relative)
 * a1: (optional) A logical value that specifies the A1 or R1C1 reference style (TRUE or omitted = A1 style)
 * sheet_text: (optional) The name of the worksheet to be used as the external reference
 * 
 * Examples:
 * - ADDRESS(2, 3) returns "$C$2"
 * - ADDRESS(2, 3, 4) returns "C2"
 * - ADDRESS(2, 3, 1, TRUE, "Sheet2") returns "Sheet2!$C$2"
 */
export const ADDRESS: FunctionDefinition = {
  name: "ADDRESS",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 2 || node.args.length > 5) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "ADDRESS function requires 2 to 5 arguments",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate row_num
    const rowResult = this.evaluateNode(node.args[0]!, context);
    if (
      rowResult.type === "error" ||
      rowResult.type === "awaiting-evaluation"
    ) {
      return rowResult;
    }

    if (rowResult.type !== "value" || rowResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "ADDRESS function row_num must be a number",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate column_num
    const colResult = this.evaluateNode(node.args[1]!, context);
    if (
      colResult.type === "error" ||
      colResult.type === "awaiting-evaluation"
    ) {
      return colResult;
    }

    if (colResult.type !== "value" || colResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "ADDRESS function column_num must be a number",
        errAddress: context.dependencyNode,
      };
    }

    const rowNum = Math.floor(rowResult.result.value);
    const colNum = Math.floor(colResult.result.value);

    if (rowNum < 1 || colNum < 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "ADDRESS function row and column numbers must be positive",
        errAddress: context.dependencyNode,
      };
    }

    // Default values
    let absNum = 1; // Absolute reference by default
    let a1Style = true;
    let sheetText = "";

    // Optional abs_num parameter
    if (node.args.length >= 3) {
      const absResult = this.evaluateNode(node.args[2]!, context);
      if (
        absResult.type === "error" ||
        absResult.type === "awaiting-evaluation"
      ) {
        return absResult;
      }

      if (absResult.type === "value" && absResult.result.type === "number") {
        absNum = Math.floor(absResult.result.value);
        if (absNum < 1 || absNum > 4) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "ADDRESS function abs_num must be between 1 and 4",
            errAddress: context.dependencyNode,
          };
        }
      }
    }

    // Optional a1 parameter
    if (node.args.length >= 4) {
      const a1Result = this.evaluateNode(node.args[3]!, context);
      if (
        a1Result.type === "error" ||
        a1Result.type === "awaiting-evaluation"
      ) {
        return a1Result;
      }

      if (a1Result.type === "value" && a1Result.result.type === "boolean") {
        a1Style = a1Result.result.value;
      }
    }

    // Optional sheet_text parameter
    if (node.args.length === 5) {
      const sheetResult = this.evaluateNode(node.args[4]!, context);
      if (
        sheetResult.type === "error" ||
        sheetResult.type === "awaiting-evaluation"
      ) {
        return sheetResult;
      }

      if (sheetResult.type === "value" && sheetResult.result.type === "string") {
        sheetText = sheetResult.result.value;
      }
    }

    // Build the address string
    let address = "";

    if (a1Style) {
      // A1 style reference
      const colLetter = columnNumberToLetter(colNum);
      
      switch (absNum) {
        case 1: // $A$1
          address = `$${colLetter}$${rowNum}`;
          break;
        case 2: // A$1
          address = `${colLetter}$${rowNum}`;
          break;
        case 3: // $A1
          address = `$${colLetter}${rowNum}`;
          break;
        case 4: // A1
          address = `${colLetter}${rowNum}`;
          break;
      }
    } else {
      // R1C1 style reference
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "ADDRESS function R1C1 style not yet implemented",
        errAddress: context.dependencyNode,
      };
    }

    // Add sheet prefix if provided
    if (sheetText) {
      address = `${sheetText}!${address}`;
    }

    return {
      type: "value",
      result: { type: "string", value: address },
    };
  },
};

/**
 * Helper function to convert column number to letter(s)
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
