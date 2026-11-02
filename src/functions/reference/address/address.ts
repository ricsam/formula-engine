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
} from "../../../core/types";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import { getSuppliedArgument } from "../../../parser/ast";

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
 * - ADDRESS(2, 3, 2, FALSE) returns "R2C[3]"
 *
 * Omitted arguments (a comma with nothing) take their defaults, as in Excel:
 * `ADDRESS(1, 1,, FALSE)` is "R1C1" and `ADDRESS(1, 1,,, "Sheet2")` is
 * "Sheet2!$A$1".
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

    // Evaluate row_num and column_num; an omitted one is 0, which is invalid
    const zero: FunctionEvaluationResult = {
      type: "value",
      result: { type: "number", value: 0 },
    };
    const rowNode = getSuppliedArgument(node, 0);
    const rowResult = rowNode ? this.evaluateNode(rowNode, context) : zero;
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
    const colNode = getSuppliedArgument(node, 1);
    const colResult = colNode ? this.evaluateNode(colNode, context) : zero;
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
    const absNode = getSuppliedArgument(node, 2);
    if (absNode) {
      const absResult = this.evaluateNode(absNode, context);
      if (
        absResult.type === "error" ||
        absResult.type === "awaiting-evaluation"
      ) {
        return absResult;
      }

      if (absResult.type !== "value" || absResult.result.type !== "number") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "ADDRESS function abs_num must be a number",
          errAddress: context.dependencyNode,
        };
      }
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

    // Optional a1 parameter: TRUE for A1 style, FALSE for R1C1 style
    const a1Node = getSuppliedArgument(node, 3);
    if (a1Node) {
      const a1Result = this.evaluateNode(a1Node, context);
      if (
        a1Result.type === "error" ||
        a1Result.type === "awaiting-evaluation"
      ) {
        return a1Result;
      }

      if (a1Result.type === "value" && a1Result.result.type === "boolean") {
        a1Style = a1Result.result.value;
      } else if (
        a1Result.type === "value" &&
        a1Result.result.type === "number"
      ) {
        a1Style = a1Result.result.value !== 0;
      } else {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "ADDRESS function a1 must be a logical value",
          errAddress: context.dependencyNode,
        };
      }
    }

    // Optional sheet_text parameter
    const sheetNode = getSuppliedArgument(node, 4);
    if (sheetNode) {
      const sheetResult = this.evaluateNode(sheetNode, context);
      if (
        sheetResult.type === "error" ||
        sheetResult.type === "awaiting-evaluation"
      ) {
        return sheetResult;
      }

      if (sheetResult.type !== "value") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "ADDRESS function sheet_text must be a single value",
          errAddress: context.dependencyNode,
        };
      }
      const sheetValue = sheetResult.result;
      sheetText =
        sheetValue.type === "string"
          ? sheetValue.value
          : sheetValue.type === "number"
          ? String(sheetValue.value)
          : sheetValue.type === "boolean"
          ? sheetValue.value
            ? "TRUE"
            : "FALSE"
          : "";
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
      // R1C1 style reference: a relative part is written in brackets
      const rowPart = absNum === 1 || absNum === 2 ? `R${rowNum}` : `R[${rowNum}]`;
      const colPart = absNum === 1 || absNum === 3 ? `C${colNum}` : `C[${colNum}]`;
      address = `${rowPart}${colPart}`;
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
