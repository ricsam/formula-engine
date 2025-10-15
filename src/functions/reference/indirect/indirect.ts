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
import { parseFormula } from "src/parser/parser";

/**
 * INDIRECT function - Returns the reference specified by a text string
 * 
 * Usage: INDIRECT(ref_text, [a1])
 * 
 * ref_text: A reference to a cell that contains an A1-style reference, an R1C1-style reference,
 *           a name defined as a reference, or a reference to a cell as a text string
 * a1: (optional) A logical value that specifies what type of reference is in ref_text
 *     TRUE or omitted = A1-style reference
 *     FALSE = R1C1-style reference
 * 
 * Examples:
 * - INDIRECT("A1") returns the value in cell A1
 * - INDIRECT("B" & ROW()) returns the value in column B of the current row
 * - INDIRECT("Sheet2!A1") returns the value in cell A1 of Sheet2
 * 
 * Note: This function is volatile - it recalculates whenever the worksheet recalculates
 */
export const INDIRECT: FunctionDefinition = {
  name: "INDIRECT",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length === 0 || node.args.length > 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INDIRECT function requires 1 or 2 arguments",
        errAddress: context.originCell.cellAddress,
      };
    }

    // Evaluate ref_text
    const refTextResult = this.evaluateNode(node.args[0]!, context);
    if (
      refTextResult.type === "error" ||
      refTextResult.type === "awaiting-evaluation"
    ) {
      return refTextResult;
    }

    if (refTextResult.type !== "value" || refTextResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INDIRECT function ref_text must be a string",
        errAddress: context.originCell.cellAddress,
      };
    }

    const refText = refTextResult.result.value.trim();

    // Optional a1 parameter
    let a1Style = true;
    if (node.args.length === 2) {
      const a1Result = this.evaluateNode(node.args[1]!, context);
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

    // R1C1 style not yet supported
    if (!a1Style) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INDIRECT function does not yet support R1C1 style references",
        errAddress: context.originCell.cellAddress,
      };
    }

    // Parse the reference string as a formula (without the = sign)
    try {
      const ast = parseFormula(refText);
      
      // The parsed result should be a reference, range, or named expression
      if (ast.type !== "reference" && ast.type !== "range" && ast.type !== "named-expression") {
        return {
          type: "error",
          err: FormulaError.REF,
          message: `INDIRECT requires a valid cell or range reference, got: ${refText}`,
          errAddress: context.originCell.cellAddress,
        };
      }
      
      // Evaluate the parsed reference
      const result = this.evaluateNode(ast, context);
      return result;
      
    } catch (error) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: `INDIRECT could not parse reference: ${refText}`,
        errAddress: context.originCell.cellAddress,
      };
    }
  },
};
