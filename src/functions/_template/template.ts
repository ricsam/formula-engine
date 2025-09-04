import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type EvaluationContext,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
} from "src/core/types";

/**
 * TEMPLATE function implementation
 */
export const TEMPLATE: FunctionDefinition = {
  name: "TEMPLATE",
  evaluate: function (node, context): FunctionEvaluationResult {
    throw new Error("Function not implemented");
  },
};
