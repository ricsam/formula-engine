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
 * TEMPLATE function implementation
 */
export const TEMPLATE: FunctionDefinition = {
  name: "TEMPLATE",
  evaluate: function (node, context): FunctionEvaluationResult {
    throw new Error("Function not implemented");
  },
};
