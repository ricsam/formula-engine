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
} from "../../core/types";
import type { EvaluationContext } from "../../evaluator/evaluation-context";

/**
 * TEMPLATE function implementation
 */
export const TEMPLATE: FunctionDefinition = {
  name: "TEMPLATE",
  evaluate: function (node, context): FunctionEvaluationResult {
    throw new Error("Function not implemented");
  },
};
