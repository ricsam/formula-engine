import type { LookupOrder } from "../core/managers/range-eval-order-builder";
import {
  type ErrorEvaluationResult,
  type SingleEvaluationResult,
} from "../core/types";
import { getCellReference } from "../core/utils";
import type { EvaluationContext } from "../evaluator/evaluation-context";
import { AwaitingEvaluationError } from "../evaluator/evaluation-error";
import type { FormulaEvaluator } from "../evaluator/formula-evaluator";
import type { FunctionNode } from "../parser/ast";

/**
 * Creates an array of all values from arguments of a function node
 * Returns SingleEvaluationResult for each value found in the arguments
 *
 * This is a shared utility used by all basic aggregation functions (SUM, MIN, MAX, AVERAGE)
 * to iterate over their arguments in a consistent way.
 *
 * @param evaluator - The formula evaluator instance
 * @param node - The function node containing arguments
 * @param context - The evaluation context
 * @returns Array of SingleEvaluationResult for each value in the arguments
 */
export function createArgumentIterator(
  evaluator: FormulaEvaluator,
  node: FunctionNode,
  context: EvaluationContext,
  lookupOrder: LookupOrder
): { type: "values"; values: SingleEvaluationResult[] } | ErrorEvaluationResult {
  const results: SingleEvaluationResult[] = [];

  for (const arg of node.args) {
    const result = evaluator.evaluateNode(arg, context);

    if (result.type === "spilled-values") {
      // Get all cells in the spilled range
      const cellValues = result.evaluateAllCells.call(evaluator, {
        context,
        evaluate: result.evaluate,
        origin: context.cellAddress,
        lookupOrder,
      });

      if (cellValues.type === "awaiting-evaluation") {
        return cellValues;
      }

      if (cellValues.type === "error") {
        results.push(cellValues);
      } else {
        for (const cellValue of cellValues.values) {
          results.push(cellValue.result);
        }
      }
    } else {
      results.push(result);
    }
  }

  return { type: "values", values: results };
}
