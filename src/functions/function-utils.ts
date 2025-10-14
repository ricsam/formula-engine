import type { LookupOrder } from "src/core/managers";
import { type SingleEvaluationResult } from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import { AwaitingEvaluationError } from "src/evaluator/evaluation-error";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { FunctionNode } from "src/parser/ast";

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
): SingleEvaluationResult[] {
  const results: SingleEvaluationResult[] = [];

  for (const arg of node.args) {
    const result = evaluator.evaluateNode(arg, context);

    if (result.type === "awaiting-evaluation") {
      throw new AwaitingEvaluationError(
        context.originCell.cellAddress,
        result.waitingFor
      );
    }

    if (result.type === "error") {
      results.push(result);
    } else if (result.type === "value") {
      results.push(result);
    } else if (result.type === "spilled-values") {
      // Get all cells in the spilled range
      const cellValues = result.evaluateAllCells.call(evaluator, {
        context,
        evaluate: result.evaluate,
        origin: context.originCell.cellAddress,
        lookupOrder,
      });

      for (const cellValue of cellValues) {
        results.push(cellValue.result);
      }
    }
  }

  return results;
}
