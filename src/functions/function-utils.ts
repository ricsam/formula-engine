import { type SingleEvaluationResult } from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { FunctionNode } from "src/parser/ast";

/**
 * Creates an iterator over all arguments of a function node
 * Yields SingleEvaluationResult for each value found in the arguments
 *
 * This is a shared utility used by all basic aggregation functions (SUM, MIN, MAX, AVERAGE)
 * to iterate over their arguments in a consistent way.
 *
 * @param evaluator - The formula evaluator instance
 * @param node - The function node containing arguments
 * @param context - The evaluation context
 * @yields SingleEvaluationResult for each value in the arguments
 */
export function* createArgumentIterator(
  evaluator: FormulaEvaluator,
  node: FunctionNode,
  context: EvaluationContext
): Generator<SingleEvaluationResult, void, unknown> {
  for (const arg of node.args) {
    const result = evaluator.evaluateNode(arg, context);

    if (result.type === "error") {
      yield result;
    } else if (result.type === "value") {
      yield result;
    } else if (result.type === "spilled-values") {
      // Iterate over all cells in the spilled range
      const cellValues = result.evaluateAllCells.call(evaluator, {
        context,
        evaluate: result.evaluate,
        origin: context.currentCell,
      });

      for (const cellValue of cellValues) {
        yield cellValue.result;
      }
    }
  }
}
