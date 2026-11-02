import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import {
  getResultShape,
  type ArrayShape,
} from "../../../evaluator/array-shape";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import type { FunctionNode } from "../../../parser/ast";

/**
 * Shared implementation of ROWS and COLUMNS: evaluates the single argument and
 * returns the size of one of its dimensions.
 *
 * The size comes from the argument's spill area, so a range is measured without
 * reading or depending on its cells. An unbounded dimension (the rows of `A:A`)
 * has no finite size because this engine has no sheet limit, so it counts as
 * positive infinity, the value of the `INFINITY` literal.
 */
export function countArgumentDimension(
  this: FormulaEvaluator,
  node: FunctionNode,
  context: EvaluationContext,
  options: { functionName: string; dimension: keyof ArrayShape }
): FunctionEvaluationResult {
  if (node.args.length !== 1) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `${options.functionName} function takes exactly 1 argument`,
      errAddress: context.dependencyNode,
    };
  }

  const arrayResult = this.evaluateNode(node.args[0]!, context);
  if (
    arrayResult.type === "error" ||
    arrayResult.type === "awaiting-evaluation"
  ) {
    return arrayResult;
  }

  // A shape does not depend on where the array spills, only on its size, so
  // measuring it from the current cell adds no context dependency.
  const count = getResultShape(arrayResult, context.cellAddress)[
    options.dimension
  ];

  return {
    type: "value",
    result: Number.isFinite(count)
      ? { type: "number", value: count }
      : { type: "infinity", sign: "positive" },
  };
}

/**
 * ROWS(array)
 *
 * Returns the number of rows in a reference or array.
 *
 * - A range counts its rows without reading or depending on its cells:
 *   `ROWS(A1:B10)` is 10 whatever the cells hold, even inside the range itself.
 * - A spilled result counts the rows of the area it spills onto:
 *   `ROWS(SEQUENCE(5, 2))` is 5, `ROWS({1,2;3,4;5,6})` is 3, and
 *   `ROWS(UNIQUE(A2:A100))` is the number of distinct values.
 * - A single value is a 1x1 array: `ROWS(5)`, `ROWS("text")` and `ROWS(A1)` are 1.
 *
 * Deviations from Excel, which follow from this engine's model:
 * - A reference to the cell a formula spills from evaluates to the whole spill
 *   (Excel needs the spill operator `C1#` for that). With C1 holding
 *   `=SEQUENCE(5)`, `ROWS(C1)` is 5 where Excel returns 1. The same applies to
 *   one-cell ranges such as `ROWS($C$1:C1)`.
 * - Unbounded rows (`A:A`, `A2:A`, `SEQUENCE(INFINITY)`) count as INFINITY,
 *   because the engine has no sheet limit. Excel returns 1048576 for `ROWS(A:A)`.
 * - An error argument propagates, as in Excel: `ROWS(UNIQUE(...))` returns
 *   UNIQUE's error. Because a single-cell reference is evaluated to find out
 *   whether it spills, a reference to a cell holding an error also returns that
 *   error, where Excel returns 1.
 */
export const ROWS: FunctionDefinition = {
  name: "ROWS",
  evaluate: function (node, context): FunctionEvaluationResult {
    return countArgumentDimension.call(this, node, context, {
      functionName: "ROWS",
      dimension: "rows",
    });
  },
};
