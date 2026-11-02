import {
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import { evaluateArrayToVector } from "../tocol/tocol";

/**
 * TOROW(array, [ignore], [scan_by_column])
 * Returns the cells of an array as a single row.
 *
 * @param array - The array or range to flatten. A single value, or an error,
 *   is a 1x1 array.
 * @param ignore - [Optional] 0 keeps every cell (default), 1 skips blanks,
 *   2 skips errors, 3 skips blanks and errors. Numbers are truncated, logical
 *   values count as 1/0 and a blank counts as 0; any other value gives #VALUE!.
 * @param scan_by_column - [Optional] FALSE (default) reads the array row by
 *   row, TRUE column by column. Numbers count as TRUE unless 0.
 *
 * TOROW behaves exactly like TOCOL (see its documentation for blanks, errors,
 * #CALC! and unbounded arrays), except that the cells are laid out in one row.
 */
export const TOROW: FunctionDefinition = {
  name: "TOROW",
  evaluate: function (node, context): FunctionEvaluationResult {
    return evaluateArrayToVector.call(this, node, context, "row");
  },
};
