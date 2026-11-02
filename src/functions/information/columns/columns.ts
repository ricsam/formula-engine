import {
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import { countArgumentDimension } from "../rows/rows";

/**
 * COLUMNS(array), also callable as COLS(array)
 *
 * Returns the number of columns in a reference or array.
 *
 * - A range counts its columns without reading or depending on its cells:
 *   `COLUMNS(A1:C10)` is 3 whatever the cells hold, even inside the range itself.
 * - A spilled result counts the columns of the area it spills onto:
 *   `COLUMNS(SEQUENCE(2, 4))` is 4, `COLUMNS({1,2,3})` is 3, and
 *   `COLUMNS(UNIQUE(A1:Z1, TRUE))` is the number of distinct columns.
 * - A single value is a 1x1 array: `COLUMNS(5)` and `COLUMNS(A1)` are 1.
 *
 * Deviations from Excel, which follow from this engine's model:
 * - A reference to the cell a formula spills from evaluates to the whole spill
 *   (Excel needs the spill operator `C1#` for that). With C1 holding
 *   `=SEQUENCE(1, 5)`, `COLUMNS(C1)` is 5 where Excel returns 1.
 * - Unbounded columns (`1:1`, `B2:2`, `SEQUENCE(1, INFINITY)`) count as
 *   INFINITY, because the engine has no sheet limit. Excel returns 16384 for
 *   `COLUMNS(1:1)`.
 * - An error argument propagates, as in Excel. Because a single-cell reference is
 *   evaluated to find out whether it spills, a reference to a cell holding an
 *   error also returns that error, where Excel returns 1.
 */
export const COLUMNS: FunctionDefinition = {
  name: "COLUMNS",
  aliases: ["COLS"],
  evaluate: function (node, context): FunctionEvaluationResult {
    return countArgumentDimension.call(this, node, context, {
      functionName: "COLUMNS",
      dimension: "cols",
    });
  },
};
