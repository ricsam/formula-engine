import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import { countNonEmptyValues } from "./count-utils";

/**
 * COUNTA(value1, [value2], ...)
 *
 * Counts the values that are not empty: numbers, text, logical values and
 * errors.
 *
 * Examples:
 *   COUNTA(A1:A10) - counts the populated cells of A1:A10
 *   COUNTA(A:A) - counts the populated cells of column A
 *   COUNTA(1, "", TRUE) - returns 3
 *   COUNTA(UNIQUE(A2:A100)) - counts the distinct values
 *
 * - In a range, unpopulated cells are not counted, but a cell whose formula
 *   returns "" is, as in Excel. Cells that other formulas spill onto count like
 *   populated cells. Errors inside a range are counted, not propagated.
 * - A spilled argument (SEQUENCE, UNIQUE, TOCOL, an array literal, ...) counts
 *   its cells.
 * - A directly supplied argument counts as one value, including the literal ""
 *   and an error value. A reference to an empty cell counts 0.
 *
 * Deviations from Excel, which follow from this engine's model:
 * - A reference to the cell a formula spills from evaluates to the whole spill,
 *   so with C1 holding `=SEQUENCE(5)`, `COUNTA(C1)` counts all 5 values (Excel
 *   needs `C1#` for that and returns 1 for `COUNTA(C1)`).
 * - The engine reads an empty cell as "", so a direct single-cell reference to a
 *   formula returning "" counts 0 (Excel counts 1). Inside a range such a cell
 *   is still counted.
 * - Calling COUNTA without arguments returns #VALUE!; Excel rejects the formula.
 */
export const COUNTA: FunctionDefinition = {
  name: "COUNTA",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length === 0) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "COUNTA function requires at least 1 argument",
        errAddress: context.dependencyNode,
      };
    }

    return countNonEmptyValues(this, node.args, context);
  },
};
