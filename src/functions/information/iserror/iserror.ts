import { createIsFunction } from "../is-function";

/**
 * ISERROR(value)
 *
 * Returns TRUE when the value is an error, FALSE otherwise.
 *
 * Examples:
 *   ISERROR(NOSUCHFN()) - TRUE: #NAME?
 *   ISERROR(MATCH("x", A1:A9, 0)) - TRUE when "x" is not in A1:A9 (#N/A)
 *   IF(ISERROR(A1), "", A1 * 2) - keeps an error in A1 out of the result
 *   ISERROR(A1:A3) - spills TRUE/FALSE for each cell
 *   SUM(IF(ISERROR(A1:A9), 1, 0)) - counts the errors in A1:A9
 *
 * - Every error value counts: #N/A, #VALUE!, #REF!, #NAME?, #NUM!, #DIV/0!,
 *   #CALC!, #SPILL!, #CYCLE! and #ERROR!. ISERROR is TRUE exactly where IFERROR
 *   would return its fallback.
 * - ISERROR never returns an error itself: the error it tests is its answer.
 * - A division by zero is INFINITY in this engine, a value rather than an
 *   error, so `ISERROR(1/0)` is FALSE. Excel returns #DIV/0! there, which
 *   ISERROR reports as TRUE.
 * - An array argument is tested element by element and spills with its shape.
 *   Aggregates read every cell of a bounded result, empty cells included (see
 *   `is-function.ts`).
 */
export const ISERROR = createIsFunction(
  "ISERROR",
  (cell) => cell.type === "error"
);
