import { createIsFunction } from "../is-function";

/**
 * ISBLANK(value)
 *
 * Returns TRUE when the value is blank, FALSE otherwise.
 *
 * Examples:
 *   ISBLANK(A1) - TRUE when A1 is empty
 *   IF(ISBLANK(A1), "", A1 * 2) - guards arithmetic against an empty cell
 *   ISBLANK(A1:A3) - spills TRUE/FALSE for each cell
 *   SUM(IF(ISBLANK(A1:A9), 1, 0)) - counts the empty cells of A1:A9
 *
 * - The engine reads an empty cell as the empty string, so every value that
 *   reads as `""` is blank: an empty cell, a cell whose formula returns `""`,
 *   and `""` itself. Excel returns TRUE only for an empty cell, so
 *   `ISBLANK("")` and ISBLANK of a cell holding `=""` are FALSE there.
 * - `ISBLANK(0)`, `ISBLANK(" ")`, `ISBLANK(FALSE)` and errors give FALSE.
 * - A reference to a cell that a spill covers is blank when the spilled value
 *   is `""`, as for a blank cell spilled by `=A1:A3`; Excel shows 0 there.
 * - An array argument is tested element by element and spills with its shape.
 *   Aggregates read every cell of a bounded result, empty cells of a range
 *   included (see `is-function.ts`).
 */
export const ISBLANK = createIsFunction(
  "ISBLANK",
  (cell) =>
    cell.type === "value" &&
    cell.result.type === "string" &&
    cell.result.value === ""
);
