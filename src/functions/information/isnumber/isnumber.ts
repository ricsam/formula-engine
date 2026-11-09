import { createIsFunction } from "../is-function";

/**
 * ISNUMBER(value)
 *
 * Returns TRUE when the value is a number, FALSE otherwise.
 *
 * Examples:
 *   ISNUMBER(42) - TRUE
 *   ISNUMBER("42") - FALSE: text that looks like a number is still text
 *   ISNUMBER(A1) - FALSE when A1 is empty, holds text, a logical value or an error
 *   ISNUMBER(MATCH("x", A1:A9, 0)) - whether "x" occurs in A1:A9
 *   ISNUMBER(A1:A3) - spills TRUE/FALSE for each cell
 *
 * - INFINITY and -INFINITY, this engine's results of a division by zero, are
 *   numbers.
 * - Errors give FALSE instead of propagating, as in Excel.
 * - An array argument is tested element by element and spills with its shape.
 *   Aggregates read every cell of a bounded result, empty cells included (see
 *   `is-function.ts`).
 */
export const ISNUMBER = createIsFunction(
  "ISNUMBER",
  (cell) =>
    cell.type === "value" &&
    (cell.result.type === "number" || cell.result.type === "infinity")
);
