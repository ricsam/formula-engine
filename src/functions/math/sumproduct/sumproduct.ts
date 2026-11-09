import {
  FormulaError,
  type AwaitingEvaluationResult,
  type ErrorEvaluationResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import { getSuppliedArgument, type ASTNode } from "../../../parser/ast";
import {
  toArrayArgument,
  type ArrayArgument,
  type ArrayCellResult,
} from "../../array/array-utils";

function describeShape(array: ArrayArgument): string {
  const size = (value: number) =>
    Number.isFinite(value) ? String(value) : "unbounded";
  return `${size(array.rows)} x ${size(array.cols)}`;
}

/** The number an array cell contributes, or undefined for any other value. */
function numberOf(cell: ArrayCellResult | undefined): number | undefined {
  if (cell?.type !== "value") {
    return undefined;
  }
  switch (cell.result.type) {
    case "number":
      return cell.result.value;
    case "infinity":
      return cell.result.sign === "positive" ? Infinity : -Infinity;
    default:
      return undefined;
  }
}

/**
 * SUMPRODUCT(array1, [array2], ...)
 *
 * Multiplies the corresponding cells of arrays of the same size and returns the
 * sum of the products: `SUMPRODUCT(B2:B6, C2:C6)` is `B2*C2 + ... + B6*C6`.
 * With one array it is the sum of its numbers.
 *
 * - Every array must have the same number of rows and columns; otherwise the
 *   result is #VALUE!, as in Excel. A single value is a 1x1 array, so
 *   `SUMPRODUCT(2, 3)` is 6.
 * - A cell that is not a number in any of the arrays (text, a logical value or
 *   an empty cell) makes its product 0, as in Excel. That includes TRUE and
 *   FALSE: `SUMPRODUCT(A1:A9="x", B1:B9)` is 0 in Excel too.
 * - An error anywhere in the arrays is the result, as in Excel; with several,
 *   the first array's first error (row by row) wins.
 * - Excel workbooks often turn conditions into numbers with `--(A1:A9="x")` or
 *   `(A1:A9="x")*(B1:B9>0)`. Booleans are not numbers in this engine, so those
 *   expressions give #VALUE! cells and therefore a #VALUE! result. Use IF:
 *   `SUMPRODUCT(IF(A1:A9="x", B1:B9, 0), C1:C9)`, or SUMIFS.
 * - INFINITY and -INFINITY are multiplied and added as infinities; a product
 *   of 0 and an infinity, or a sum of opposite infinities, is #NUM!.
 * - Unbounded ranges such as `SUMPRODUCT(A:A, B:B)` work: their cells are read
 *   by the populated ones, and a position that is empty in one of the arrays
 *   adds nothing.
 * - An omitted argument (`SUMPRODUCT(A1:A3,,B1:B3)`) gives #VALUE!.
 */
export const SUMPRODUCT: FunctionDefinition = {
  name: "SUMPRODUCT",
  evaluate: function (node, context): FunctionEvaluationResult {
    const error = (
      err: FormulaError,
      message: string
    ): ErrorEvaluationResult => ({
      type: "error",
      err,
      message: `SUMPRODUCT: ${message}`,
      errAddress: context.dependencyNode,
    });

    if (node.args.length === 0) {
      return error(FormulaError.VALUE, "requires at least 1 argument");
    }
    const argNodes: ASTNode[] = [];
    for (let index = 0; index < node.args.length; index++) {
      const argNode = getSuppliedArgument(node, index);
      if (!argNode) {
        return error(FormulaError.VALUE, `array${index + 1} is omitted`);
      }
      argNodes.push(argNode);
    }

    // Evaluate every argument before returning any of them, so that one pass
    // registers all dependencies.
    const results = argNodes.map((arg) => this.evaluateNode(arg, context));
    for (const result of results) {
      if (result.type === "awaiting-evaluation") {
        return result;
      }
    }
    for (const result of results) {
      if (result.type === "error") {
        return result;
      }
    }

    const arrays = results.map((result) =>
      toArrayArgument(
        this,
        result as Exclude<FunctionEvaluationResult, AwaitingEvaluationResult>,
        context
      )
    );
    const [first] = arrays;
    for (const [index, array] of arrays.entries()) {
      if (array.rows !== first!.rows || array.cols !== first!.cols) {
        return error(
          FormulaError.VALUE,
          `array${index + 1} is ${describeShape(array)} but array1 is ` +
            `${describeShape(first!)}; the arrays must have the same size`
        );
      }
    }

    // Read every array once, keyed by position.
    const key = Number.isFinite(first!.cols)
      ? (x: number, y: number): number | string => y * first!.cols + x
      : (x: number, y: number): number | string => `${x},${y}`;
    const cellsByArray: Map<number | string, ArrayCellResult>[] = [];
    let awaiting: AwaitingEvaluationResult | undefined;
    for (const array of arrays) {
      const read = array.readCells(context);
      if (read.type === "awaiting-evaluation") {
        awaiting ??= read;
        continue;
      }
      if (read.type === "error") {
        return read;
      }
      let firstError:
        | { x: number; y: number; result: ErrorEvaluationResult }
        | undefined;
      const cells = new Map<number | string, ArrayCellResult>();
      for (const cell of read.cells) {
        if (
          cell.result.type === "error" &&
          (!firstError ||
            cell.y < firstError.y ||
            (cell.y === firstError.y && cell.x < firstError.x))
        ) {
          firstError = { x: cell.x, y: cell.y, result: cell.result };
        }
        cells.set(key(cell.x, cell.y), cell.result);
      }
      if (firstError && !awaiting) {
        return firstError.result;
      }
      cellsByArray.push(cells);
    }
    if (awaiting) {
      return awaiting;
    }

    // Only positions where every array has a number add to the sum, so walk
    // the array with the fewest cells.
    const driver = cellsByArray.reduce((smallest, cells) =>
      cells.size < smallest.size ? cells : smallest
    );
    let sum = 0;
    for (const position of driver.keys()) {
      let product = 1;
      let isNumber = true;
      for (const cells of cellsByArray) {
        const value = numberOf(cells.get(position));
        if (value === undefined) {
          isNumber = false;
          break;
        }
        product *= value;
      }
      if (!isNumber) {
        continue;
      }
      sum += product;
      if (Number.isNaN(sum)) {
        return error(
          FormulaError.NUM,
          "a product of 0 and an infinity, or a sum of opposite infinities, is undefined"
        );
      }
    }

    if (sum === Infinity || sum === -Infinity) {
      return {
        type: "value",
        result: { type: "infinity", sign: sum > 0 ? "positive" : "negative" },
      };
    }
    return { type: "value", result: { type: "number", value: sum } };
  },
};
