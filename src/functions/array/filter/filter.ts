import {
  FormulaError,
  type AwaitingEvaluationResult,
  type CellAddress,
  type CellInRangeResult,
  type ErrorEvaluationResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "../../../core/types";
import { captureEvaluationErrors } from "../../../core/utils";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import { getSuppliedArgument } from "../../../parser/ast";
import { isTruthy } from "../../logical/if/if";
import {
  createSpillArea,
  getSpillWindow,
  outsideResultError,
  toArrayArgument,
  type ArrayCell,
} from "../array-utils";

function describeSize(size: number): string {
  return Number.isFinite(size) ? String(size) : "unbounded";
}

/**
 * FILTER(array, include, [if_empty])
 * Returns the rows (or columns) of an array for which `include` is true.
 *
 * @param array - The range or array to filter. A single value is a 1x1 array.
 * @param include - One column as tall as `array`, which selects rows, or one
 *   row as wide as `array`, which selects columns. Rows are tried first, so a
 *   one-cell `include` selects the single row of a one-row array. Any other
 *   shape gives #VALUE!.
 * @param if_empty - [Optional] Returned, as is, when nothing is selected.
 *   Without it an empty selection is #CALC!, as in Excel; an omitted
 *   `if_empty` (a trailing comma) keeps that default.
 *
 * - `include` values are tested like IF conditions: FALSE, 0 and empty text
 *   (an empty cell) exclude, every other value includes. Excel gives #VALUE!
 *   for text in `include`; this engine includes non-empty text, as IF treats
 *   it as true.
 * - An error in `include` is the result, even where `array` is not selected,
 *   as in Excel. Errors in `array` stay in the cells they reach.
 * - `include` is read once, when FILTER is evaluated; the result maps each of
 *   its cells to a selected cell of `array` and reads it when it is read, so a
 *   one-cell result is returned as a single value.
 * - Conditions are usually comparisons such as `B2:B9="x"`. Booleans are not
 *   numbers in this engine, so Excel's `(B2:B9="x")*(C2:C9>5)` gives #VALUE!
 *   cells and therefore a #VALUE! result; combine conditions with IF instead:
 *   `IF(B2:B9="x", C2:C9>5)` is true where both are.
 * - Unbounded ranges such as `FILTER(A:A, B:B="x")` work: `include` is read by
 *   its populated cells (and the cells that operators compute from them), and
 *   the cells it leaves out count as blank, so they are not selected. Their
 *   heights are both unbounded, so such ranges must start in the same row to
 *   line up.
 */
export const FILTER: FunctionDefinition = {
  name: "FILTER",
  evaluate: function (node, context): FunctionEvaluationResult {
    const error = (
      err: FormulaError,
      message: string
    ): ErrorEvaluationResult => ({
      type: "error",
      err,
      message: `FILTER: ${message}`,
      errAddress: context.dependencyNode,
    });

    if (node.args.length < 2 || node.args.length > 3) {
      return error(FormulaError.VALUE, "takes 2 or 3 arguments");
    }
    const arrayNode = getSuppliedArgument(node, 0);
    if (!arrayNode) {
      return error(FormulaError.VALUE, "array is required");
    }
    const includeNode = getSuppliedArgument(node, 1);
    if (!includeNode) {
      return error(FormulaError.VALUE, "include is required");
    }

    // Evaluate both arrays before returning either, so that one pass registers
    // all dependencies. if_empty is only evaluated when it is needed.
    const arrayResult = this.evaluateNode(arrayNode, context);
    const includeResult = this.evaluateNode(includeNode, context);
    if (arrayResult.type === "awaiting-evaluation") {
      return arrayResult;
    }
    if (includeResult.type === "awaiting-evaluation") {
      return includeResult;
    }
    if (arrayResult.type === "error") {
      return arrayResult;
    }
    if (includeResult.type === "error") {
      return includeResult;
    }

    const array = toArrayArgument(this, arrayResult, context);
    const include = toArrayArgument(this, includeResult, context);

    let byRow: boolean;
    if (include.cols === 1 && include.rows === array.rows) {
      byRow = true;
    } else if (include.rows === 1 && include.cols === array.cols) {
      byRow = false;
    } else {
      return error(
        FormulaError.VALUE,
        `include must be one column as tall as array (${describeSize(
          array.rows
        )} rows) or one row as wide as it (${describeSize(
          array.cols
        )} columns), got ${describeSize(include.rows)} x ${describeSize(
          include.cols
        )}`
      );
    }

    const read = include.readCells(context);
    if (read.type !== "cells") {
      return read;
    }
    const indexOf = (cell: ArrayCell) => (byRow ? cell.y : cell.x);
    const includeCells = [...read.cells].sort(
      (a, b) => indexOf(a) - indexOf(b)
    );
    /** The selected rows (or columns) of `array`, in order. */
    const selected: number[] = [];
    for (const cell of includeCells) {
      if (cell.result.type === "error") {
        return cell.result;
      }
      if (isTruthy(cell.result.result)) {
        selected.push(indexOf(cell));
      }
    }

    if (selected.length === 0) {
      const ifEmptyNode = getSuppliedArgument(node, 2);
      if (!ifEmptyNode) {
        return error(FormulaError.CALC, "no rows or columns are selected");
      }
      return this.evaluateNode(ifEmptyNode, context);
    }

    const rows = byRow ? selected.length : array.rows;
    const cols = byRow ? array.cols : selected.length;
    const cellAt = (
      x: number,
      y: number,
      cellContext: EvaluationContext
    ): SingleEvaluationResult =>
      byRow
        ? array.cellAt(x, selected[y]!, cellContext)
        : array.cellAt(selected[x]!, y, cellContext);

    if (rows === 1 && cols === 1) {
      return cellAt(0, 0, context);
    }

    const spillArea = (origin: CellAddress) =>
      createSpillArea(origin, rows, cols);

    return {
      type: "spilled-values",
      source: "FILTER function",
      spillArea,
      evaluate: (offset, cellContext) => {
        if (
          offset.x < 0 ||
          offset.y < 0 ||
          offset.x >= cols ||
          offset.y >= rows
        ) {
          return outsideResultError("FILTER", context);
        }
        return cellAt(offset.x, offset.y, cellContext);
      },
      evaluateAllCells: ({
        intersection,
        context: readContext,
        origin,
        lookupOrder,
      }) => {
        const window = getSpillWindow(spillArea(origin), origin, intersection);
        if (!window) {
          return { type: "values", values: [] };
        }

        const values: CellInRangeResult[] = [];
        if (array.sparse) {
          // Read the populated cells once and keep those of the selection.
          const cells = array.readCells(readContext);
          if (cells.type !== "cells") {
            return cells;
          }
          const position = new Map<number, number>();
          selected.forEach((source, index) => position.set(source, index));
          for (const cell of cells.cells) {
            const index = position.get(byRow ? cell.y : cell.x);
            if (index === undefined) {
              continue;
            }
            const x = byRow ? cell.x : index;
            const y = byRow ? index : cell.y;
            if (
              x >= window.x0 &&
              x <= window.x1 &&
              y >= window.y0 &&
              y <= window.y1
            ) {
              values.push({ result: cell.result, relativePos: { x, y } });
            }
          }
          values.sort(
            lookupOrder === "row-major"
              ? (a, b) =>
                  a.relativePos.y - b.relativePos.y ||
                  a.relativePos.x - b.relativePos.x
              : (a, b) =>
                  a.relativePos.x - b.relativePos.x ||
                  a.relativePos.y - b.relativePos.y
          );
          return { type: "values", values };
        }

        // A dense array is bounded, so these loops end.
        let awaiting: AwaitingEvaluationResult | undefined;
        const visit = (x: number, y: number) => {
          const result = captureEvaluationErrors(
            readContext.dependencyNode,
            () => cellAt(x, y, readContext)
          );
          if (result.type === "awaiting-evaluation") {
            awaiting ??= result;
            return;
          }
          values.push({ result, relativePos: { x, y } });
        };
        if (lookupOrder === "row-major") {
          for (let y = window.y0; y <= window.y1; y++) {
            for (let x = window.x0; x <= window.x1; x++) visit(x, y);
          }
        } else {
          for (let x = window.x0; x <= window.x1; x++) {
            for (let y = window.y0; y <= window.y1; y++) visit(x, y);
          }
        }
        return awaiting ?? { type: "values", values };
      },
    };
  },
};
