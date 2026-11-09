import {
  FormulaError,
  type CellAddress,
  type CellInRangeResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SpreadsheetRange,
} from "../../../core/types";
import { getSpillAreaShape } from "../../../evaluator/array-shape";
import { getSuppliedArgument } from "../../../parser/ast";
import {
  createSpillArea,
  getSpillWindow,
  outsideResultError,
} from "../array-utils";

/**
 * TRANSPOSE(array)
 * Swaps the rows and columns of an array: the first row of `array` becomes the
 * first column of the result.
 *
 * - The result spills: a 2 x 3 array gives a 3 x 2 result. A single value, or a
 *   one-cell array, is returned as a single value.
 * - The result maps each of its cells to one cell of `array` and reads nothing
 *   up front, so it follows edits of the source cells like any spill.
 * - Errors in `array` stay in their cells; an error argument is returned.
 * - Blank cells of a range read as empty text, as everywhere in the engine.
 *   Excel shows 0 for them.
 * - Aggregates read the cells `array` provides, transposed: for a range, its
 *   populated cells.
 * - Unbounded arrays stay lazy: `TRANSPOSE(A:A)` is a single row with unbounded
 *   columns.
 */
export const TRANSPOSE: FunctionDefinition = {
  name: "TRANSPOSE",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "TRANSPOSE function takes exactly 1 argument",
        errAddress: context.dependencyNode,
      };
    }
    const arrayNode = getSuppliedArgument(node, 0);
    if (!arrayNode) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "TRANSPOSE: array is required",
        errAddress: context.dependencyNode,
      };
    }

    const array = this.evaluateNode(arrayNode, context);
    if (array.type !== "spilled-values") {
      // A single value, an error or a pending evaluation.
      return array;
    }

    const shape = getSpillAreaShape(array.spillArea(context.cellAddress));
    const rows = shape.cols;
    const cols = shape.rows;
    if (rows === 1 && cols === 1) {
      return array.evaluate({ x: 0, y: 0 }, context);
    }
    const spillArea = (origin: CellAddress) =>
      createSpillArea(origin, rows, cols);

    return {
      type: "spilled-values",
      source: "TRANSPOSE function",
      spillArea,
      evaluate: (offset, cellContext) => {
        if (
          offset.x < 0 ||
          offset.y < 0 ||
          offset.x >= cols ||
          offset.y >= rows
        ) {
          return outsideResultError("TRANSPOSE", context);
        }
        return array.evaluate({ x: offset.y, y: offset.x }, cellContext);
      },
      evaluateAllCells: function ({
        intersection,
        context: readContext,
        origin,
        lookupOrder,
      }) {
        const window = getSpillWindow(spillArea(origin), origin, intersection);
        if (!window) {
          return { type: "values", values: [] };
        }

        // The same block of `array`, with rows and columns swapped. `array`
        // reports positions relative to the origin too, also within an
        // intersection.
        const end = (offset: number): SpreadsheetRange["end"]["col"] =>
          Number.isFinite(offset)
            ? { type: "number", value: offset }
            : { type: "infinity", sign: "positive" };
        const sourceIntersection: SpreadsheetRange = {
          start: {
            col: origin.colIndex + window.y0,
            row: origin.rowIndex + window.x0,
          },
          end: {
            col: end(origin.colIndex + window.y1),
            row: end(origin.rowIndex + window.x1),
          },
        };
        // Pass the array's own evaluate: forwarding the caller's could apply a
        // wrapper's mapping twice.
        const cells = array.evaluateAllCells.call(this, {
          context: readContext,
          origin,
          evaluate: array.evaluate,
          intersection: sourceIntersection,
          lookupOrder: lookupOrder === "row-major" ? "col-major" : "row-major",
        });
        if (cells.type !== "values") {
          return cells;
        }

        const values: CellInRangeResult[] = [];
        for (const cell of cells.values) {
          const x = cell.relativePos.y;
          const y = cell.relativePos.x;
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
      },
    };
  },
};
