import {
  FormulaError,
  type CellInRangeResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpilledValuesEvaluationResult,
  type ValueEvaluationResult,
} from "../../core/types";
import { captureEvaluationErrors } from "../../core/utils";
import { hasCriteriaRangeReader } from "../../evaluator/criteria-range-cache";
import type { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  blankResult,
  getSpillWindow,
  type ArrayCellResult,
} from "../array/array-utils";

/**
 * Shared implementation of the IS functions (ISNUMBER, ISBLANK): each tests one
 * value and returns TRUE or FALSE, never an error.
 *
 * - An error value is tested like any other value (and is not a number nor
 *   blank), so `ISNUMBER(1/"x")` is FALSE rather than #VALUE!.
 * - A spilled argument, such as a range or `SEQUENCE(3)`, is tested element by
 *   element and the result spills with the argument's shape.
 * - Aggregates read every cell of a bounded result through `evaluateAllCells`,
 *   empty cells of a range included, so `SUM(IF(ISBLANK(A1:A9), 1, 0))` counts
 *   the blank cells. An unbounded result such as `ISBLANK(A:A)` reports only the
 *   cells its argument reports: for a range, the populated ones.
 */

/** A bounded result larger than this is read sparsely, like an unbounded one. */
const MAX_DENSE_CELLS = 1 << 24;

function booleanResult(value: boolean): ValueEvaluationResult {
  return { type: "value", result: { type: "boolean", value } };
}

/**
 * The spilled result of testing every cell of `argument`.
 */
function createTestedArray(
  name: string,
  argument: SpilledValuesEvaluationResult,
  test: (cell: ArrayCellResult) => boolean,
  evaluator: FormulaEvaluator
): SpilledValuesEvaluationResult {
  // The cells a raw range leaves out of evaluateAllCells are empty.
  const isRange = hasCriteriaRangeReader(argument, evaluator);

  return {
    type: "spilled-values",
    source: `${name} function`,
    spillArea: (origin) => argument.spillArea(origin),
    evaluate: (offset, context) => {
      const cell = captureEvaluationErrors(context.dependencyNode, () =>
        argument.evaluate(offset, context)
      );
      if (cell.type === "awaiting-evaluation") {
        return cell;
      }
      return booleanResult(test(cell));
    },
    evaluateAllCells: function (options) {
      // The argument's own cells: forwarding options.evaluate, which is this
      // result's evaluate, would test them twice.
      const cells = argument.evaluateAllCells.call(this, {
        ...options,
        evaluate: argument.evaluate,
      });
      if (cells.type !== "values") {
        return cells;
      }
      const tested = (cell: SingleEvaluationResult): SingleEvaluationResult =>
        cell.type === "awaiting-evaluation" ? cell : booleanResult(test(cell));

      const window = getSpillWindow(
        argument.spillArea(options.origin),
        options.origin,
        options.intersection
      );
      if (!window) {
        return { type: "values", values: [] };
      }
      const width = window.x1 - window.x0 + 1;
      const height = window.y1 - window.y0 + 1;
      if (!Number.isFinite(width * height) || width * height > MAX_DENSE_CELLS) {
        return {
          type: "values",
          values: cells.values.map((cell) => ({
            relativePos: cell.relativePos,
            result: tested(cell.result),
          })),
        };
      }

      // Report every cell of the window. Cells the argument left out are empty
      // in a range; any other array is asked for them one by one.
      const reported = new Map<number, SingleEvaluationResult>();
      for (const { relativePos, result } of cells.values) {
        const x = relativePos.x - window.x0;
        const y = relativePos.y - window.y0;
        if (x >= 0 && y >= 0 && x < width && y < height) {
          reported.set(y * width + x, result);
        }
      }
      const blank = booleanResult(test(blankResult()));
      const values: CellInRangeResult[] = [];
      const visit = (x: number, y: number) => {
        const cell = reported.get((y - window.y0) * width + (x - window.x0));
        let result: SingleEvaluationResult;
        if (cell) {
          result = tested(cell);
        } else if (isRange) {
          result = blank;
        } else {
          result = tested(
            captureEvaluationErrors(options.context.dependencyNode, () =>
              argument.evaluate({ x, y }, options.context)
            )
          );
        }
        values.push({ relativePos: { x, y }, result });
      };
      if (options.lookupOrder === "row-major") {
        for (let y = window.y0; y <= window.y1; y++) {
          for (let x = window.x0; x <= window.x1; x++) visit(x, y);
        }
      } else {
        for (let x = window.x0; x <= window.x1; x++) {
          for (let y = window.y0; y <= window.y1; y++) visit(x, y);
        }
      }
      return { type: "values", values };
    },
  };
}

/**
 * Creates an IS function that returns `test(value)` for its single argument.
 */
export function createIsFunction(
  name: string,
  test: (cell: ArrayCellResult) => boolean
): FunctionDefinition {
  return {
    name,
    evaluate: function (node, context): FunctionEvaluationResult {
      if (node.args.length !== 1) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: `${name} function takes exactly 1 argument`,
          errAddress: context.dependencyNode,
        };
      }

      const result = this.evaluateNode(node.args[0]!, context);
      if (result.type === "awaiting-evaluation") {
        return result;
      }
      if (result.type === "spilled-values") {
        return createTestedArray(name, result, test, this);
      }
      return booleanResult(test(result));
    },
  };
}
