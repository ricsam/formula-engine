import {
  FormulaError,
  type CellAddress,
  type CellInRangeResult,
  type ErrorEvaluationResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpreadsheetRange,
} from "../../../core/types";
import { getRangeIntersection, isRangeOneCell } from "../../../core/utils";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import { EvaluationError } from "../../../evaluator/evaluation-error";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import { getSuppliedArgument, type FunctionNode } from "../../../parser/ast";

const valueError = (
  message: string,
  context: EvaluationContext
): ErrorEvaluationResult => ({
  type: "error",
  err: FormulaError.VALUE,
  message,
  errAddress: context.dependencyNode,
});

/**
 * Reads an optional argument as a single value, or returns `undefined` when the
 * call omits it so that the caller applies the default.
 */
function readSingleArgument(
  evaluator: FormulaEvaluator,
  node: FunctionNode,
  index: number,
  name: string,
  context: EvaluationContext
): SingleEvaluationResult | undefined {
  const argument = getSuppliedArgument(node, index);
  if (!argument) {
    return undefined;
  }
  const result = evaluator.evaluateNode(argument, context);
  if (result.type === "spilled-values") {
    return valueError(
      `SEQUENCE ${name} must be a single value, not an array`,
      context
    );
  }
  return result;
}

/** Reads the rows or columns argument: a whole number of at least 1, or INFINITY. */
function readDimension(
  evaluator: FormulaEvaluator,
  node: FunctionNode,
  index: number,
  name: "rows" | "columns",
  context: EvaluationContext
): number | ErrorEvaluationResult {
  const result = readSingleArgument(evaluator, node, index, name, context);
  if (!result) {
    return 1;
  }
  if (result.type !== "value") {
    return result;
  }
  const value = result.result;
  if (value.type === "infinity" && value.sign === "positive") {
    return Infinity;
  }
  if (value.type !== "number") {
    return valueError(
      `SEQUENCE ${name} must be a number or INFINITY`,
      context
    );
  }
  const count = Math.floor(value.value);
  if (!(count >= 1)) {
    return valueError(`SEQUENCE ${name} must be at least 1`, context);
  }
  return count;
}

/** Reads the start or step argument: a number. */
function readNumber(
  evaluator: FormulaEvaluator,
  node: FunctionNode,
  index: number,
  name: "start" | "step",
  context: EvaluationContext
): number | ErrorEvaluationResult {
  const result = readSingleArgument(evaluator, node, index, name, context);
  if (!result) {
    return 1;
  }
  if (result.type !== "value") {
    return result;
  }
  if (result.result.type !== "number") {
    return valueError(`SEQUENCE ${name} must be a number`, context);
  }
  return result.result.value;
}

/**
 * SEQUENCE(rows, [columns], [start], [step])
 * Generates a sequence of numbers in an array, filled row by row.
 *
 * @param rows - The number of rows (default: 1), or INFINITY
 * @param columns - [Optional] The number of columns (default: 1), or INFINITY
 * @param start - [Optional] The first number in the sequence (default: 1)
 * @param step - [Optional] The amount to increment each subsequent value (default: 1)
 *
 * An omitted argument takes its default, as in Excel: `SEQUENCE(,3)` is the row
 * {1,2,3} and `SEQUENCE(3,,10)` is the column {10;11;12}. Rows and columns are
 * truncated to whole numbers and must be at least 1.
 *
 * Every argument must be a single value. Excel lifts SEQUENCE over an array
 * argument (one sequence per element, each reduced to its first value); this
 * engine returns #VALUE! for an array argument instead.
 *
 * An unbounded sequence still fills row by row while the number of columns is
 * finite. With INFINITY columns, every row of a multi-row sequence repeats the
 * first row, and an INFINITY x INFINITY sequence holds `start + (x + y) * step`
 * at column offset x and row offset y.
 */
export const SEQUENCE: FunctionDefinition = {
  name: "SEQUENCE",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 1 || node.args.length > 4) {
      return valueError("SEQUENCE function takes 1 to 4 arguments", context);
    }

    const rows = readDimension(this, node, 0, "rows", context);
    if (typeof rows !== "number") {
      return rows;
    }
    const columns = readDimension(this, node, 1, "columns", context);
    if (typeof columns !== "number") {
      return columns;
    }
    const start = readNumber(this, node, 2, "start", context);
    if (typeof start !== "number") {
      return start;
    }
    const step = readNumber(this, node, 3, "step", context);
    if (typeof step !== "number") {
      return step;
    }

    const isRowsInfinite = rows === Infinity;
    const isColumnsInfinite = columns === Infinity;

    const spillArea = (origin: CellAddress): SpreadsheetRange => ({
      start: {
        col: origin.colIndex,
        row: origin.rowIndex,
      },
      end: {
        col: isColumnsInfinite
          ? { type: "infinity", sign: "positive" }
          : { type: "number", value: origin.colIndex + columns - 1 },
        row: isRowsInfinite
          ? { type: "infinity", sign: "positive" }
          : { type: "number", value: origin.rowIndex + rows - 1 },
      },
    });

    if (isRangeOneCell(spillArea(context.cellAddress))) {
      return {
        type: "value",
        result: { type: "number", value: start },
      };
    }

    /** The value at column offset x and row offset y of the sequence. */
    const valueAt = (
      x: number,
      y: number,
      readerContext: EvaluationContext
    ): SingleEvaluationResult => {
      if (x < 0 || y < 0 || x >= columns || y >= rows) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "SEQUENCE: requested cell is outside the sequence",
          errAddress: readerContext.dependencyNode,
        };
      }

      let sequenceIndex: number;
      if (!isColumnsInfinite) {
        // Row by row: (0,0), (1,0), ..., (0,1), (1,1), ...
        sequenceIndex = y * columns + x;
      } else if (isRowsInfinite) {
        // No row-by-row order exists over unbounded rows and columns.
        sequenceIndex = y + x;
      } else {
        // A row never ends, so every row repeats the first one.
        sequenceIndex = x;
      }

      return {
        type: "value",
        result: { type: "number", value: start + sequenceIndex * step },
      };
    };

    return {
      type: "spilled-values",
      spillArea,
      source: "SEQUENCE function",
      evaluate: (spillOffset, readerContext) =>
        valueAt(spillOffset.x, spillOffset.y, readerContext),
      // Dense. The values are read with this sequence's own `valueAt`:
      // `options.evaluate` may belong to a wrapper that maps the cells it gets.
      evaluateAllCells: function ({
        intersection,
        context,
        origin,
        lookupOrder,
      }) {
        let range = spillArea(origin);
        if (intersection) {
          const overlap = getRangeIntersection(range, intersection);
          if (!overlap) {
            return { type: "values", values: [] };
          }
          range = overlap;
        }
        if (
          range.end.row.type === "infinity" ||
          range.end.col.type === "infinity"
        ) {
          throw new EvaluationError(
            FormulaError.REF,
            `Can not evaluate all cells over an infinite range`
          );
        }

        const firstX = range.start.col - origin.colIndex;
        const lastX = range.end.col.value - origin.colIndex;
        const firstY = range.start.row - origin.rowIndex;
        const lastY = range.end.row.value - origin.rowIndex;
        const results: CellInRangeResult[] = [];
        const read = (x: number, y: number) => {
          results.push({
            result: valueAt(x, y, context),
            relativePos: { x, y },
          });
        };
        if (lookupOrder === "col-major") {
          for (let x = firstX; x <= lastX; x++) {
            for (let y = firstY; y <= lastY; y++) read(x, y);
          }
        } else {
          for (let y = firstY; y <= lastY; y++) {
            for (let x = firstX; x <= lastX; x++) read(x, y);
          }
        }

        return { type: "values", values: results };
      },
    };
  },
};
