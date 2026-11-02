import {
  FormulaError,
  type AwaitingEvaluationResult,
  type CellAddress,
  type ErrorEvaluationResult,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpreadsheetRange,
  type ValueEvaluationResult,
} from "../../core/types";
import {
  captureEvaluationErrors,
  getRangeIntersection,
} from "../../core/utils";
import { getSpillAreaShape } from "../../evaluator/array-shape";
import { hasCriteriaRangeReader } from "../../evaluator/criteria-range-cache";
import type { EvaluationContext } from "../../evaluator/evaluation-context";
import { EvaluationError } from "../../evaluator/evaluation-error";
import type { FormulaEvaluator } from "../../evaluator/formula-evaluator";

/**
 * Helpers for array functions (TOCOL, TOROW, HSTACK) that view an evaluated
 * argument as a two-dimensional array of cells.
 *
 * Cells are addressed by offsets `x` (column) and `y` (row) from the array's
 * top-left cell. A dimension is `Infinity` when the array is unbounded, as in
 * `A:A`, `2:2` or `A2:A`.
 *
 * The engine reads an empty cell as the empty string, so a blank cell and a cell
 * holding `""` look the same: both count as blank here.
 */

/** A cell's value or error; never awaiting-evaluation. */
export type ArrayCellResult = Exclude<
  SingleEvaluationResult,
  AwaitingEvaluationResult
>;

export type ArrayCell = {
  x: number;
  y: number;
  result: ArrayCellResult;
};

type ArrayCells = {
  type: "cells";
  cells: ArrayCell[];
};

/** An evaluated argument viewed as an array of cells. */
export type ArrayArgument = {
  rows: number;
  cols: number;
  /**
   * True when `readCells` returns only the populated cells (raw ranges and
   * unbounded arrays): every cell missing from the result is blank. Otherwise
   * `readCells` returns every cell of the array.
   */
  sparse: boolean;
  /**
   * Reads one cell lazily, adding its dependencies to `context`. The offsets
   * must lie inside the array.
   */
  cellAt: (
    x: number,
    y: number,
    context: EvaluationContext
  ) => SingleEvaluationResult;
  /**
   * Reads every cell once. Returns the first awaiting-evaluation result instead
   * when any cell is not computed yet, after reading the others, so that one
   * pass registers every dependency. Order is unspecified.
   */
  readCells: (context: EvaluationContext) => ArrayCells | ErrorEvaluationResult;
};

/** True for a cell that reads as the empty string. */
export function isBlankResult(result: SingleEvaluationResult): boolean {
  return (
    result.type === "value" &&
    result.result.type === "string" &&
    result.result.value === ""
  );
}

/** A fresh blank cell result. */
export function blankResult(): ValueEvaluationResult {
  return { type: "value", result: { type: "string", value: "" } };
}

/**
 * Wraps an evaluated argument as an array. A single value or an error is a 1x1
 * array; a spilled result keeps its shape. Handle awaiting-evaluation results
 * before calling this.
 *
 * Arrays are read in one of two ways:
 * - Raw ranges and unbounded arrays are read through `evaluateAllCells`, which
 *   is sparse: only populated cells are returned and the reader depends on a
 *   range node that tracks later edits. This is the only way to read an
 *   unbounded array. Cells that an unbounded computed array, such as `A:A*2`,
 *   leaves out of `evaluateAllCells` count as blank. A raw range is recognized
 *   by its registered range reader rather than by `sourceRange`, which computed
 *   results such as ROW(A1:A3) carry too.
 * - Other spilled results are computed arrays, which are dense and cheap to
 *   read cell by cell over their bounded area.
 */
export function toArrayArgument(
  evaluator: FormulaEvaluator,
  result: Exclude<FunctionEvaluationResult, AwaitingEvaluationResult>,
  context: EvaluationContext
): ArrayArgument {
  if (result.type !== "spilled-values") {
    // A value keeps only its content: the array is a value, not a reference.
    const cell: ArrayCellResult =
      result.type === "value"
        ? { type: "value", result: result.result }
        : result;
    return {
      rows: 1,
      cols: 1,
      sparse: false,
      cellAt: () => cell,
      readCells: () => ({
        type: "cells",
        cells: [{ x: 0, y: 0, result: cell }],
      }),
    };
  }

  const spilled = result;
  const { rows, cols } = getSpillAreaShape(
    spilled.spillArea(context.cellAddress)
  );
  const sparse =
    !Number.isFinite(rows) ||
    !Number.isFinite(cols) ||
    hasCriteriaRangeReader(spilled, evaluator);

  const cellAt = (x: number, y: number, cellContext: EvaluationContext) =>
    spilled.evaluate({ x, y }, cellContext);

  return {
    rows,
    cols,
    sparse,
    cellAt,
    readCells: (readContext) => {
      if (!sparse) {
        return readDenseCells(rows, cols, cellAt, readContext);
      }

      const values = spilled.evaluateAllCells.call(evaluator, {
        context: readContext,
        evaluate: spilled.evaluate,
        origin: readContext.cellAddress,
        lookupOrder: "row-major",
      });
      if (values.type !== "values") {
        return values;
      }

      const cells: ArrayCell[] = [];
      let awaiting: AwaitingEvaluationResult | undefined;
      for (const { result: cell, relativePos } of values.values) {
        if (cell.type === "awaiting-evaluation") {
          awaiting ??= cell;
          continue;
        }
        const { x, y } = relativePos;
        if (x < 0 || y < 0 || x >= cols || y >= rows) {
          continue;
        }
        cells.push({ x, y, result: cell });
      }
      return awaiting ?? { type: "cells", cells };
    },
  };
}

/**
 * Reads every cell of a bounded array, row by row. Thrown evaluation errors
 * become error cells. Keeps reading past cells that are awaiting evaluation so
 * one pass registers every dependency, then returns the first of them.
 */
function readDenseCells(
  rows: number,
  cols: number,
  cellAt: ArrayArgument["cellAt"],
  context: EvaluationContext
): ArrayCells | AwaitingEvaluationResult {
  const cells: ArrayCell[] = [];
  let awaiting: AwaitingEvaluationResult | undefined;
  for (let y = 0; y < rows; y++) {
    for (let x = 0; x < cols; x++) {
      const result = captureEvaluationErrors(context.dependencyNode, () =>
        cellAt(x, y, context)
      );
      if (result.type === "awaiting-evaluation") {
        awaiting ??= result;
        continue;
      }
      cells.push({ x, y, result });
    }
  }
  return awaiting ?? { type: "cells", cells };
}

/** The area of a `rows` x `cols` result spilled from `origin`. */
export function createSpillArea(
  origin: CellAddress,
  rows: number,
  cols: number
): SpreadsheetRange {
  return {
    start: { col: origin.colIndex, row: origin.rowIndex },
    end: {
      col: Number.isFinite(cols)
        ? { type: "number", value: origin.colIndex + cols - 1 }
        : { type: "infinity", sign: "positive" },
      row: Number.isFinite(rows)
        ? { type: "number", value: origin.rowIndex + rows - 1 }
        : { type: "infinity", sign: "positive" },
    },
  };
}

/**
 * Offsets from a spill's origin, with inclusive ends that are `Infinity` when
 * unbounded.
 */
type SpillWindow = {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
};

/**
 * The offsets from `origin` of the cells that `evaluateAllCells` has to return:
 * the spill area, limited to `intersection` when one is given. Undefined when
 * they do not overlap.
 */
export function getSpillWindow(
  spillArea: SpreadsheetRange,
  origin: CellAddress,
  intersection?: SpreadsheetRange
): SpillWindow | undefined {
  const range = intersection
    ? getRangeIntersection(spillArea, intersection)
    : spillArea;
  if (!range) {
    return undefined;
  }
  return {
    x0: range.start.col - origin.colIndex,
    y0: range.start.row - origin.rowIndex,
    x1:
      range.end.col.type === "number"
        ? range.end.col.value - origin.colIndex
        : Infinity,
    y1:
      range.end.row.type === "number"
        ? range.end.row.value - origin.rowIndex
        : Infinity,
  };
}

function invalidOption(
  name: string,
  message: string,
  context: EvaluationContext
): ErrorEvaluationResult {
  return {
    type: "error",
    err: FormulaError.VALUE,
    message: `${name} ${message}`,
    errAddress: context.dependencyNode,
  };
}

/**
 * Reads a scalar whole-number option such as TOCOL's `ignore`. Numbers are
 * truncated toward zero and logical values count as 1 and 0. A blank value
 * counts as 0, since an empty cell reads as `""`. Other text and arrays give
 * #VALUE!; an error is returned as is. Handle awaiting-evaluation first.
 */
export function readIntegerOption(
  result: Exclude<FunctionEvaluationResult, AwaitingEvaluationResult>,
  name: string,
  context: EvaluationContext
): number | ErrorEvaluationResult {
  if (result.type === "error") {
    return result;
  }
  if (result.type === "spilled-values") {
    return invalidOption(name, "must be a single value", context);
  }
  const value = result.result;
  switch (value.type) {
    case "number":
      return Math.trunc(value.value);
    case "boolean":
      return value.value ? 1 : 0;
    case "infinity":
      return value.sign === "positive" ? Infinity : -Infinity;
    case "string":
      if (value.value === "") {
        return 0;
      }
      return invalidOption(name, "must be a number, got text", context);
  }
}

/**
 * Reads a scalar logical option such as TOCOL's `scan_by_column`. Numbers are
 * TRUE unless 0, and a blank value is FALSE. Other text and arrays give
 * #VALUE!; an error is returned as is. Handle awaiting-evaluation first.
 */
export function readLogicalOption(
  result: Exclude<FunctionEvaluationResult, AwaitingEvaluationResult>,
  name: string,
  context: EvaluationContext
): boolean | ErrorEvaluationResult {
  if (result.type === "error") {
    return result;
  }
  if (result.type === "spilled-values") {
    return invalidOption(name, "must be a single logical value", context);
  }
  const value = result.result;
  switch (value.type) {
    case "boolean":
      return value.value;
    case "number":
      return value.value !== 0;
    case "infinity":
      return true;
    case "string":
      if (value.value === "") {
        return false;
      }
      return invalidOption(name, "must be a logical value, got text", context);
  }
}

/** The result for a read outside a spilled result. */
export function outsideResultError(
  name: string,
  context: EvaluationContext
): ErrorEvaluationResult {
  return {
    type: "error",
    err: FormulaError.REF,
    message: `${name}: requested cell is outside the result`,
    errAddress: context.dependencyNode,
  };
}

/** Thrown by `evaluateAllCells` when it would have to list infinitely many cells. */
export function throwInfiniteCells(name: string): never {
  throw new EvaluationError(
    FormulaError.REF,
    `${name}: can not evaluate all cells over an infinite range`
  );
}
