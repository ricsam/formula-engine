import {
  FormulaError,
  type AwaitingEvaluationResult,
  type CellAddress,
  type ErrorEvaluationResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpilledValuesEvaluationResult,
} from "../../../core/types";
import { captureEvaluationErrors } from "../../../core/utils";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import { getSuppliedArgument, type FunctionNode } from "../../../parser/ast";
import {
  blankResult,
  createSpillArea,
  getSpillWindow,
  isBlankResult,
  outsideResultError,
  readIntegerOption,
  readLogicalOption,
  throwInfiniteCells,
  toArrayArgument,
  type ArrayArgument,
  type ArrayCell,
  type ArrayCellResult,
} from "../array-utils";

/**
 * TOCOL(array, [ignore], [scan_by_column])
 * Returns the cells of an array as a single column.
 *
 * @param array - The array or range to flatten. A single value, or an error,
 *   is a 1x1 array.
 * @param ignore - [Optional] 0 keeps every cell (default), 1 skips blanks,
 *   2 skips errors, 3 skips blanks and errors. Numbers are truncated, logical
 *   values count as 1/0 and a blank counts as 0; any other value gives #VALUE!.
 * @param scan_by_column - [Optional] FALSE (default) reads the array row by
 *   row, TRUE column by column. Numbers count as TRUE unless 0.
 *
 * Errors that are kept stay as error cells of the result. When nothing is left
 * the result is #CALC!, and a one-cell result is returned as a single value.
 * Omitted options, as in `TOCOL(A1:B2,,TRUE)`, keep their defaults; an omitted
 * array gives #VALUE!.
 *
 * The engine reads an empty cell as `""`, so a cell holding `""` (for example
 * from a formula) is skipped as a blank too.
 *
 * Unbounded arrays such as `A:A`, `A2:A` or `A:B` are read by their populated
 * cells. Skipping blanks drops the empty tail, so the result is bounded. Keeping
 * blanks gives an unbounded result, which works when the scan still reaches
 * every cell: a single row or column, rows of a bounded number of columns read
 * row by row, or columns of a bounded number of rows read column by column.
 * Other unbounded scans, such as `TOCOL(A:B, 0, TRUE)` whose first column never
 * ends, give #VALUE!. An unbounded computed array (for example `A:A*2`) is read
 * through its populated cells as well; the cells it leaves out count as blank.
 *
 * When nothing is skipped, the result only maps each of its cells to one cell of
 * the array and reads nothing up front. Otherwise the array is read once during
 * evaluation and the result serves the surviving cells.
 */
export const TOCOL: FunctionDefinition = {
  name: "TOCOL",
  evaluate: function (node, context): FunctionEvaluationResult {
    return evaluateArrayToVector.call(this, node, context, "column");
  },
};

/** Whether the result is one column (TOCOL) or one row (TOROW). */
type VectorOrientation = "column" | "row";

type IndexedCell = { index: number; result: SingleEvaluationResult };

type VectorCells = { type: "cells"; cells: IndexedCell[] };

/**
 * The flattened array: its cells numbered 0, 1, 2, … in scan order, after the
 * skipped cells are removed. `length` is `Infinity` for an unbounded result.
 */
type Vector = {
  type: "vector";
  length: number;
  /** Reads cell `index`, which must be below `length`. */
  cellAt: (index: number, context: EvaluationContext) => SingleEvaluationResult;
  /**
   * The cells numbered `first` to `last` (inclusive, may be `Infinity`) in
   * order. Blank cells may be missing when the array is read sparsely.
   */
  cellsIn: (
    first: number,
    last: number,
    context: EvaluationContext
  ) => VectorCells | ErrorEvaluationResult;
};

/** Numbering of an array's cells in scan order. */
type ScanOrder = {
  index: (x: number, y: number) => number;
  position: (index: number) => { x: number; y: number };
};

/**
 * Numbers the cells of a `rows` x `cols` array in scan order. Undefined when
 * the scan would never finish its first row (or column), because that line is
 * unbounded and another one follows it.
 */
function getScanOrder(
  rows: number,
  cols: number,
  byColumn: boolean
): ScanOrder | undefined {
  if (rows === 1) {
    return { index: (x) => x, position: (index) => ({ x: index, y: 0 }) };
  }
  if (cols === 1) {
    return { index: (_x, y) => y, position: (index) => ({ x: 0, y: index }) };
  }
  if (!byColumn && Number.isFinite(cols)) {
    return {
      index: (x, y) => y * cols + x,
      position: (index) => ({ x: index % cols, y: Math.floor(index / cols) }),
    };
  }
  if (byColumn && Number.isFinite(rows)) {
    return {
      index: (x, y) => x * rows + y,
      position: (index) => ({ x: Math.floor(index / rows), y: index % rows }),
    };
  }
  return undefined;
}

/**
 * Reads cells `first` to `last` one by one, turning thrown evaluation errors
 * into error cells. Keeps reading past cells that are awaiting evaluation so
 * every dependency is registered, then returns the first of them.
 */
function readIndexRange(
  first: number,
  last: number,
  name: string,
  cellAt: (index: number) => SingleEvaluationResult,
  context: EvaluationContext
): VectorCells | AwaitingEvaluationResult {
  if (!Number.isFinite(last)) {
    throwInfiniteCells(name);
  }
  const cells: IndexedCell[] = [];
  let awaiting: AwaitingEvaluationResult | undefined;
  for (let index = first; index <= last; index++) {
    const result = captureEvaluationErrors(context.dependencyNode, () =>
      cellAt(index)
    );
    if (result.type === "awaiting-evaluation") {
      awaiting ??= result;
      continue;
    }
    cells.push({ index, result });
  }
  return awaiting ?? { type: "cells", cells };
}

/**
 * Keeps every cell: the result maps its cells to cells of the array, without
 * reading the array up front.
 */
function createMappedVector(
  array: ArrayArgument,
  order: ScanOrder,
  name: string
): Vector {
  const length = array.rows * array.cols;
  const cellAt = (index: number, context: EvaluationContext) => {
    const { x, y } = order.position(index);
    return array.cellAt(x, y, context);
  };
  return {
    type: "vector",
    length,
    cellAt,
    cellsIn: (first, last, context) => {
      if (!array.sparse) {
        return readIndexRange(
          first,
          Math.min(last, length - 1),
          name,
          (index) => cellAt(index, context),
          context
        );
      }
      const read = array.readCells(context);
      if (read.type !== "cells") {
        return read;
      }
      const cells: IndexedCell[] = [];
      for (const cell of read.cells) {
        const index = order.index(cell.x, cell.y);
        if (index >= first && index <= last) {
          cells.push({ index, result: cell.result });
        }
      }
      cells.sort((a, b) => a.index - b.index);
      return { type: "cells", cells };
    },
  };
}

/**
 * Number of entries of the ascending array `values` that are below `limit`.
 */
function countBelow(values: readonly number[], limit: number): number {
  let low = 0;
  let high = values.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (values[middle]! < limit) {
      low = middle + 1;
    } else {
      high = middle;
    }
  }
  return low;
}

/**
 * Skips errors but keeps blanks. The array is read once; the result numbers the
 * remaining cells by skipping the scan positions of the errors, which works for
 * unbounded arrays too.
 */
function createErrorSkippingVector(
  array: ArrayArgument,
  order: ScanOrder,
  context: EvaluationContext
): Vector | ErrorEvaluationResult {
  const read = array.readCells(context);
  if (read.type !== "cells") {
    return read;
  }

  const errorIndexSet = new Set<number>();
  const values = new Map<number, ArrayCellResult>();
  for (const cell of read.cells) {
    const index = order.index(cell.x, cell.y);
    if (cell.result.type === "error") {
      errorIndexSet.add(index);
    } else {
      values.set(index, cell.result);
    }
  }
  const errorIndexes = Array.from(errorIndexSet).sort((a, b) => a - b);

  // `errorIndexes[i] - i` counts the kept cells before the i-th error, so the
  // kept cell number `index` comes after every error whose count is <= index.
  const keptBefore = errorIndexes.map((errorIndex, i) => errorIndex - i);
  const sourceIndex = (index: number) =>
    index + countBelow(keptBefore, index + 1);

  return {
    type: "vector",
    length: array.rows * array.cols - errorIndexes.length,
    cellAt: (index) => values.get(sourceIndex(index)) ?? blankResult(),
    cellsIn: (first, last) => {
      const cells: IndexedCell[] = [];
      for (const [source, result] of values) {
        const index = source - countBelow(errorIndexes, source);
        if (index >= first && index <= last) {
          cells.push({ index, result });
        }
      }
      cells.sort((a, b) => a.index - b.index);
      return { type: "cells", cells };
    },
  };
}

/**
 * Skips blanks, and errors when `skipErrors` is set. The array is read once and
 * the result serves the surviving cells, which are always finitely many.
 */
function createBlankSkippingVector(
  array: ArrayArgument,
  byColumn: boolean,
  skipErrors: boolean,
  context: EvaluationContext
): Vector | ErrorEvaluationResult {
  const read = array.readCells(context);
  if (read.type !== "cells") {
    return read;
  }

  const kept = read.cells.filter(
    (cell) =>
      !isBlankResult(cell.result) &&
      !(skipErrors && cell.result.type === "error")
  );
  kept.sort(
    byColumn
      ? (a: ArrayCell, b: ArrayCell) => a.x - b.x || a.y - b.y
      : (a: ArrayCell, b: ArrayCell) => a.y - b.y || a.x - b.x
  );

  return {
    type: "vector",
    length: kept.length,
    cellAt: (index) => kept[index]!.result,
    cellsIn: (first, last) => {
      const end = Math.min(last, kept.length - 1);
      const cells: IndexedCell[] = [];
      for (let index = first; index <= end; index++) {
        cells.push({ index, result: kept[index]!.result });
      }
      return { type: "cells", cells };
    },
  };
}

function createVectorResult(
  vector: Vector,
  orientation: VectorOrientation,
  name: string,
  context: EvaluationContext
): SpilledValuesEvaluationResult {
  const rows = orientation === "column" ? vector.length : 1;
  const cols = orientation === "column" ? 1 : vector.length;
  const spillArea = (origin: CellAddress) =>
    createSpillArea(origin, rows, cols);

  return {
    type: "spilled-values",
    spillArea,
    source: `${name} function`,
    evaluate: (spillOffset, cellContext) => {
      const index = orientation === "column" ? spillOffset.y : spillOffset.x;
      const across = orientation === "column" ? spillOffset.x : spillOffset.y;
      if (across !== 0 || index < 0 || index >= vector.length) {
        return outsideResultError(name, context);
      }
      return vector.cellAt(index, cellContext);
    },
    evaluateAllCells: ({ intersection, context: readContext, origin }) => {
      const window = getSpillWindow(spillArea(origin), origin, intersection);
      if (!window) {
        return { type: "values", values: [] };
      }
      const cells = vector.cellsIn(
        orientation === "column" ? window.y0 : window.x0,
        orientation === "column" ? window.y1 : window.x1,
        readContext
      );
      if (cells.type !== "cells") {
        return cells;
      }
      return {
        type: "values",
        values: cells.cells.map(({ index, result }) => ({
          result,
          relativePos:
            orientation === "column" ? { x: 0, y: index } : { x: index, y: 0 },
        })),
      };
    },
  };
}

/**
 * Shared implementation of TOCOL and TOROW: flattens the array argument into
 * one column or one row.
 */
export function evaluateArrayToVector(
  this: FormulaEvaluator,
  node: FunctionNode,
  context: EvaluationContext,
  orientation: VectorOrientation
): FunctionEvaluationResult {
  const name = orientation === "column" ? "TOCOL" : "TOROW";
  const error = (
    err: FormulaError,
    message: string
  ): ErrorEvaluationResult => ({
    type: "error",
    err,
    message: `${name}: ${message}`,
    errAddress: context.dependencyNode,
  });

  if (node.args.length < 1 || node.args.length > 3) {
    return error(FormulaError.VALUE, "takes 1 to 3 arguments");
  }
  const arrayNode = getSuppliedArgument(node, 0);
  if (!arrayNode) {
    return error(FormulaError.VALUE, "array is required");
  }
  // Omitted options, as in TOCOL(A1:B2,,TRUE), keep their defaults.
  const ignoreNode = getSuppliedArgument(node, 1);
  const scanNode = getSuppliedArgument(node, 2);

  // Evaluate every argument before returning any of them, so that one pass
  // registers all dependencies.
  const arrayResult = this.evaluateNode(arrayNode, context);
  const ignoreResult = ignoreNode && this.evaluateNode(ignoreNode, context);
  const scanResult = scanNode && this.evaluateNode(scanNode, context);
  if (arrayResult.type === "awaiting-evaluation") {
    return arrayResult;
  }
  if (ignoreResult?.type === "awaiting-evaluation") {
    return ignoreResult;
  }
  if (scanResult?.type === "awaiting-evaluation") {
    return scanResult;
  }

  let ignore = 0;
  if (ignoreResult) {
    const value = readIntegerOption(ignoreResult, `${name} ignore`, context);
    if (typeof value !== "number") {
      return value;
    }
    ignore = value;
  }
  if (ignore !== 0 && ignore !== 1 && ignore !== 2 && ignore !== 3) {
    return error(
      FormulaError.VALUE,
      `ignore must be 0, 1, 2 or 3, got ${ignore}`
    );
  }

  let byColumn = false;
  if (scanResult) {
    const value = readLogicalOption(
      scanResult,
      `${name} scan_by_column`,
      context
    );
    if (typeof value !== "boolean") {
      return value;
    }
    byColumn = value;
  }

  const array = toArrayArgument(this, arrayResult, context);
  const skipBlanks = ignore === 1 || ignore === 3;
  const skipErrors = ignore === 2 || ignore === 3;

  let vector: Vector | ErrorEvaluationResult;
  if (skipBlanks) {
    vector = createBlankSkippingVector(array, byColumn, skipErrors, context);
  } else {
    const order = getScanOrder(array.rows, array.cols, byColumn);
    if (!order) {
      return error(
        FormulaError.VALUE,
        byColumn
          ? "cannot keep the blanks of an array with unbounded rows when scanning by column"
          : "cannot keep the blanks of an array with unbounded columns when scanning by row"
      );
    }
    vector = skipErrors
      ? createErrorSkippingVector(array, order, context)
      : createMappedVector(array, order, name);
  }
  if (vector.type !== "vector") {
    return vector;
  }

  if (vector.length === 0) {
    return error(FormulaError.CALC, "no cells are left in the result");
  }
  if (vector.length === 1) {
    return vector.cellAt(0, context);
  }
  return createVectorResult(vector, orientation, name, context);
}
