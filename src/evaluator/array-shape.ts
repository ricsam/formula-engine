import type {
  CellAddress,
  EvaluateAllCellsResult,
  SpilledValuesEvaluationResult,
  SpreadsheetRange,
  ValueEvaluationResult,
} from "../core/types";

/**
 * The size of a value viewed as an array. Unbounded dimensions, such as the
 * rows of `A:A`, are `Infinity`.
 */
export type ArrayShape = {
  rows: number;
  cols: number;
};

/** Returns the number of rows and columns covered by a spill area. */
export function getSpillAreaShape(area: SpreadsheetRange): ArrayShape {
  return {
    rows:
      area.end.row.type === "number"
        ? area.end.row.value - area.start.row + 1
        : Infinity,
    cols:
      area.end.col.type === "number"
        ? area.end.col.value - area.start.col + 1
        : Infinity,
  };
}

/**
 * Returns the shape of an evaluated value. A single value is a 1x1 array; a
 * spilled value has the shape of the area it occupies when spilled from
 * `origin`, normally the cell being evaluated (`context.cellAddress`).
 *
 * Error and awaiting-evaluation results have no shape: handle them first.
 */
export function getResultShape(
  result: ValueEvaluationResult | SpilledValuesEvaluationResult,
  origin: CellAddress
): ArrayShape {
  if (result.type === "value") {
    return { rows: 1, cols: 1 };
  }
  return getSpillAreaShape(result.spillArea(origin));
}

/**
 * Moves the `relativePos` of `evaluateAllCells` results by `offset`. Use it to
 * re-base cells read from part of an array, such as a sub-range or a spill
 * anchored elsewhere, onto the start of the array they are reported for.
 * Errors pass through, and a zero offset returns the same result object.
 */
export function offsetCellResults(
  result: EvaluateAllCellsResult,
  offset: { x: number; y: number }
): EvaluateAllCellsResult {
  if (result.type !== "values" || (offset.x === 0 && offset.y === 0)) {
    return result;
  }
  return {
    type: "values",
    values: result.values.map((cell) => ({
      result: cell.result,
      relativePos: {
        x: cell.relativePos.x + offset.x,
        y: cell.relativePos.y + offset.y,
      },
    })),
  };
}
