import {
  FormulaError,
  type CellAddress,
  type CellInRangeResult,
  type CellValue,
  type ErrorEvaluationResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpreadsheetRange,
} from "../../../core/types";
import { getRangeIntersection, isRangeOneCell } from "../../../core/utils";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import { EvaluationError } from "../../../evaluator/evaluation-error";

/**
 * UNIQUE(array, [by_col], [exactly_once])
 * Returns the distinct rows (or columns) of an array.
 *
 * @param array - The range or array to deduplicate
 * @param by_col - [Optional] FALSE (default) compares rows, TRUE compares columns
 * @param exactly_once - [Optional] TRUE keeps only rows/columns that occur exactly once
 *
 * Comparison follows spreadsheet semantics: values must share a type to be equal,
 * and text is compared case-insensitively. Empty cells inside `array` compare as
 * empty text, which is how the engine reads any unpopulated cell.
 *
 * `array` may be unbounded along the dimension that gets deduplicated, because the
 * unbounded tail of empty rows/columns collapses into a single entry. It may not be
 * unbounded along the other dimension, since that would produce an unbounded result.
 */

const EMPTY_VALUE: CellValue = { type: "string", value: "" };

/** Separator that cannot appear in a value key, keeping vector keys unambiguous. */
const KEY_SEPARATOR = "\u0000";

function valueKey(value: CellValue): string {
  switch (value.type) {
    case "string":
      // Spreadsheet text comparison is case-insensitive.
      return `s${value.value.toLowerCase()}`;
    case "number":
      return `n${value.value}`;
    case "boolean":
      return `b${value.value}`;
    case "infinity":
      return `i${value.sign}`;
  }
}

function vectorKey(values: CellValue[]): string {
  return values.map(valueKey).join(KEY_SEPARATOR);
}

/**
 * Reads an optional logical argument. Booleans and numbers are accepted, matching
 * how spreadsheets treat 0/1 as FALSE/TRUE.
 */
function readLogicalArgument(
  result: FunctionEvaluationResult,
  argumentName: string,
  context: EvaluationContext
): boolean | ErrorEvaluationResult {
  if (result.type === "error" || result.type === "awaiting-evaluation") {
    return result;
  }

  if (result.type === "spilled-values") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `UNIQUE ${argumentName} must be a single logical value`,
      errAddress: context.dependencyNode,
    };
  }

  if (result.result.type === "boolean") {
    return result.result.value;
  }

  if (result.result.type === "number") {
    return result.result.value !== 0;
  }

  return {
    type: "error",
    err: FormulaError.VALUE,
    message: `UNIQUE ${argumentName} must be a logical value, got ${result.result.type}`,
    errAddress: context.dependencyNode,
  };
}

type DistinctEntry = {
  /** Index of the first occurrence along the deduplicated dimension. */
  sourceIndex: number;
  /** Occurrence count, capped at 2 because only "exactly once" matters. */
  occurrences: number;
};

export const UNIQUE: FunctionDefinition = {
  name: "UNIQUE",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 1 || node.args.length > 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "UNIQUE function takes 1 to 3 arguments",
        errAddress: context.dependencyNode,
      };
    }

    const arrayResult = this.evaluateNode(node.args[0]!, context);
    if (
      arrayResult.type === "error" ||
      arrayResult.type === "awaiting-evaluation"
    ) {
      return arrayResult;
    }

    let byColumn = false;
    if (node.args[1]) {
      const byColumnArg = readLogicalArgument(
        this.evaluateNode(node.args[1], context),
        "by_col",
        context
      );
      if (typeof byColumnArg !== "boolean") {
        return byColumnArg;
      }
      byColumn = byColumnArg;
    }

    let exactlyOnce = false;
    if (node.args[2]) {
      const exactlyOnceArg = readLogicalArgument(
        this.evaluateNode(node.args[2], context),
        "exactly_once",
        context
      );
      if (typeof exactlyOnceArg !== "boolean") {
        return exactlyOnceArg;
      }
      exactlyOnce = exactlyOnceArg;
    }

    // A single value is already distinct and occurs exactly once.
    if (arrayResult.type === "value") {
      return { type: "value", result: arrayResult.result };
    }

    const inputArea = arrayResult.spillArea(context.cellAddress);
    const rowsUnbounded = inputArea.end.row.type === "infinity";
    const columnsUnbounded = inputArea.end.col.type === "infinity";

    // The dimension that survives deduplication becomes the result width/height,
    // so it has to be bounded.
    if (byColumn ? rowsUnbounded : columnsUnbounded) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: byColumn
          ? "UNIQUE cannot compare columns of an array with an unbounded number of rows"
          : "UNIQUE cannot compare rows of an array with an unbounded number of columns",
        errAddress: context.dependencyNode,
      };
    }

    const inputCells = arrayResult.evaluateAllCells.call(this, {
      context,
      evaluate: arrayResult.evaluate,
      origin: context.cellAddress,
      lookupOrder: byColumn ? "row-major" : "col-major",
    });
    if (inputCells.type !== "values") {
      return inputCells;
    }

    // Positions are relative to the start of the input area.
    const values = new Map<string, CellValue>();
    let maxX = -1;
    let maxY = -1;
    for (const cell of inputCells.values) {
      if (
        cell.result.type === "error" ||
        cell.result.type === "awaiting-evaluation"
      ) {
        return cell.result;
      }
      const { x, y } = cell.relativePos;
      values.set(`${x},${y}`, cell.result.result);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
    }

    const valueAt = (x: number, y: number): CellValue =>
      values.get(`${x},${y}`) ?? EMPTY_VALUE;

    // Unbounded dimensions are only populated up to the last non-empty cell. One
    // extra index past that stands in for the unbounded tail of empty cells.
    const rowCount =
      inputArea.end.row.type === "number"
        ? inputArea.end.row.value - inputArea.start.row + 1
        : maxY + 1;
    const columnCount =
      inputArea.end.col.type === "number"
        ? inputArea.end.col.value - inputArea.start.col + 1
        : maxX + 1;

    const comparedCount = byColumn ? columnCount : rowCount;
    const keptCount = byColumn ? rowCount : columnCount;
    const hasUnboundedTail = byColumn ? columnsUnbounded : rowsUnbounded;

    const vectorAt = (index: number): CellValue[] => {
      const vector: CellValue[] = [];
      for (let offset = 0; offset < keptCount; offset++) {
        vector.push(
          byColumn ? valueAt(index, offset) : valueAt(offset, index)
        );
      }
      return vector;
    };

    const entriesByKey = new Map<string, DistinctEntry>();
    const distinctEntries: DistinctEntry[] = [];

    const addOccurrences = (sourceIndex: number, occurrences: number) => {
      const key = vectorKey(vectorAt(sourceIndex));
      const existing = entriesByKey.get(key);
      if (existing) {
        existing.occurrences = Math.min(2, existing.occurrences + occurrences);
        return;
      }
      const entry: DistinctEntry = { sourceIndex, occurrences };
      entriesByKey.set(key, entry);
      distinctEntries.push(entry);
    };

    for (let index = 0; index < comparedCount; index++) {
      addOccurrences(index, 1);
    }

    if (hasUnboundedTail) {
      // `comparedCount` is the first index past every populated cell, so reading it
      // yields the all-empty vector. The tail repeats it endlessly, which is never
      // "exactly once".
      addOccurrences(comparedCount, 2);
    }

    const resultEntries = exactlyOnce
      ? distinctEntries.filter((entry) => entry.occurrences === 1)
      : distinctEntries;

    if (resultEntries.length === 0) {
      return {
        type: "error",
        err: FormulaError.CALC,
        message: "UNIQUE returned an empty array",
        errAddress: context.dependencyNode,
      };
    }

    const resultRows = byColumn ? keptCount : resultEntries.length;
    const resultColumns = byColumn ? resultEntries.length : keptCount;

    const resultAt = (x: number, y: number): SingleEvaluationResult => {
      const entry = byColumn ? resultEntries[x] : resultEntries[y];
      if (!entry || x < 0 || y < 0 || x >= resultColumns || y >= resultRows) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "UNIQUE: requested cell is outside the result",
          errAddress: context.dependencyNode,
        };
      }
      return {
        type: "value",
        result: byColumn
          ? valueAt(entry.sourceIndex, y)
          : valueAt(x, entry.sourceIndex),
      };
    };

    const spillArea = (origin: CellAddress): SpreadsheetRange => ({
      start: {
        col: origin.colIndex,
        row: origin.rowIndex,
      },
      end: {
        col: { type: "number", value: origin.colIndex + resultColumns - 1 },
        row: { type: "number", value: origin.rowIndex + resultRows - 1 },
      },
    });

    if (isRangeOneCell(spillArea(context.cellAddress))) {
      return resultAt(0, 0);
    }

    return {
      type: "spilled-values",
      spillArea,
      source: "UNIQUE function",
      evaluate: (spillOffset) => resultAt(spillOffset.x, spillOffset.y),
      evaluateAllCells: ({ evaluate, intersection, context, origin }) => {
        let range = spillArea(origin);
        if (intersection) {
          const intersected = getRangeIntersection(range, intersection);
          if (!intersected) {
            throw new EvaluationError(
              FormulaError.REF,
              "UNIQUE: intersection does not overlap the result"
            );
          }
          range = intersected;
        }
        if (
          range.end.row.type === "infinity" ||
          range.end.col.type === "infinity"
        ) {
          throw new EvaluationError(
            FormulaError.REF,
            "UNIQUE: can not evaluate all cells over an infinite range"
          );
        }

        const results: CellInRangeResult[] = [];
        for (let row = range.start.row; row <= range.end.row.value; row++) {
          for (let col = range.start.col; col <= range.end.col.value; col++) {
            const relativePos = {
              x: col - origin.colIndex,
              y: row - origin.rowIndex,
            };
            results.push({
              result: evaluate(relativePos, context),
              relativePos,
            });
          }
        }

        return { type: "values", values: results };
      },
    };
  },
};
