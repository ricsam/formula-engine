import {
  FormulaError,
  type CellValue,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type SpilledValuesEvaluationResult,
  type CellAddress,
  type SpreadsheetRange,
  type ErrorEvaluationResult,
} from "../../../core/types";
import { getRangeIntersection } from "../../../core/utils";
import { getSpillAreaShape } from "../../../evaluator/array-shape";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import { getSuppliedArgument, type FunctionNode } from "../../../parser/ast";

/**
 * INDEX function - Returns a value from a table or array
 * INDEX(array, row_num, [column_num])
 *
 * STRICT TYPE CHECKING:
 * - array: range/array only
 * - row_num: number only (integer, 1-based)
 * - column_num: number only (integer, 1-based, optional, default 1)
 *
 * A row_num or column_num of 0 selects the whole column or row, as in Excel:
 * `INDEX(A1:C3, 0, 2)` is B1:B3 and `INDEX(A1:C3, 2, 0)` is A2:C2, and both 0
 * select the whole array. A selection from a range is that range (it reads like
 * B1:B3); a selection from a computed array spills its cells.
 *
 * An omitted row_num or column_num (a comma with nothing) is 0, as in Excel:
 * `INDEX(A1:C3,,2)` is the second column.
 */

// Helper function to get value from array at specific position
function getValueFromArray(
  this: FormulaEvaluator,
  arrayResult: SpilledValuesEvaluationResult,
  row: number,
  col: number,
  context: EvaluationContext
): CellValue | ErrorEvaluationResult {
  const dims = arrayResult.spillArea(context.cellAddress);

  // Convert 1-based indices to 0-based
  const rowIndex = row - 1;
  const colIndex = col - 1;

  // Calculate actual cell position
  const actualRow = dims.start.row + rowIndex;
  const actualCol = dims.start.col + colIndex;

  // Check bounds
  if (
    actualRow < 0 ||
    (dims.end.row.type === "number" && actualRow > dims.end.row.value)
  ) {
    return {
      type: "error",
      err: FormulaError.REF,
      message: `INDEX: row_num ${row} is out of range`,
      errAddress: context.dependencyNode,
    };
  }

  if (
    actualCol < 0 ||
    (dims.end.col.type === "number" && actualCol > dims.end.col.value)
  ) {
    return {
      type: "error",
      err: FormulaError.REF,
      message: `INDEX: column_num ${col} is out of range`,
      errAddress: context.dependencyNode,
    };
  }

  const spilledAddress: CellAddress = {
    colIndex: actualCol,
    rowIndex: actualRow,
    sheetName: context.cellAddress.sheetName,
    workbookName: context.cellAddress.workbookName,
  };

  const spill = {
    address: spilledAddress,
    spillOffset: {
      x: actualCol - dims.start.col,
      y: actualRow - dims.start.row,
    },
  };

  const spillResult = arrayResult.evaluate(spill.spillOffset, context);

  if (!spillResult) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "INDEX: Unable to retrieve value from array",
      errAddress: context.dependencyNode,
    };
  }

  if (
    spillResult.type === "error" ||
    spillResult.type === "awaiting-evaluation"
  ) {
    return spillResult;
  }

  return spillResult.result;
}

/**
 * Reads row_num or column_num as a whole number. An omitted argument is 0.
 */
function readIndexArgument(
  this: FormulaEvaluator,
  node: FunctionNode,
  index: number,
  name: "row_num" | "column_num",
  context: EvaluationContext
): number | ErrorEvaluationResult {
  const argument = getSuppliedArgument(node, index);
  if (!argument) {
    return 0;
  }
  const result = this.evaluateNode(argument, context);
  if (result.type === "error" || result.type === "awaiting-evaluation") {
    return result;
  }
  if (result.type === "spilled-values") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message:
        "INDEX: Spilled row_num/column_num arguments not yet implemented",
      errAddress: context.dependencyNode,
    };
  }
  if (result.result.type !== "number") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `INDEX ${name} must be number, got ${result.result.type}`,
      errAddress: context.dependencyNode,
    };
  }
  const value = Math.floor(result.result.value);
  if (value < 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `INDEX ${name} must be >= 0, got ${value}`,
      errAddress: context.dependencyNode,
    };
  }
  return value;
}

/**
 * The cells of a computed array that INDEX selects with a row_num or
 * column_num of 0: `width` x `height` cells starting at column offset `x0` and
 * row offset `y0` of the array.
 */
function createArraySelection(
  arrayResult: SpilledValuesEvaluationResult,
  selection: { x0: number; y0: number; width: number; height: number },
  context: EvaluationContext
): SpilledValuesEvaluationResult {
  const { x0, y0, width, height } = selection;
  const spillArea = (origin: CellAddress): SpreadsheetRange => ({
    start: { col: origin.colIndex, row: origin.rowIndex },
    end: {
      col:
        width === Infinity
          ? { type: "infinity", sign: "positive" }
          : { type: "number", value: origin.colIndex + width - 1 },
      row:
        height === Infinity
          ? { type: "infinity", sign: "positive" }
          : { type: "number", value: origin.rowIndex + height - 1 },
    },
  });

  return {
    type: "spilled-values",
    source: "INDEX row or column",
    spillArea,
    evaluate: (offset, readerContext) => {
      if (
        offset.x < 0 ||
        offset.y < 0 ||
        offset.x >= width ||
        offset.y >= height
      ) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "INDEX: requested cell is outside the selection",
          errAddress: context.dependencyNode,
        };
      }
      return arrayResult.evaluate(
        { x: x0 + offset.x, y: y0 + offset.y },
        readerContext
      );
    },
    evaluateAllCells: function ({
      intersection,
      context: readerContext,
      origin,
      lookupOrder,
    }) {
      let block = spillArea(origin);
      if (intersection) {
        const overlap = getRangeIntersection(block, intersection);
        if (!overlap) {
          return { type: "values", values: [] };
        }
        block = overlap;
      }

      // The same cells of the array, which reports positions relative to the
      // origin too. Pass the array's own evaluate: forwarding the caller's
      // could apply a wrapper's mapping twice.
      const shift = (end: SpreadsheetRange["end"]["col"], by: number) =>
        end.type === "number" ? { type: "number" as const, value: end.value + by } : end;
      const cells = arrayResult.evaluateAllCells.call(this, {
        context: readerContext,
        origin,
        lookupOrder,
        evaluate: arrayResult.evaluate,
        intersection: {
          start: { col: block.start.col + x0, row: block.start.row + y0 },
          end: { col: shift(block.end.col, x0), row: shift(block.end.row, y0) },
        },
      });
      if (cells.type !== "values") {
        return cells;
      }

      const values = [];
      for (const cell of cells.values) {
        const x = cell.relativePos.x - x0;
        const y = cell.relativePos.y - y0;
        if (x >= 0 && y >= 0 && x < width && y < height) {
          values.push({ result: cell.result, relativePos: { x, y } });
        }
      }
      return { type: "values", values };
    },
  };
}

export const INDEX: FunctionDefinition = {
  name: "INDEX",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 2 || node.args.length > 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INDEX function takes 2 or 3 arguments",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate array argument
    const arrayResult = this.evaluateNode(node.args[0]!, context);
    if (
      arrayResult.type === "error" ||
      arrayResult.type === "awaiting-evaluation"
    ) {
      return arrayResult;
    }

    // Evaluate row_num argument (0 when omitted)
    const rowNum = readIndexArgument.call(this, node, 1, "row_num", context);
    if (typeof rowNum !== "number") {
      return rowNum;
    }

    // Evaluate column_num argument: 1 when missing, 0 when omitted
    const colNum =
      node.args.length > 2
        ? readIndexArgument.call(this, node, 2, "column_num", context)
        : 1;
    if (typeof colNum !== "number") {
      return colNum;
    }

    // A single value is a 1x1 array.
    if (arrayResult.type === "value") {
      if (rowNum > 1 || colNum > 1) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: `INDEX: Single value can only be accessed at [1,1], got [${rowNum},${colNum}]`,
          errAddress: context.dependencyNode,
        };
      }
      return {
        type: "value",
        result: arrayResult.result,
      } satisfies ValueEvaluationResult;
    }

    const shape = getSpillAreaShape(arrayResult.spillArea(context.cellAddress));
    if (rowNum > shape.rows) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: `INDEX: row_num ${rowNum} is out of range`,
        errAddress: context.dependencyNode,
      };
    }
    if (colNum > shape.cols) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: `INDEX: column_num ${colNum} is out of range`,
        errAddress: context.dependencyNode,
      };
    }

    const selection = {
      x0: colNum === 0 ? 0 : colNum - 1,
      y0: rowNum === 0 ? 0 : rowNum - 1,
      width: colNum === 0 ? shape.cols : 1,
      height: rowNum === 0 ? shape.rows : 1,
    };

    if (selection.width === 1 && selection.height === 1) {
      const result = getValueFromArray.call(
        this,
        arrayResult,
        selection.y0 + 1,
        selection.x0 + 1,
        context
      );

      if (result.type === "error" || result.type === "awaiting-evaluation") {
        return result;
      }

      return {
        type: "value",
        result,
      } satisfies ValueEvaluationResult;
    }

    // A whole row or column of a range is a range itself.
    if (arrayResult.sourceRange) {
      const { range, sheetName, workbookName } = arrayResult.sourceRange;
      const startCol = range.start.col + selection.x0;
      const startRow = range.start.row + selection.y0;
      return this.evaluateRange(
        {
          type: "range",
          range: {
            start: { col: startCol, row: startRow },
            end: {
              col:
                colNum === 0 ? range.end.col : { type: "number", value: startCol },
              row:
                rowNum === 0 ? range.end.row : { type: "number", value: startRow },
            },
          },
          sheetName,
          workbookName,
          isAbsolute: {
            start: { col: true, row: true },
            end: { col: true, row: true },
          },
        },
        context
      );
    }

    return createArraySelection(arrayResult, selection, context);
  },
};
