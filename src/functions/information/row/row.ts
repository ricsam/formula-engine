import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { ReferenceNode, RangeNode } from "src/parser/ast";
import { EvaluationError } from "src/evaluator/evaluation-error";
import { getRangeIntersection, isCellAddress, isRangeAddress } from "src/core/utils";

/**
 * ROW function - Returns the row number of a reference
 *
 * Usage: ROW([reference])
 *
 * If reference is omitted, returns the row number of the cell in which the formula appears.
 *
 * Examples:
 * - ROW() returns the row number of the current cell
 * - ROW(A5) returns 5
 * - ROW(A5:A10) returns an array of row numbers {5;6;7;8;9;10}
 */
export const ROW: FunctionDefinition = {
  name: "ROW",
  evaluate: function (node, context): FunctionEvaluationResult {
    // If no arguments, return the row of the current cell (1-based)
    if (node.args.length === 0) {
      return {
        type: "value",
        result: {
          type: "number",
          value: context.originCell.cellAddress.rowIndex + 1,
        },
      };
    }

    if (node.args.length > 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "ROW function takes at most 1 argument",
        errAddress: context.originCell.cellAddress,
      };
    }

    const argNode = node.args[0]!;

    // Evaluate the argument
    const argResult = this.evaluateNode(argNode, context);
    
    // Handle errors and awaiting evaluation
    if (
      argResult.type === "error" ||
      argResult.type === "awaiting-evaluation"
    ) {
      return argResult;
    }

    // Check if we have a cell or range reference
    const cellOrRange =
      argResult.type === "spilled-values"
        ? (argResult.sourceCell ?? argResult.sourceRange)
        : argResult.sourceCell;

    // If we don't have a sourceCell or sourceRange, the argument is invalid
    if (!cellOrRange) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "ROW function requires a cell or range reference",
        errAddress: context.originCell.cellAddress,
      };
    }

    // If we have a sourceCell (single cell reference), return its row
    if (isCellAddress(cellOrRange)) {
      return {
        type: "value",
        result: { type: "number", value: cellOrRange.rowIndex + 1 },
      };
    }

    // If we have a sourceRange, return spilled row numbers
    if (isRangeAddress(cellOrRange) && argResult.type === "spilled-values") {
      const sourceRange = cellOrRange.range;

      // Return as spilled values - evaluate returns row numbers based on position
      return {
        type: "spilled-values",
        spillArea: (origin) => argResult.spillArea(origin),
        sourceRange: cellOrRange,
        source: "ROW with range/array argument",
        evaluate: (spillOffset, evalContext) => {
          // If we have a sourceRange, use it to get the actual row number
          if (sourceRange) {
            const actualRow = sourceRange.start.row + spillOffset.y;
            return {
              type: "value",
              result: { type: "number", value: actualRow + 1 }, // 1-based
            };
          }

          // Otherwise, use spillArea to get the row number relative to where it appears
          const spillArea = argResult.spillArea(
            evalContext.originCell.cellAddress
          );
          const actualRow = spillArea.start.row + spillOffset.y;
          return {
            type: "value",
            result: { type: "number", value: actualRow + 1 }, // 1-based
          };
        },
        evaluateAllCells: ({ evaluate, intersection, context, origin }) => {
          // Determine the actual range to evaluate
          let rangeToEvaluate = cellOrRange.range;

          // Apply intersection if provided
          if (intersection) {
            // Use the intersection range
            const newRange = getRangeIntersection(
              rangeToEvaluate,
              intersection
            );
            if (!newRange) {
              return [];
            }
            rangeToEvaluate = newRange;
          }

          // Check if the range is infinite
          if (
            rangeToEvaluate.end.row.type === "infinity" ||
            rangeToEvaluate.end.col.type === "infinity"
          ) {
            throw new EvaluationError(
              FormulaError.VALUE,
              "ROW function cannot evaluate all cells over an infinite range"
            );
          }

          // Extract unique row numbers from the range
          const startRow = rangeToEvaluate.start.row;
          const endRow = rangeToEvaluate.end.row.value;

          const rowSet = new Set<number>();
          for (let row = startRow; row <= endRow; row++) {
            rowSet.add(row + 1); // 1-based
          }
          const rows = Array.from(rowSet).sort((a, b) => a - b);

          const results = [];
          for (let i = 0; i < rows.length; i++) {
            const relativePos = { x: 0, y: i };
            const evaled = evaluate(relativePos, context);
            results.push({
              result: evaled!,
              relativePos,
            });
          }
          return results;
        },
      };
    }

    // If we reach here, cellOrRange type is unknown
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "ROW function requires a cell or range reference",
      errAddress: context.originCell.cellAddress,
    };
  },
};
