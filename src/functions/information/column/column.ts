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
  type CellInRangeResult,
} from "../../../core/types";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import type { ReferenceNode, RangeNode } from "../../../parser/ast";
import { EvaluationError } from "../../../evaluator/evaluation-error";
import {
  getRangeIntersection,
  isCellAddress,
  isRangeAddress,
} from "../../../core/utils";

/**
 * COLUMN function - Returns the column number of a reference
 *
 * Usage: COLUMN([reference])
 *
 * If reference is omitted, returns the column number of the cell in which the formula appears.
 *
 * Examples:
 * - COLUMN() returns the column number of the current cell
 * - COLUMN(C5) returns 3
 * - COLUMN(A1:E1) returns an array of column numbers {1,2,3,4,5}
 */
export const COLUMN: FunctionDefinition = {
  name: "COLUMN",
  evaluate: function (node, context): FunctionEvaluationResult {
    // If no arguments, return the column of the current cell (1-based)
    if (node.args.length === 0) {
      context.addContextDependency("col");
      return {
        type: "value",
        result: { type: "number", value: context.cellAddress.colIndex + 1 },
      };
    }

    if (node.args.length > 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "COLUMN function takes at most 1 argument",
        errAddress: context.dependencyNode,
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
        message: "COLUMN function requires a cell or range reference",
        errAddress: context.dependencyNode,
      };
    }

    // If we have a sourceCell (single cell reference), return its column
    if (isCellAddress(cellOrRange)) {
      return {
        type: "value",
        result: { type: "number", value: cellOrRange.colIndex + 1 },
      };
    }

    // If we have a sourceRange, return spilled column numbers
    if (isRangeAddress(cellOrRange) && argResult.type === "spilled-values") {
      const sourceRange = cellOrRange.range;

      // Return as spilled values - evaluate returns column numbers based on position
      return {
        type: "spilled-values",
        spillArea: (origin) => argResult.spillArea(origin),
        sourceRange: cellOrRange,
        source: "COLUMN with range/array argument",
        evaluate: (spillOffset, evalContext) => {
          // If we have a sourceRange, use it to get the actual column number
          if (sourceRange) {
            const actualCol = sourceRange.start.col + spillOffset.x;
            return {
              type: "value",
              result: { type: "number", value: actualCol + 1 }, // 1-based
            };
          }

          // Otherwise, use spillArea to get the column number relative to where it appears
          const spillArea = argResult.spillArea(evalContext.cellAddress);
          const actualCol = spillArea.start.col + spillOffset.x;
          return {
            type: "value",
            result: { type: "number", value: actualCol + 1 }, // 1-based
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
              return {
                type: "error",
                err: FormulaError.VALUE,
                message:
                  "Invalid intersection range for COLUMN function",
                errAddress: context.dependencyNode,
              };
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
              "COLUMN function cannot evaluate all cells over an infinite range"
            );
          }

          // Extract unique column numbers from the range
          const startCol = rangeToEvaluate.start.col;
          const endCol = rangeToEvaluate.end.col.value;

          const colSet = new Set<number>();
          for (let col = startCol; col <= endCol; col++) {
            colSet.add(col + 1); // 1-based
          }
          const cols = Array.from(colSet).sort((a, b) => a - b);

          const results: CellInRangeResult[] = [];
          for (let i = 0; i < cols.length; i++) {
            const relativePos = { x: i, y: 0 };
            const evaled = evaluate(relativePos, context);
            results.push({
              result: evaled!,
              relativePos,
            });
          }
          return { type: "values", values: results };
        },
      };
    }

    // If we reach here, cellOrRange type is unknown
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "COLUMN function requires a cell or range reference",
      errAddress: context.dependencyNode,
    };
  },
};
