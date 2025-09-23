import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type CellInfinity,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import { parseCriteria, matchesParsedCriteria } from "../../criteria-parser";

/**
 * MAXIFS function - Returns the maximum value among cells specified by multiple criteria
 *
 * Usage: MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
 *
 * max_range: The range of cells from which you want the maximum
 * criteria_range1: The first range in which to evaluate criteria
 * criteria1: The criteria to apply to criteria_range1
 * criteria_range2, criteria2: Optional additional criteria pairs
 *
 * Examples:
 *   MAXIFS(B1:B10, A1:A10, "Apple") - max of B1:B10 where A1:A10 = "Apple"
 *   MAXIFS(C1:C10, A1:A10, "Apple", B1:B10, ">10") - max of C1:C10 where A1:A10 = "Apple" AND B1:B10 > 10
 *
 * Note:
 * - All criteria must be satisfied for a value to be included
 * - Only numeric values are considered for the maximum
 * - Returns error if no values match all criteria
 */
export const MAXIFS: FunctionDefinition = {
  name: "MAXIFS",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Must have at least 3 arguments (max_range, criteria_range1, criteria1)
    // Additional arguments must come in pairs (criteria_range, criteria)
    if (node.args.length < 3 || (node.args.length - 1) % 2 !== 0) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message:
          "MAXIFS function requires an odd number of arguments (min 3): max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...",
      };
    }

    // Evaluate max range
    const maxRangeResult = this.evaluateNode(node.args[0]!, context);
    if (maxRangeResult.type === "error") {
      return maxRangeResult;
    }

    // Parse criteria pairs
    const criteriaPairs: Array<{
      rangeResult: FunctionEvaluationResult;
      parsedCriteria: ReturnType<typeof parseCriteria>;
    }> = [];

    for (let i = 1; i < node.args.length; i += 2) {
      const criteriaRangeResult = this.evaluateNode(node.args[i]!, context);
      if (criteriaRangeResult.type === "error") {
        return criteriaRangeResult;
      }

      const criteriaResult = this.evaluateNode(node.args[i + 1]!, context);
      if (criteriaResult.type === "error") {
        return criteriaResult;
      }

      if (criteriaResult.type !== "value") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "MAXIFS criteria must be single values",
        };
      }

      const parsedCriteria = parseCriteria(criteriaResult.result);
      if (parsedCriteria.type === "error") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: parsedCriteria.message,
        };
      }

      criteriaPairs.push({
        rangeResult: criteriaRangeResult,
        parsedCriteria,
      });
    }

    let maxValue = -Infinity;
    let hasValues = false;

    // Special case: handle empty/non-empty cell criteria
    const hasEmptyCriteria = criteriaPairs.some(
      ({ parsedCriteria }) =>
        (parsedCriteria.type === "exact" &&
          parsedCriteria.value.type === "string" &&
          parsedCriteria.value.value === "") ||
        (parsedCriteria.type === "comparison" &&
          parsedCriteria.operator === "<>" &&
          parsedCriteria.value.type === "string" &&
          parsedCriteria.value.value === "")
    );

    if (hasEmptyCriteria && maxRangeResult.type === "spilled-values") {
      // Need special handling for empty cell criteria
      const maxSpillArea = maxRangeResult.spillArea(context.currentCell);

      // Check if this is an infinite range
      if (
        maxSpillArea.end.col.type === "infinity" ||
        maxSpillArea.end.row.type === "infinity"
      ) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Cannot process infinite ranges with empty cell criteria",
        };
      }

      // Finite range - iterate over all cells in spill areas
      if (
        maxSpillArea.end.col.type === "number" &&
        maxSpillArea.end.row.type === "number"
      ) {
        for (
          let row = maxSpillArea.start.row;
          row <= maxSpillArea.end.row.value;
          row++
        ) {
          for (
            let col = maxSpillArea.start.col;
            col <= maxSpillArea.end.col.value;
            col++
          ) {
            const maxCell = maxRangeResult.evaluate(
              {
                x: col - maxSpillArea.start.col,
                y: row - maxSpillArea.start.row,
              },
              context
            );

            // Check all criteria for this position
            let allMatch = true;
            for (let i = 0; i < criteriaPairs.length; i++) {
              const { rangeResult, parsedCriteria } = criteriaPairs[i]!;

              let criteriaCell;
              if (rangeResult.type === "spilled-values") {
                const criteriaSpillArea = rangeResult.spillArea(
                  context.currentCell
                );
                const criteriaRow =
                  row - maxSpillArea.start.row + criteriaSpillArea.start.row;
                const criteriaCol =
                  col - maxSpillArea.start.col + criteriaSpillArea.start.col;

                if (
                  criteriaRow <=
                    (criteriaSpillArea.end.row.type === "number"
                      ? criteriaSpillArea.end.row.value
                      : Infinity) &&
                  criteriaCol <=
                    (criteriaSpillArea.end.col.type === "number"
                      ? criteriaSpillArea.end.col.value
                      : Infinity)
                ) {
                  criteriaCell = rangeResult.evaluate(
                    {
                      x: criteriaCol - criteriaSpillArea.start.col,
                      y: criteriaRow - criteriaSpillArea.start.row,
                    },
                    context
                  );
                }
              } else if (rangeResult.type === "value") {
                criteriaCell = rangeResult;
              }

              // Check if criteria matches (including empty cell logic)
              if (
                parsedCriteria.type === "exact" &&
                parsedCriteria.value.type === "string" &&
                parsedCriteria.value.value === ""
              ) {
                // Match empty cells
                const isEmpty =
                  !criteriaCell ||
                  (criteriaCell.type === "value" &&
                    criteriaCell.result.type === "string" &&
                    criteriaCell.result.value === "");
                if (!isEmpty) {
                  allMatch = false;
                  break;
                }
              } else if (
                parsedCriteria.type === "comparison" &&
                parsedCriteria.operator === "<>" &&
                parsedCriteria.value.type === "string" &&
                parsedCriteria.value.value === ""
              ) {
                // Match non-empty cells
                const isEmpty =
                  !criteriaCell ||
                  (criteriaCell.type === "value" &&
                    criteriaCell.result.type === "string" &&
                    criteriaCell.result.value === "");
                if (isEmpty) {
                  allMatch = false;
                  break;
                }
              } else {
                // Regular criteria matching
                if (
                  !criteriaCell ||
                  criteriaCell.type !== "value" ||
                  !matchesParsedCriteria(criteriaCell.result, parsedCriteria)
                ) {
                  allMatch = false;
                  break;
                }
              }
            }

            if (allMatch && maxCell?.type === "value") {
              if (maxCell.result.type === "number") {
                maxValue = Math.max(maxValue, maxCell.result.value);
                hasValues = true;
              } else if (
                maxCell.result.type === "infinity" &&
                maxCell.result.sign === "positive"
              ) {
                return {
                  type: "value",
                  result: maxCell.result,
                };
              }
            }
          }
        }

        if (!hasValues) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "No numeric values match all criteria",
          };
        }

        return {
          type: "value",
          result: { type: "number", value: maxValue },
        };
      }
    }

    // Handle different max range types
    if (maxRangeResult.type === "value") {
      // Single value case - check if all criteria match
      let allMatch = true;
      for (const { rangeResult, parsedCriteria } of criteriaPairs) {
        if (rangeResult.type === "value") {
          if (!matchesParsedCriteria(rangeResult.result, parsedCriteria)) {
            allMatch = false;
            break;
          }
        } else {
          // Single max value but range criteria - doesn't make sense
          allMatch = false;
          break;
        }
      }

      if (allMatch && maxRangeResult.result.type === "number") {
        maxValue = maxRangeResult.result.value;
        hasValues = true;
      } else if (allMatch && maxRangeResult.result.type === "infinity") {
        return {
          type: "value",
          result: maxRangeResult.result,
        };
      }
    } else if (maxRangeResult.type === "spilled-values") {
      // Range case - iterate over all values and check criteria
      const maxValues = Array.from(
        maxRangeResult.evaluateAllCells.call(this, {
          context,
          evaluate: maxRangeResult.evaluate,
          origin: context.currentCell,
        })
      );

      // Get all criteria values for parallel iteration
      const criteriaValueArrays: Array<Array<FunctionEvaluationResult>> =
        criteriaPairs.map(({ rangeResult }) => {
          if (rangeResult.type === "spilled-values") {
            return Array.from(
              rangeResult.evaluateAllCells.call(this, {
                context,
                evaluate: rangeResult.evaluate,
                origin: context.currentCell,
              })
            ).map((cell) => cell.result);
          } else if (rangeResult.type === "value") {
            // Single criteria value - repeat for all max values
            return Array(maxValues.length).fill(rangeResult);
          } else {
            return [];
          }
        });

      // Check that all arrays have compatible lengths
      const minLength = Math.min(
        maxValues.length,
        ...criteriaValueArrays.map((arr) => arr.length)
      );

      for (let i = 0; i < minLength; i++) {
        const maxCell = maxValues[i];

        if (maxCell?.result.type === "error") {
          continue; // Skip error cells
        }

        if (maxCell?.result.type === "value") {
          // Check if all criteria match for this position
          let allMatch = true;
          for (let j = 0; j < criteriaPairs.length; j++) {
            const criteriaCell = criteriaValueArrays[j]![i];
            if (criteriaCell?.type === "error") {
              allMatch = false;
              break;
            }
            if (criteriaCell?.type === "value") {
              if (
                !matchesParsedCriteria(
                  criteriaCell.result,
                  criteriaPairs[j]!.parsedCriteria
                )
              ) {
                allMatch = false;
                break;
              }
            } else {
              allMatch = false;
              break;
            }
          }

          if (allMatch) {
            if (maxCell.result.result.type === "number") {
              maxValue = Math.max(maxValue, maxCell.result.result.value);
              hasValues = true;
            } else if (
              maxCell.result.result.type === "infinity" &&
              maxCell.result.result.sign === "positive"
            ) {
              // Positive infinity is always the maximum
              return {
                type: "value",
                result: maxCell.result.result,
              };
            }
          }
        }
      }
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid max range argument type",
      };
    }

    if (!hasValues) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "No numeric values match all criteria",
      };
    }

    return {
      type: "value",
      result: { type: "number", value: maxValue },
    };
  },
};
