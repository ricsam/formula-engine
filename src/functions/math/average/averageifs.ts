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
import { 
  hasEmptyCriteria, 
  checkAllCriteriaMatch,
  processValuesWithCriteria,
  type CriteriaPair
} from "../../multi-criteria-utils";
import { parseCriteria, matchesParsedCriteria } from "../../criteria-parser";

/**
 * AVERAGEIFS function - Calculates the average of cells in a range that meet multiple criteria
 * 
 * Usage: AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
 * 
 * average_range: The range of cells to average
 * criteria_range1: The first range to evaluate against criteria1
 * criteria1: The first criteria to match against
 * criteria_range2, criteria2: Optional additional criteria pairs
 * 
 * Examples:
 *   AVERAGEIFS(B1:B10, A1:A10, "Apple") - averages B1:B10 where A1:A10 = "Apple"
 *   AVERAGEIFS(C1:C10, A1:A10, "Apple", B1:B10, ">10") - averages C1:C10 where A1:A10 = "Apple" AND B1:B10 > 10
 * 
 * Note:
 * - All criteria must be satisfied for a value to be included
 * - Only numeric values are included in the average calculation
 * - Returns error if no values match all criteria
 */
export const AVERAGEIFS: FunctionDefinition = {
  name: "AVERAGEIFS",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Must have at least 3 arguments (average_range, criteria_range1, criteria1)
    // Additional arguments must come in pairs (criteria_range, criteria)
    if (node.args.length < 3 || (node.args.length - 1) % 2 !== 0) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message:
          "AVERAGEIFS function requires an odd number of arguments (min 3): average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...",
      };
    }

    // Evaluate average range
    const averageRangeResult = this.evaluateNode(node.args[0]!, context);
    if (averageRangeResult.type === "error") {
      return averageRangeResult;
    }

    // Parse criteria pairs manually (avoiding the utility for now to debug)
    const criteriaPairs: CriteriaPair[] = [];

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
          message: "Criteria must be single values",
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

    // Special case: handle empty/non-empty cell criteria with spilled values
    if (hasEmptyCriteria(criteriaPairs) && averageRangeResult.type === "spilled-values") {
      const averageSpillArea = averageRangeResult.spillArea(context.currentCell);

      // Check if this is an infinite range
      if (
        averageSpillArea.end.col.type === "infinity" ||
        averageSpillArea.end.row.type === "infinity"
      ) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Cannot process infinite ranges with empty cell criteria",
        };
      }

      // Finite range - iterate over all cells in spill areas
      if (
        averageSpillArea.end.col.type === "number" &&
        averageSpillArea.end.row.type === "number"
      ) {
        let sum = 0;
        let count = 0;

        for (
          let row = averageSpillArea.start.row;
          row <= averageSpillArea.end.row.value;
          row++
        ) {
          for (
            let col = averageSpillArea.start.col;
            col <= averageSpillArea.end.col.value;
            col++
          ) {
            const averageCell = averageRangeResult.evaluate(
              {
                x: col - averageSpillArea.start.col,
                y: row - averageSpillArea.start.row,
              },
              context
            );

            // Check all criteria for this position
            if (checkAllCriteriaMatch(criteriaPairs, averageSpillArea, row, col, context)) {
              if (averageCell?.type === "value") {
                if (averageCell.result.type === "number") {
                  sum += averageCell.result.value;
                  count++;
                } else if (averageCell.result.type === "infinity") {
                  return {
                    type: "value",
                    result: averageCell.result,
                  };
                }
              }
            }
          }
        }

        if (count === 0) {
          return {
            type: "error",
            err: FormulaError.DIV0,
            message: "No numeric values match all criteria",
          };
        }

        return {
          type: "value",
          result: { type: "number", value: sum / count },
        };
      }
    }

    // Process values with criteria - simplified implementation
    let sum = 0;
    let count = 0;

    // Handle different average range types
    if (averageRangeResult.type === "value") {
      // Single value case - check if all criteria match
      let allMatch = true;
      for (const { rangeResult, parsedCriteria } of criteriaPairs) {
        if (rangeResult.type === "value") {
          if (!matchesParsedCriteria(rangeResult.result, parsedCriteria)) {
            allMatch = false;
            break;
          }
        } else {
          // Single value but range criteria - doesn't make sense
          allMatch = false;
          break;
        }
      }

      if (allMatch && averageRangeResult.result.type === "number") {
        sum = averageRangeResult.result.value;
        count = 1;
      } else if (allMatch && averageRangeResult.result.type === "infinity") {
        return {
          type: "value",
          result: averageRangeResult.result,
        };
      }
    } else if (averageRangeResult.type === "spilled-values") {
      // Range case - iterate over all values and check criteria
      const averageValues = Array.from(
        averageRangeResult.evaluateAllCells.call(this, {
          context,
          evaluate: averageRangeResult.evaluate,
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
            // Single criteria value - repeat for all average values
            return Array(averageValues.length).fill(rangeResult);
          } else {
            return [];
          }
        });

      // Check that all arrays have compatible lengths
      const minLength = Math.min(
        averageValues.length,
        ...criteriaValueArrays.map((arr) => arr.length)
      );

      for (let i = 0; i < minLength; i++) {
        const averageCell = averageValues[i];

        if (averageCell?.result.type === "error") {
          continue; // Skip error cells
        }

        if (averageCell?.result.type === "value") {
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
            if (averageCell.result.result.type === "number") {
              sum += averageCell.result.result.value;
              count++;
            } else if (averageCell.result.result.type === "infinity") {
              return {
                type: "value",
                result: averageCell.result.result,
              };
            }
          }
        }
      }
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid average range argument type",
      };
    }

    if (count === 0) {
      return {
        type: "error",
        err: FormulaError.DIV0,
        message: "No numeric values match all criteria",
      };
    }

    return {
      type: "value",
      result: { type: "number", value: sum / count },
    };
  },
};
