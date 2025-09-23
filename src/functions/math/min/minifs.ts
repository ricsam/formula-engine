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
 * MINIFS function - Returns the minimum value among cells specified by multiple criteria
 *
 * Usage: MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
 *
 * min_range: The range of cells from which you want the minimum
 * criteria_range1: The first range in which to evaluate criteria
 * criteria1: The criteria to apply to criteria_range1
 * criteria_range2, criteria2: Optional additional criteria pairs
 *
 * Examples:
 *   MINIFS(B1:B10, A1:A10, "Apple") - min of B1:B10 where A1:A10 = "Apple"
 *   MINIFS(C1:C10, A1:A10, "Apple", B1:B10, ">10") - min of C1:C10 where A1:A10 = "Apple" AND B1:B10 > 10
 *
 * Note:
 * - All criteria must be satisfied for a value to be included
 * - Only numeric values are considered for the minimum
 * - Returns error if no values match all criteria
 */
export const MINIFS: FunctionDefinition = {
  name: "MINIFS",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Must have at least 3 arguments (min_range, criteria_range1, criteria1)
    // Additional arguments must come in pairs (criteria_range, criteria)
    if (node.args.length < 3 || (node.args.length - 1) % 2 !== 0) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message:
          "MINIFS function requires an odd number of arguments (min 3): min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...",
      };
    }

    // Evaluate min range
    const minRangeResult = this.evaluateNode(node.args[0]!, context);
    if (minRangeResult.type === "error") {
      return minRangeResult;
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
          message: "MINIFS criteria must be single values",
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

    let minValue = Infinity;
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

    if (hasEmptyCriteria && minRangeResult.type === "spilled-values") {
      // Need special handling for empty cell criteria
      const minSpillArea = minRangeResult.spillArea(context.currentCell);

      // Check if this is an infinite range
      if (
        minSpillArea.end.col.type === "infinity" ||
        minSpillArea.end.row.type === "infinity"
      ) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Cannot process infinite ranges with empty cell criteria",
        };
      }

      // Finite range - iterate over all cells in spill areas
      if (
        minSpillArea.end.col.type === "number" &&
        minSpillArea.end.row.type === "number"
      ) {
        for (
          let row = minSpillArea.start.row;
          row <= minSpillArea.end.row.value;
          row++
        ) {
          for (
            let col = minSpillArea.start.col;
            col <= minSpillArea.end.col.value;
            col++
          ) {
            const minCell = minRangeResult.evaluate(
              {
                x: col - minSpillArea.start.col,
                y: row - minSpillArea.start.row,
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
                  row - minSpillArea.start.row + criteriaSpillArea.start.row;
                const criteriaCol =
                  col - minSpillArea.start.col + criteriaSpillArea.start.col;

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

            if (allMatch && minCell?.type === "value") {
              if (minCell.result.type === "number") {
                minValue = Math.min(minValue, minCell.result.value);
                hasValues = true;
              } else if (
                minCell.result.type === "infinity" &&
                minCell.result.sign === "negative"
              ) {
                return {
                  type: "value",
                  result: minCell.result,
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
          result: { type: "number", value: minValue },
        };
      }
    }

    // Handle different min range types
    if (minRangeResult.type === "value") {
      // Single value case - check if all criteria match
      let allMatch = true;
      for (const { rangeResult, parsedCriteria } of criteriaPairs) {
        if (rangeResult.type === "value") {
          if (!matchesParsedCriteria(rangeResult.result, parsedCriteria)) {
            allMatch = false;
            break;
          }
        } else {
          // Single min value but range criteria - doesn't make sense
          allMatch = false;
          break;
        }
      }

      if (allMatch && minRangeResult.result.type === "number") {
        minValue = minRangeResult.result.value;
        hasValues = true;
      } else if (allMatch && minRangeResult.result.type === "infinity") {
        return {
          type: "value",
          result: minRangeResult.result,
        };
      }
    } else if (minRangeResult.type === "spilled-values") {
      // Range case - iterate over all values and check criteria
      const minValues = Array.from(
        minRangeResult.evaluateAllCells.call(this, {
          context,
          evaluate: minRangeResult.evaluate,
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
            // Single criteria value - repeat for all min values
            return Array(minValues.length).fill(rangeResult);
          } else {
            return [];
          }
        });

      // Check that all arrays have compatible lengths
      const minLength = Math.min(
        minValues.length,
        ...criteriaValueArrays.map((arr) => arr.length)
      );

      for (let i = 0; i < minLength; i++) {
        const minCell = minValues[i];

        if (minCell?.result.type === "error") {
          continue; // Skip error cells
        }

        if (minCell?.result.type === "value") {
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
            if (minCell.result.result.type === "number") {
              minValue = Math.min(minValue, minCell.result.result.value);
              hasValues = true;
            } else if (
              minCell.result.result.type === "infinity" &&
              minCell.result.result.sign === "negative"
            ) {
              // Negative infinity is always the minimum
              return {
                type: "value",
                result: minCell.result.result,
              };
            }
          }
        }
      }
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid min range argument type",
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
      result: { type: "number", value: minValue },
    };
  },
};
