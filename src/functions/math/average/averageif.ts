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
 * AVERAGEIF function - Calculates the average of cells in a range that meet a criteria
 * 
 * Usage: AVERAGEIF(range, criteria, [average_range])
 * 
 * range: The range of cells to evaluate against the criteria
 * criteria: The criteria to match against. Can be:
 *   - Exact value: "Apple", 42
 *   - Comparison: ">10", "<=5", "<>0"
 *   - Wildcards: "App*", "?ruit"
 * average_range: Optional. The range to average. If omitted, uses the range parameter
 * 
 * Examples:
 *   AVERAGEIF(A1:A10, "Apple") - averages cells in A1:A10 that contain "Apple"
 *   AVERAGEIF(B1:B10, ">10") - averages cells in B1:B10 with values greater than 10
 *   AVERAGEIF(A1:A10, "Apple", B1:B10) - averages B1:B10 where A1:A10 contains "Apple"
 * 
 * Note:
 * - Supports type coercion for comparisons
 * - Case-sensitive string matching
 * - Wildcards: * matches any sequence, ? matches any single character
 * - Only numeric values are included in the average calculation
 */
export const AVERAGEIF: FunctionDefinition = {
  name: "AVERAGEIF",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 2 || node.args.length > 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "AVERAGEIF function takes 2 or 3 arguments",
      };
    }

    // Evaluate range argument (criteria range)
    const rangeResult = this.evaluateNode(node.args[0]!, context);
    if (rangeResult.type === "error") {
      return rangeResult;
    }

    // Evaluate criteria argument
    const criteriaResult = this.evaluateNode(node.args[1]!, context);
    if (criteriaResult.type === "error") {
      return criteriaResult;
    }

    // Criteria must be a single value
    if (criteriaResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "AVERAGEIF criteria must be a single value",
      };
    }

    // Evaluate average range (optional third argument)
    let averageRangeResult: FunctionEvaluationResult = rangeResult; // Default to same as criteria range
    if (node.args.length === 3) {
      averageRangeResult = this.evaluateNode(node.args[2]!, context);
      if (averageRangeResult.type === "error") {
        return averageRangeResult;
      }
    }

    // Parse the criteria using the criteria parser
    const parsedCriteria = parseCriteria(criteriaResult.result);
    if (parsedCriteria.type === "error") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: parsedCriteria.message,
      };
    }

    let sum = 0;
    let count = 0;

    // Special case: averaging values where criteria matches empty/non-empty cells
    if (rangeResult.type === "spilled-values" && 
        ((parsedCriteria.type === "exact" && 
          parsedCriteria.value.type === "string" && 
          parsedCriteria.value.value === "") ||
         (parsedCriteria.type === "comparison" && 
          parsedCriteria.operator === "<>" && 
          parsedCriteria.value.type === "string" && 
          parsedCriteria.value.value === ""))) {
      
      const spillArea = rangeResult.spillArea(context.currentCell);
      
      // Check if this is an infinite range
      if (spillArea.end.col.type === "infinity" || spillArea.end.row.type === "infinity") {
        // Can't average infinite empty cells
        return {
          type: "error",
          err: FormulaError.DIV0,
          message: "Cannot average infinite empty cells",
        };
      }
      
      // Finite range - check empty cells and average corresponding values
      if (spillArea.end.col.type === "number" && spillArea.end.row.type === "number" && 
          averageRangeResult.type === "spilled-values") {
        
        const averageSpillArea = averageRangeResult.spillArea(context.currentCell);
        
        for (let row = spillArea.start.row; row <= spillArea.end.row.value; row++) {
          for (let col = spillArea.start.col; col <= spillArea.end.col.value; col++) {
            const criteriaCell = rangeResult.evaluate(
              { x: col - spillArea.start.col, y: row - spillArea.start.row }, 
              context
            );
            
            // Check if criteria cell matches the condition (empty or non-empty)
            const isEmpty = !criteriaCell || 
                           (criteriaCell.type === "value" && 
                            criteriaCell.result.type === "string" && 
                            criteriaCell.result.value === "");
            
            const shouldInclude = (parsedCriteria.type === "exact" && isEmpty) || 
                                 (parsedCriteria.type === "comparison" && 
                                  parsedCriteria.operator === "<>" && !isEmpty);
            
            if (shouldInclude) {
              
              // Get corresponding cell from average range
              const avgRow = row - spillArea.start.row + averageSpillArea.start.row;
              const avgCol = col - spillArea.start.col + averageSpillArea.start.col;
              
              // Check bounds
              if (avgRow <= (averageSpillArea.end.row.type === "number" ? averageSpillArea.end.row.value : Infinity) &&
                  avgCol <= (averageSpillArea.end.col.type === "number" ? averageSpillArea.end.col.value : Infinity)) {
                
                const averageCell = averageRangeResult.evaluate(
                  { x: avgCol - averageSpillArea.start.col, y: avgRow - averageSpillArea.start.row }, 
                  context
                );
                
                if (averageCell?.type === "value") {
                  if (averageCell.result.type === "number") {
                    sum += averageCell.result.value;
                    count++;
                  } else if (averageCell.result.type === "infinity") {
                    return {
                      type: "value",
                      result: { type: "infinity", sign: averageCell.result.sign },
                    };
                  }
                }
              }
            }
          }
        }
        
        if (count === 0) {
          return {
            type: "error",
            err: FormulaError.DIV0,
            message: "No numeric values match the criteria",
          };
        }
        
        return {
          type: "value",
          result: { type: "number", value: sum / count },
        };
      }
    }

    // Handle different range types
    if (rangeResult.type === "value" && averageRangeResult.type === "value") {
      // Single value case
      if (matchesParsedCriteria(rangeResult.result, parsedCriteria)) {
        // Check if the average value is numeric
        if (averageRangeResult.result.type === "number") {
          sum = averageRangeResult.result.value;
          count = 1;
        } else if (averageRangeResult.result.type === "infinity") {
          return {
            type: "value",
            result: { type: "infinity", sign: averageRangeResult.result.sign },
          };
        }
        // Non-numeric values are ignored (like Excel)
      }
    } else if (rangeResult.type === "spilled-values" && averageRangeResult.type === "spilled-values") {
      // Range case - iterate over both ranges in parallel
      const criteriaValues = Array.from(rangeResult.evaluateAllCells.call(this, {
        context,
        evaluate: rangeResult.evaluate,
        origin: context.currentCell,
      }));

      const averageValues = Array.from(averageRangeResult.evaluateAllCells.call(this, {
        context,
        evaluate: averageRangeResult.evaluate,
        origin: context.currentCell,
      }));

      // Both ranges should have the same size for proper parallel iteration
      const minLength = Math.min(criteriaValues.length, averageValues.length);

      for (let i = 0; i < minLength; i++) {
        const criteriaCell = criteriaValues[i];
        const averageCell = averageValues[i];

        if (criteriaCell?.result.type === "error" || averageCell?.result.type === "error") {
          // Skip error cells
          continue;
        }

        if (criteriaCell?.result.type === "value" && averageCell?.result.type === "value") {
          if (matchesParsedCriteria(criteriaCell.result.result, parsedCriteria)) {
            // Only include numeric values in the average
            if (averageCell.result.result.type === "number") {
              sum += averageCell.result.result.value;
              count++;
            } else if (averageCell.result.result.type === "infinity") {
              // If any value is infinity, return infinity
              return {
                type: "value",
                result: { type: "infinity", sign: averageCell.result.result.sign },
              };
            }
            // Non-numeric values are ignored
          }
        }
      }
    } else if (rangeResult.type === "value" && averageRangeResult.type === "spilled-values") {
      // Single criteria value, range of average values
      if (matchesParsedCriteria(rangeResult.result, parsedCriteria)) {
        const averageValues = averageRangeResult.evaluateAllCells.call(this, {
          context,
          evaluate: averageRangeResult.evaluate,
          origin: context.currentCell,
        });

        for (const averageCell of averageValues) {
          if (averageCell?.result.type === "error") {
            continue;
          }
          if (averageCell?.result.type === "value") {
            if (averageCell.result.result.type === "number") {
              sum += averageCell.result.result.value;
              count++;
            } else if (averageCell.result.result.type === "infinity") {
              return {
                type: "value",
                result: { type: "infinity", sign: averageCell.result.result.sign },
              };
            }
          }
        }
      }
    } else if (rangeResult.type === "spilled-values" && averageRangeResult.type === "value") {
      // Range of criteria values, single average value
      const criteriaValues = rangeResult.evaluateAllCells.call(this, {
        context,
        evaluate: rangeResult.evaluate,
        origin: context.currentCell,
      });

      for (const criteriaCell of criteriaValues) {
        if (criteriaCell?.result.type === "error") {
          continue;
        }
        if (criteriaCell?.result.type === "value") {
          if (matchesParsedCriteria(criteriaCell.result.result, parsedCriteria)) {
            if (averageRangeResult.result.type === "number") {
              sum += averageRangeResult.result.value;
              count++;
            } else if (averageRangeResult.result.type === "infinity") {
              return {
                type: "value",
                result: { type: "infinity", sign: averageRangeResult.result.sign },
              };
            }
          }
        }
      }
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid range argument types",
      };
    }

    // Return error if no matching numeric values found
    if (count === 0) {
      return {
        type: "error",
        err: FormulaError.DIV0,
        message: "No numeric values match the criteria",
      };
    }

    return {
      type: "value",
      result: { type: "number", value: sum / count },
    };
  },
};
