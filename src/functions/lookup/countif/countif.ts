import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type EvaluationContext,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
} from "src/core/types";
import type { FormulaEngine } from "src/core/engine";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import { OpenRangeEvaluator } from "../../math/open-range-evaluator";
import { parseCriteria, matchesParsedCriteria } from "../criteria-parser";

/**
 * COUNTIF function - Counts cells in a range that meet a criteria
 * 
 * Usage: COUNTIF(range, criteria)
 * 
 * range: The range of cells to evaluate
 * criteria: The criteria to match against. Can be:
 *   - Exact value: "Apple", 42
 *   - Comparison: ">10", "<=5", "<>0"
 *   - Wildcards: "App*", "?ruit"
 * 
 * Examples:
 *   COUNTIF(A1:A10, "Apple") - counts cells containing "Apple"
 *   COUNTIF(B1:B10, ">10") - counts cells with values greater than 10
 *   COUNTIF(C1:C10, "App*") - counts cells starting with "App"
 * 
 * Note:
 * - Supports type coercion for comparisons
 * - Case-sensitive string matching
 * - Wildcards: * matches any sequence, ? matches any single character
 */



/**
 * COUNTIF function implementation
 */
export const COUNTIF: FunctionDefinition = {
  name: "COUNTIF",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "COUNTIF function takes exactly 2 arguments",
      };
    }

    // Evaluate range argument
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
        message: "COUNTIF criteria must be a single value",
      };
    }

    let count = 0;

    // Parse the criteria using the new parser
    const parsedCriteria = parseCriteria(criteriaResult.result);
    if (parsedCriteria.type === "error") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: parsedCriteria.message,
      };
    }

    // Special case: counting empty cells over infinite ranges
    if (rangeResult.type === "spilled-values" && 
        parsedCriteria.type === "exact" && 
        parsedCriteria.value.type === "string" && 
        parsedCriteria.value.value === "") {
      
      const spillArea = rangeResult.spillArea(context.currentCell);
      
      // Check if this is an infinite range
      if (spillArea.end.col.type === "infinity" || spillArea.end.row.type === "infinity") {
        // Return infinity for infinite empty cell count
        return {
          type: "value",
          result: { type: "infinity", sign: "positive" },
        };
      }
      
      // Finite range - count empty cells by iterating over all cells in spill area
      if (spillArea.end.col.type === "number" && spillArea.end.row.type === "number") {
        let emptyCount = 0;
        
        for (let row = spillArea.start.row; row <= spillArea.end.row.value; row++) {
          for (let col = spillArea.start.col; col <= spillArea.end.col.value; col++) {
            const cellResult = rangeResult.evaluate(
              { x: col - spillArea.start.col, y: row - spillArea.start.row }, 
              context
            );
            
            // If cell is undefined/null or empty string, count it
            if (!cellResult || 
                (cellResult.type === "value" && 
                 cellResult.result.type === "string" && 
                 cellResult.result.value === "")) {
              emptyCount++;
            }
          }
        }
        
        return {
          type: "value",
          result: { type: "number", value: emptyCount },
        };
      }
    }

    // Handle different range types (normal case)
    if (rangeResult.type === "value") {
      // Single value case
      if (matchesParsedCriteria(rangeResult.result, parsedCriteria)) {
        count = 1;
      }
    } else if (rangeResult.type === "spilled-values") {
      // Range case - use evaluateAllCells to iterate over non-empty cells
      const values = rangeResult.evaluateAllCells.call(this, {
        context,
        evaluate: rangeResult.evaluate,
        origin: context.currentCell,
      });

      for (const cellResult of values) {
        if (cellResult.type === "error") {
          // Skip error cells (like Excel does)
          continue;
        }
        if (cellResult.type === "value") {
          if (matchesParsedCriteria(cellResult.result, parsedCriteria)) {
            count++;
          }
        }
      }
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid range argument type",
      };
    }

    return {
      type: "value",
      result: { type: "number", value: count },
    };
  },
};
