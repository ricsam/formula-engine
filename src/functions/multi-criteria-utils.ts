import {
  FormulaError,
  type CellValue,
  type FunctionEvaluationResult,
  type FunctionDefinition,
  type SpreadsheetRange,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { FunctionNode } from "src/parser/ast";
import { parseCriteria, matchesParsedCriteria, type ParsedCriteria } from "./criteria-parser";

/**
 * Criteria pair for multi-criteria functions
 */
export interface CriteriaPair {
  rangeResult: FunctionEvaluationResult;
  parsedCriteria: ParsedCriteria;
}

/**
 * Parse and validate criteria pairs from function arguments
 * Starting from argument index 1 (after the value range)
 */
export function parseCriteriaPairs(
  node: FunctionNode,
  context: EvaluationContext,
  evaluateNode: (node: any, context: EvaluationContext) => FunctionEvaluationResult
): CriteriaPair[] | { type: "error"; err: FormulaError; message: string } {
  const criteriaPairs: CriteriaPair[] = [];

  for (let i = 1; i < node.args.length; i += 2) {
    const criteriaRangeResult = evaluateNode(node.args[i]!, context);
    if (criteriaRangeResult.type === "error") {
      return criteriaRangeResult;
    }

    const criteriaResult = evaluateNode(node.args[i + 1]!, context);
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

  return criteriaPairs;
}

/**
 * Check if any criteria involves empty/non-empty cell matching
 */
export function hasEmptyCriteria(criteriaPairs: CriteriaPair[]): boolean {
  return criteriaPairs.some(
    ({ parsedCriteria }) =>
      (parsedCriteria.type === "exact" &&
        parsedCriteria.value.type === "string" &&
        parsedCriteria.value.value === "") ||
      (parsedCriteria.type === "comparison" &&
        parsedCriteria.operator === "<>" &&
        parsedCriteria.value.type === "string" &&
        parsedCriteria.value.value === "")
  );
}

/**
 * Check if a cell matches empty criteria (empty string or non-empty)
 */
export function matchesEmptyCriteria(
  criteriaCell: FunctionEvaluationResult | undefined,
  parsedCriteria: ParsedCriteria
): boolean {
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
    return isEmpty;
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
    return !isEmpty;
  }
  return false;
}

/**
 * Check if all criteria match for a given position
 */
export function checkAllCriteriaMatch(
  criteriaPairs: CriteriaPair[],
  valueSpillArea: SpreadsheetRange,
  row: number,
  col: number,
  context: EvaluationContext
): boolean {
  for (const { rangeResult, parsedCriteria } of criteriaPairs) {
    let criteriaCell;
    
    if (rangeResult.type === "spilled-values") {
      const criteriaSpillArea = rangeResult.spillArea(context.currentCell);
      const criteriaRow = row - valueSpillArea.start.row + criteriaSpillArea.start.row;
      const criteriaCol = col - valueSpillArea.start.col + criteriaSpillArea.start.col;

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
    const isEmptyCriteria =
      (parsedCriteria.type === "exact" &&
        parsedCriteria.value.type === "string" &&
        parsedCriteria.value.value === "") ||
      (parsedCriteria.type === "comparison" &&
        parsedCriteria.operator === "<>" &&
        parsedCriteria.value.type === "string" &&
        parsedCriteria.value.value === "");

    if (isEmptyCriteria) {
      if (!matchesEmptyCriteria(criteriaCell, parsedCriteria)) {
        return false;
      }
    } else {
      // Regular criteria matching
      if (
        !criteriaCell ||
        criteriaCell.type !== "value" ||
        !matchesParsedCriteria(criteriaCell.result, parsedCriteria)
      ) {
        return false;
      }
    }
  }

  return true;
}

/**
 * Process values with criteria for range-based evaluation
 * This is a generic utility that can be used by AVERAGEIFS, SUMIFS, COUNTIFS, MAXIFS, MINIFS, etc.
 */
export function processValuesWithCriteria<T>(
  valueRangeResult: FunctionEvaluationResult,
  criteriaPairs: CriteriaPair[],
  context: EvaluationContext,
  evaluatorContext: any,
  processor: (value: number) => T,
  combiner: (results: T[]) => T | { type: "error"; err: FormulaError; message: string },
  infinityHandler?: (sign: "positive" | "negative") => T
): T | { type: "error"; err: FormulaError; message: string } {
  const results: T[] = [];

  // Handle different value range types
  if (valueRangeResult.type === "value") {
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

    if (allMatch) {
      if (valueRangeResult.result.type === "number") {
        results.push(processor(valueRangeResult.result.value));
      } else if (valueRangeResult.result.type === "infinity" && infinityHandler) {
        return infinityHandler(valueRangeResult.result.sign);
      }
    }
  } else if (valueRangeResult.type === "spilled-values") {
    // Range case - iterate over all values and check criteria
    const valueValues = Array.from(
      valueRangeResult.evaluateAllCells.call(evaluatorContext, {
        context,
        evaluate: valueRangeResult.evaluate,
        origin: context.currentCell,
      })
    );

    // Get all criteria values for parallel iteration
    const criteriaValueArrays: Array<Array<FunctionEvaluationResult>> =
      criteriaPairs.map(({ rangeResult }) => {
        if (rangeResult.type === "spilled-values") {
          return Array.from(
            rangeResult.evaluateAllCells.call(evaluatorContext, {
              context,
              evaluate: rangeResult.evaluate,
              origin: context.currentCell,
            })
          ).map((cell) => cell.result);
        } else if (rangeResult.type === "value") {
          // Single criteria value - repeat for all values
          return Array(valueValues.length).fill(rangeResult);
        } else {
          return [];
        }
      });

    // Check that all arrays have compatible lengths
    const minLength = Math.min(
      valueValues.length,
      ...criteriaValueArrays.map((arr) => arr.length)
    );

    for (let i = 0; i < minLength; i++) {
      const valueCell = valueValues[i];

      if (valueCell?.result.type === "error") {
        continue; // Skip error cells
      }

      if (valueCell?.result.type === "value") {
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
          if (valueCell.result.result.type === "number") {
            results.push(processor(valueCell.result.result.value));
          } else if (valueCell.result.result.type === "infinity" && infinityHandler) {
            return infinityHandler(valueCell.result.result.sign);
          }
        }
      }
    }
  } else {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Invalid value range argument type",
    };
  }

  return combiner(results);
}

/**
 * Handle empty cell criteria for spilled values - used by functions that need special empty cell handling
 */
export function processEmptyCriteriaSpilledValues<T>(
  valueRangeResult: FunctionEvaluationResult,
  criteriaPairs: CriteriaPair[],
  context: EvaluationContext,
  processor: (value: number) => void,
  infinityHandler: (sign: "positive" | "negative") => T,
  noValuesError: string
): T | { type: "error"; err: FormulaError; message: string } {
  if (valueRangeResult.type !== "spilled-values") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Expected spilled values for empty criteria processing",
    };
  }

  const valueSpillArea = valueRangeResult.spillArea(context.currentCell);

  // Check if this is an infinite range
  if (
    valueSpillArea.end.col.type === "infinity" ||
    valueSpillArea.end.row.type === "infinity"
  ) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Cannot process infinite ranges with empty cell criteria",
    };
  }

  // Finite range - iterate over all cells in spill areas
  if (
    valueSpillArea.end.col.type === "number" &&
    valueSpillArea.end.row.type === "number"
  ) {
    let hasValues = false;

    for (
      let row = valueSpillArea.start.row;
      row <= valueSpillArea.end.row.value;
      row++
    ) {
      for (
        let col = valueSpillArea.start.col;
        col <= valueSpillArea.end.col.value;
        col++
      ) {
        const valueCell = valueRangeResult.evaluate(
          {
            x: col - valueSpillArea.start.col,
            y: row - valueSpillArea.start.row,
          },
          context
        );

        // Check all criteria for this position
        if (checkAllCriteriaMatch(criteriaPairs, valueSpillArea, row, col, context)) {
          if (valueCell?.type === "value") {
            if (valueCell.result.type === "number") {
              processor(valueCell.result.value);
              hasValues = true;
            } else if (valueCell.result.type === "infinity") {
              return infinityHandler(valueCell.result.sign);
            }
          }
        }
      }
    }

    if (!hasValues) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: noValuesError,
      };
    }
  }

  return {
    type: "error",
    err: FormulaError.VALUE,
    message: "Unexpected error in empty criteria processing",
  };
}
