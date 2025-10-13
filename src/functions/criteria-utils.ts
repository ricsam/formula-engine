import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type CellInfinity,
  type ValueEvaluationResult,
  type ErrorEvaluationResult,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { FunctionNode, ASTNode } from "src/parser/ast";
import {
  parseCriteria,
  matchesParsedCriteria,
  type ParsedCriteria,
} from "./criteria-parser";
import { flags } from "src/debug/flags";
import { getRelativeRange } from "src/core/utils";
import type { LookupOrder } from "src/core/managers";
import {
  AwaitingEvaluationError,
  EvaluationError,
} from "src/evaluator/evaluation-error";

/**
 * Criteria pair for criteria-based functions
 */
export interface CriteriaPair {
  rangeResult: FunctionEvaluationResult;
  parsedCriteria: ParsedCriteria;
}

/**
 * Parse criteria pairs from function arguments
 * @param node - The function node
 * @param context - Evaluation context
 * @param evaluateNode - Node evaluation function
 * @param startIndex - Index to start parsing criteria pairs from (1 for IFS functions, 0 for COUNTIFS)
 * @returns Array of criteria pairs or error
 */
export function parseCriteriaPairs(
  this: FormulaEvaluator,
  node: FunctionNode,
  context: EvaluationContext,
  evaluateNode: (
    node: ASTNode,
    context: EvaluationContext
  ) => FunctionEvaluationResult,
  startIndex: number = 1
): CriteriaPair[] | ErrorEvaluationResult {
  const criteriaPairs: CriteriaPair[] = [];

  for (let i = startIndex; i < node.args.length; i += 2) {
    const criteriaRangeResult = evaluateNode.call(this, node.args[i]!, context);
    if (
      criteriaRangeResult.type === "error" ||
      criteriaRangeResult.type === "awaiting-evaluation"
    ) {
      return criteriaRangeResult;
    }

    const criteriaResult = evaluateNode.call(this, node.args[i + 1]!, context);
    if (
      criteriaResult.type === "error" ||
      criteriaResult.type === "awaiting-evaluation"
    ) {
      return criteriaResult;
    }

    let result: ValueEvaluationResult;

    if (criteriaResult.type === "spilled-values") {
      // just take the first spilled value
      const firstSpilledValue = criteriaResult.evaluate(
        { x: 0, y: 0 },
        context
      );
      if (
        firstSpilledValue.type === "error" ||
        firstSpilledValue.type === "awaiting-evaluation"
      ) {
        return firstSpilledValue;
      } else {
        result = firstSpilledValue;
      }
    } else {
      result = criteriaResult;
    }

    const parsedCriteria = parseCriteria(result.result);
    if (parsedCriteria.type === "error") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: parsedCriteria.message,
        errAddress: context.originCell.cellAddress,
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
      parsedCriteria.type === "exact" &&
      parsedCriteria.value.type === "string" &&
      parsedCriteria.value.value === ""
  );
}

/**
 * Returns matching values from ranges based on criteria
 * @param evaluator - The function evaluator context
 * @param valueRangeResult - The range of values to process
 * @param criteriaPairs - Array of criteria pairs to match against
 * @param context - Evaluation context
 * @returns Array of SingleEvaluationResult for each matching cell - includes both values and errors
 * Consuming functions decide whether to skip errors or propagate them
 */
export function processMultiCriteriaValues(
  evaluator: FormulaEvaluator,
  valueRangeResult: FunctionEvaluationResult,
  criteriaPairs: CriteriaPair[],
  context: EvaluationContext,
  lookupOrder: LookupOrder
): SingleEvaluationResult[] {
  const results: SingleEvaluationResult[] = [];

  // Check if this is an empty criteria case
  if (
    hasEmptyCriteria(criteriaPairs) &&
    valueRangeResult.type === "spilled-values"
  ) {
    return handleEmptyCriteriaSpilledValues(
      valueRangeResult,
      criteriaPairs,
      context
    );
  }

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
      results.push(valueRangeResult); // Add the entire result (ValueEvaluationResult)
    }
  } else if (valueRangeResult.type === "spilled-values") {
    // Range case - first validate dimensions using spillArea for efficiency
    const valueSpillArea = valueRangeResult.spillArea(
      context.originCell.cellAddress
    );

    // Check that all criteria ranges have compatible dimensions
    for (const { rangeResult } of criteriaPairs) {
      if (rangeResult.type === "spilled-values") {
        const criteriaSpillArea = rangeResult.spillArea(
          context.originCell.cellAddress
        );

        // Compare dimensions using relative ranges to get width/height
        const valueRelRange = getRelativeRange(
          valueSpillArea,
          context.originCell.cellAddress
        );
        const criteriaRelRange = getRelativeRange(
          criteriaSpillArea,
          context.originCell.cellAddress
        );

        // Check if dimensions are compatible
        const widthsMatch =
          (valueRelRange.width.type === "infinity" &&
            criteriaRelRange.width.type === "infinity") ||
          (valueRelRange.width.type === "number" &&
            criteriaRelRange.width.type === "number" &&
            valueRelRange.width.value === criteriaRelRange.width.value);

        const heightsMatch =
          (valueRelRange.height.type === "infinity" &&
            criteriaRelRange.height.type === "infinity") ||
          (valueRelRange.height.type === "number" &&
            criteriaRelRange.height.type === "number" &&
            valueRelRange.height.value === criteriaRelRange.height.value);

        if (!widthsMatch || !heightsMatch) {
          // Return #VALUE! error for dimension mismatch
          throw new EvaluationError(
            FormulaError.VALUE,
            "Criteria range dimensions do not match value range dimensions"
          );
        }
      }
      // Single values (type === "value") are compatible with any range size
    }

    // Get all cells to simplify position-based pairing
    const valueResults = valueRangeResult.evaluateAllCells.call(evaluator, {
      context,
      evaluate: valueRangeResult.evaluate,
      origin: context.originCell.cellAddress,
      lookupOrder,
    });

    // Get criteria ranges (or track if single value)
    const criteriaResultsArrays = criteriaPairs.map(({ rangeResult }) => {
      if (rangeResult.type === "spilled-values") {
        const criteriaResults = rangeResult.evaluateAllCells.call(evaluator, {
          context,
          evaluate: rangeResult.evaluate,
          origin: context.originCell.cellAddress,
          lookupOrder,
        });
        return { type: "array" as const, results: criteriaResults };
      } else {
        // Single value - we'll use it for all positions
        return { type: "single" as const, value: rangeResult };
      }
    });

    // Create maps from position to value for each criteria range
    const criteriaMaps = criteriaResultsArrays.map((criteriaData) => {
      if (criteriaData.type === "single") {
        return null; // Single values don't need a map
      }
      const map = new Map<string, SingleEvaluationResult>();
      for (const { result, relativePos } of criteriaData.results) {
        const key = `${relativePos.x},${relativePos.y}`;
        map.set(key, result);
      }
      return map;
    });

    // Iterate through each value position
    for (const valueCell of valueResults) {
      if (!valueCell) {
        continue;
      }

      const posKey = `${valueCell.relativePos.x},${valueCell.relativePos.y}`;

      // Check if all criteria match for this position
      let allMatch = true;
      for (let j = 0; j < criteriaPairs.length; j++) {
        let criteriaCell: SingleEvaluationResult | undefined;

        if (criteriaResultsArrays[j]!.type === "single") {
          // Single value - use it for all positions
          criteriaCell = criteriaResultsArrays[j]!.value;
        } else {
          // Array - look up by position, treat missing as empty
          criteriaCell = criteriaMaps[j]?.get(posKey);
        }

        // If criteriaCell is undefined, treat it as an empty cell
        if (!criteriaCell) {
          // Check if the criteria matches empty cells
          const parsedCriteria = criteriaPairs[j]!.parsedCriteria;
          const emptyValue = { type: "string" as const, value: "" };
          if (!matchesParsedCriteria(emptyValue, parsedCriteria)) {
            allMatch = false;
            break;
          }
        } else if (criteriaCell.type === "error") {
          allMatch = false;
          break;
        } else if (criteriaCell.type === "value") {
          if (
            !matchesParsedCriteria(
              criteriaCell.result,
              criteriaPairs[j]!.parsedCriteria
            )
          ) {
            allMatch = false;
            break;
          }
        } else if (criteriaCell.type === "awaiting-evaluation") {
          throw new AwaitingEvaluationError(
            context.originCell.cellAddress,
            criteriaCell.waitingFor
          );
        } else {
          allMatch = false;
          break;
        }
      }

      if (allMatch) {
        results.push(valueCell.result); // Add the entire SingleEvaluationResult (includes errors and values)
      }
    }
  }

  return results;
}

/**
 * Validate multi-criteria function arguments (odd number, min 3)
 */
export function validateMultiCriteriaArgs(
  functionName: string,
  argCount: number,
  context: EvaluationContext
): ErrorEvaluationResult | null {
  if (argCount < 3 || (argCount - 1) % 2 !== 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `${functionName} function requires an odd number of arguments (min 3): value_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...`,
      errAddress: context.originCell.cellAddress,
    };
  }
  return null;
}

/**
 * Validate COUNTIFS function arguments (even number, min 2)
 */
export function validateCountifsArgs(
  argCount: number,
  context: EvaluationContext
): ErrorEvaluationResult | null {
  if (argCount < 2 || argCount % 2 !== 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message:
        "COUNTIFS function requires an even number of arguments (min 2): criteria_range1, criteria1, [criteria_range2, criteria2], ...",
      errAddress: context.originCell.cellAddress,
    };
  }
  return null;
}

/**
 * Validate single criteria function arguments (2 or 3)
 */
export function validateSingleCriteriaArgs(
  functionName: string,
  argCount: number,
  context: EvaluationContext
): ErrorEvaluationResult | null {
  if (argCount < 2 || argCount > 3) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `${functionName} function takes 2 or 3 arguments`,
      errAddress: context.originCell.cellAddress,
    };
  }
  return null;
}

/**
 * Result type for processInfinity - either return infinity immediately or continue with updated state
 */
export type ProcessInfinityResult<T> =
  | { type: "infinity"; infinity: CellInfinity }
  | { type: "state"; state: T };

/**
 * Handle empty cell criteria for spilled values - returns matching values for empty criteria cases
 * This is needed when criteria like "=" (empty) or "<>" (non-empty) are used with ranges
 *
 * @param valueRangeResult - The range of values to process
 * @param criteriaPairs - Array of criteria pairs
 * @param context - Evaluation context
 * @returns Array of matching values
 */
function handleEmptyCriteriaSpilledValues(
  valueRangeResult: FunctionEvaluationResult,
  criteriaPairs: CriteriaPair[],
  context: EvaluationContext
): SingleEvaluationResult[] {
  if (
    !hasEmptyCriteria(criteriaPairs) ||
    valueRangeResult.type !== "spilled-values"
  ) {
    throw new Error("Not an empty criteria case");
  }
  const results: SingleEvaluationResult[] = [];
  const valueSpillArea = valueRangeResult.spillArea(
    context.originCell.cellAddress
  );

  // Check if this is an infinite range and we are counting empty cells - throw error
  if (
    valueSpillArea.end.col.type === "infinity" ||
    valueSpillArea.end.row.type === "infinity"
  ) {
    throw new EvaluationError(
      FormulaError.VALUE,
      "Can not process infinite ranges with empty cell criteria"
    );
  }

  // Finite range - iterate over all cells in spill areas
  if (
    valueSpillArea.end.col.type === "number" &&
    valueSpillArea.end.row.type === "number"
  ) {
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
        let allMatch = true;
        for (const { rangeResult, parsedCriteria } of criteriaPairs) {
          let criteriaCell: SingleEvaluationResult | undefined;

          if (rangeResult.type === "spilled-values") {
            const criteriaSpillArea = rangeResult.spillArea(
              context.originCell.cellAddress
            );
            const criteriaRow =
              row - valueSpillArea.start.row + criteriaSpillArea.start.row;
            const criteriaCol =
              col - valueSpillArea.start.col + criteriaSpillArea.start.col;

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
            // Check if criteria cell matches empty condition
            const isEmpty =
              !criteriaCell ||
              (criteriaCell.type === "value" &&
                criteriaCell.result.type === "string" &&
                criteriaCell.result.value === "");

            const shouldMatch =
              parsedCriteria.type === "exact" ? isEmpty : !isEmpty;
            if (!shouldMatch) {
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

        if (allMatch && valueCell) {
          results.push(valueCell); // Add the matching cell result
        }
      }
    }
  }

  return results;
}
