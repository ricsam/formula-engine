import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type CellInfinity,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { FunctionNode, ASTNode } from "src/parser/ast";
import {
  parseCriteria,
  matchesParsedCriteria,
  type ParsedCriteria,
} from "./criteria-parser";

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
): CriteriaPair[] | { type: "error"; err: FormulaError; message: string } {
  const criteriaPairs: CriteriaPair[] = [];

  for (let i = startIndex; i < node.args.length; i += 2) {
    const criteriaRangeResult = evaluateNode.call(this, node.args[i]!, context);
    if (criteriaRangeResult.type === "error") {
      return criteriaRangeResult;
    }

    const criteriaResult = evaluateNode.call(this, node.args[i + 1]!, context);
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
 * Generator that yields matching values from ranges based on criteria
 * Much more efficient than creating intermediate arrays
 * @param evaluator - The function evaluator context
 * @param valueRangeResult - The range of values to process
 * @param criteriaPairs - Array of criteria pairs to match against
 * @param context - Evaluation context
 * @yields SingleEvaluationResult for each matching cell - includes both values and errors
 * Consuming functions decide whether to skip errors or propagate them
 */
export function* processMultiCriteriaValues(
  evaluator: FormulaEvaluator,
  valueRangeResult: FunctionEvaluationResult,
  criteriaPairs: CriteriaPair[],
  context: EvaluationContext
): Generator<SingleEvaluationResult, void, unknown> {
  // Check if this is an empty criteria case
  if (
    hasEmptyCriteria(criteriaPairs) &&
    valueRangeResult.type === "spilled-values"
  ) {
    return yield* handleEmptyCriteriaSpilledValues(
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
      yield valueRangeResult; // Yield the entire result (ValueEvaluationResult)
    }
  } else if (valueRangeResult.type === "spilled-values") {
    // Range case - iterate over all values and check criteria
    const valueValues = Array.from(
      valueRangeResult.evaluateAllCells.call(evaluator, {
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
            rangeResult.evaluateAllCells.call(evaluator, {
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

      if (!valueCell) {
        continue; // Skip undefined cells
      }

      // Check if all criteria match for this position
      let allMatch = true;
      for (let j = 0; j < criteriaPairs.length; j++) {
        const criteriaCell = criteriaValueArrays[j]?.[i];
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
        yield valueCell.result; // Yield the entire SingleEvaluationResult (includes errors and values)
      }
    }
  }
}

/**
 * Validate multi-criteria function arguments (odd number, min 3)
 */
export function validateMultiCriteriaArgs(
  functionName: string,
  argCount: number
): { type: "error"; err: FormulaError; message: string } | null {
  if (argCount < 3 || (argCount - 1) % 2 !== 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `${functionName} function requires an odd number of arguments (min 3): value_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...`,
    };
  }
  return null;
}

/**
 * Validate COUNTIFS function arguments (even number, min 2)
 */
export function validateCountifsArgs(
  argCount: number
): { type: "error"; err: FormulaError; message: string } | null {
  if (argCount < 2 || argCount % 2 !== 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message:
        "COUNTIFS function requires an even number of arguments (min 2): criteria_range1, criteria1, [criteria_range2, criteria2], ...",
    };
  }
  return null;
}

/**
 * Validate single criteria function arguments (2 or 3)
 */
export function validateSingleCriteriaArgs(
  functionName: string,
  argCount: number
): { type: "error"; err: FormulaError; message: string } | null {
  if (argCount < 2 || argCount > 3) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `${functionName} function takes 2 or 3 arguments`,
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
 * Handle empty cell criteria for spilled values - returns a generator if this is an empty criteria case
 * This is needed when criteria like "=" (empty) or "<>" (non-empty) are used with ranges
 *
 * @param valueRangeResult - The range of values to process
 * @param criteriaPairs - Array of criteria pairs
 * @param context - Evaluation context
 * @returns Generator for matching values or null if not an empty criteria case
 */
function* handleEmptyCriteriaSpilledValues(
  valueRangeResult: FunctionEvaluationResult,
  criteriaPairs: CriteriaPair[],
  context: EvaluationContext
): Generator<SingleEvaluationResult, void, unknown> {
  if (
    !hasEmptyCriteria(criteriaPairs) ||
    valueRangeResult.type !== "spilled-values"
  ) {
    throw new Error("Not an empty criteria case");
  }
  const valueSpillArea = valueRangeResult.spillArea(context.currentCell);

  // Check if this is an infinite range - yield error
  if (
    valueSpillArea.end.col.type === "infinity" ||
    valueSpillArea.end.row.type === "infinity"
  ) {
    yield {
      type: "error",
      err: FormulaError.VALUE,
      message: "Cannot process infinite ranges with empty cell criteria",
    };
    return;
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
          let criteriaCell;

          if (rangeResult.type === "spilled-values") {
            const criteriaSpillArea = rangeResult.spillArea(
              context.currentCell
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
          yield valueCell; // Yield the matching cell result
        }
      }
    }
  }
}
