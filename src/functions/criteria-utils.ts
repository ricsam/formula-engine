import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type CellInfinity,
  type ValueEvaluationResult,
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

    if (criteriaResult.type === "awaiting-evaluation") {
      continue;
    }

    let result: ValueEvaluationResult;

    if (criteriaResult.type === "spilled-values") {
      // just take the first spilled value
      const firstSpilledValue = criteriaResult.evaluate(
        { x: 0, y: 0 },
        context
      );
      if (firstSpilledValue.type === "error") {
        return firstSpilledValue;
      } else if (firstSpilledValue.type === "awaiting-evaluation") {
        continue;
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
  context: EvaluationContext,
  lookupOrder: LookupOrder
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
    if (flags.isProfiling) {
      console.time("criteriaValueArrays");
    }
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
          yield {
            type: "error",
            err: FormulaError.VALUE,
            message:
              "Criteria range dimensions do not match value range dimensions",
          };
          return;
        }
      }
      // Single values (type === "value") are compatible with any range size
    }

    // Now use efficient iterator-based processing (no Array.from conversion)
    const valueIterator = valueRangeResult.evaluateAllCells.call(evaluator, {
      context,
      evaluate: valueRangeResult.evaluate,
      origin: context.originCell.cellAddress,
      lookupOrder,
    });

    const criteriaIterators = criteriaPairs.map(({ rangeResult }) => {
      if (rangeResult.type === "spilled-values") {
        return rangeResult.evaluateAllCells.call(evaluator, {
          context,
          evaluate: rangeResult.evaluate,
          origin: context.originCell.cellAddress,
          lookupOrder,
        });
      } else {
        // Single value - create a repeating iterator
        const singleValue = rangeResult;
        return (function* () {
          while (true) {
            yield { result: singleValue };
          }
        })();
      }
    });

    // Create synchronized iteration using a manual approach
    const valueIteratorResult = valueIterator[Symbol.iterator]();
    const criteriaIteratorResults = criteriaIterators.map((iter) =>
      iter[Symbol.iterator]()
    );
    let i = 0;

    const durations: number[] = [];

    let valueNext = valueIteratorResult.next();
    while (!valueNext.done) {
      let duration = performance.now();
      const valueCell = valueNext.value;
      i++;

      if (!valueCell) {
        valueNext = valueIteratorResult.next();
        continue; // Skip undefined cells
      }

      if (flags.isProfiling && i === 1) {
        console.time("step 1");
      }

      if (flags.isProfiling && i === 1) {
        console.log("@criteriaIteratorResults", criteriaIteratorResults);
      }

      const criteriaDurations: number[] = [];
      // Get corresponding criteria values for this position
      const criteriaNextResults = criteriaIteratorResults.map(
        function criteriaNextResults(iter, innerIndex) {
          const duration = performance.now();
          if (
            flags.isProfiling &&
            i === 1 &&
            innerIndex === 0 &&
            !flags.profilingNamespaces["criteria-utils-profiled"]
          ) {
            flags.profilingNamespaces["criteria-utils"] = true;
            // console.profile("What is going on here?");
          }
          const result = iter.next();
          if (
            flags.isProfiling &&
            i === 1 &&
            innerIndex === 0 &&
            !flags.profilingNamespaces["criteria-utils-profiled"]
          ) {
            flags.profilingNamespaces["criteria-utils"] = false;
            // console.profileEnd("What is going on here?");
          }
          if (
            flags.isProfiling &&
            i === 1 &&
            innerIndex === 0 &&
            !flags.profilingNamespaces["criteria-utils-profiled"]
          ) {
            flags.profilingNamespaces["criteria-utils-profiled"] = true;
          }
          criteriaDurations.push(performance.now() - duration);
          return result;
        }
      );

      if (flags.isProfiling && i === 1) {
        console.log("@criteriaDurations", criteriaDurations);
      }

      if (flags.isProfiling && i === 1) {
        console.timeEnd("step 1");
      }

      // Check if any criteria iterator is done (shouldn't happen if dimensions match)
      if (criteriaNextResults.some((result) => result.done)) {
        break;
      }

      if (flags.isProfiling && i === 1) {
        console.time("step 2");
      }

      // Check if all criteria match for this position
      let allMatch = true;
      for (let j = 0; j < criteriaPairs.length; j++) {
        const criteriaCell = criteriaNextResults[j]?.value?.result;

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

      if (flags.isProfiling && i === 1) {
        console.timeEnd("step 2");
      }

      if (allMatch) {
        if (flags.isProfiling) {
          console.log(
            "All match for value cell and criteria cells",
            valueCell.result,
            criteriaNextResults,
            i
          );
        }
        yield valueCell.result; // Yield the entire SingleEvaluationResult (includes errors and values)
      }

      // Advance to next value
      if (flags.isProfiling && i === 1) {
        console.time("step 3");
      }
      valueNext = valueIteratorResult.next();
      durations.push(performance.now() - duration);
      if (flags.isProfiling && i === 1) {
        console.timeEnd("step 3");
      }
    }

    if (flags.isProfiling) {
      const averageDuration =
        durations.reduce((a, b) => a + b, 0) / durations.length;
      console.log("average duration: " + averageDuration);
      const maxDurationIndex = durations.indexOf(Math.max(...durations));
      console.log("min duration: " + Math.min(...durations));
      console.log("max duration: " + Math.max(...durations));
      console.log("max duration index: " + maxDurationIndex);

      // Log top 10 slowest indexes
      const indexedDurations = durations.map((duration, index) => ({
        duration,
        index,
      }));
      const top10Slowest = indexedDurations
        .sort((a, b) => b.duration - a.duration)
        .slice(0, 10);
      console.log("top 10 slowest indexes:", top10Slowest);

      console.log("processed " + i + " values");
      console.timeEnd("criteriaValueArrays");
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
  const valueSpillArea = valueRangeResult.spillArea(
    context.originCell.cellAddress
  );

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
          yield valueCell; // Yield the matching cell result
        }
      }
    }
  }
}
