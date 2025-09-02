import {
  FormulaError,
  type CellString,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type EvaluationContext,
  type CellAddress,
  type SpreadsheetRange,
  type ErrorEvaluationResult,
  type SingleEvaluationResult,
} from "src/core/types";
import type { FormulaEngine } from "src/core/engine";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

/**
 * Strictly extracts string value without type coercion
 */
export function convertToString(result: FunctionEvaluationResult): string {
  if (result.type !== "value") {
    throw new Error(FormulaError.VALUE);
  }

  if (result.result.type === "string") {
    return result.result.value;
  } else {
    // No type coercion - only strings are valid
    throw new Error(FormulaError.VALUE);
  }
}

/**
 * Strictly extracts numeric value without type coercion
 */
export function extractNumericValue(result: FunctionEvaluationResult): number {
  if (result.type !== "value") {
    throw new Error(FormulaError.VALUE);
  }

  if (result.result.type === "number") {
    return result.result.value;
  } else {
    // No type coercion - only numbers are valid
    throw new Error(FormulaError.VALUE);
  }
}

/**
 * MID substring operation for MID function
 */
export function midOperation(
  textResult: FunctionEvaluationResult,
  startNumResult: FunctionEvaluationResult,
  numCharsResult: FunctionEvaluationResult
): { type: "value"; result: CellString } | ErrorEvaluationResult {
  const textStr = convertToString(textResult);
  const startNum = extractNumericValue(startNumResult);
  const numChars = extractNumericValue(numCharsResult);

  // Validate startNum and numChars
  if (startNum < 1) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "StartNum argument must be a positive number",
    };
  }
  if (numChars < 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "NumChars argument must be a positive number",
    };
  }

  // Convert to 0-based index (Excel uses 1-based)
  const startIndex = Math.floor(startNum) - 1;
  const length = Math.floor(numChars);

  // Extract substring
  const result = textStr.substring(startIndex, startIndex + length);

  return { type: "value", result: { type: "string", value: result } };
}

/**
 * Helper for creating spilled-values result for MID function (3 parameters)
 */
export function createMidSpilledResult(
  this: FormulaEvaluator,
  {
    textResult,
    startNumResult,
    numCharsResult,
    context,
  }: {
    textResult: FunctionEvaluationResult;
    startNumResult: FunctionEvaluationResult;
    numCharsResult: FunctionEvaluationResult;
    context: EvaluationContext;
  }
): FunctionEvaluationResult {
  const hasSpilledText = textResult.type === "spilled-values";
  const hasSpilledStart = startNumResult.type === "spilled-values";
  const hasSpilledNum = numCharsResult.type === "spilled-values";

  if (!hasSpilledText && !hasSpilledStart && !hasSpilledNum) {
    throw new Error("createMidSpilledResult called without spilled values");
  }

  return {
    type: "spilled-values",
    spillArea: (origin: CellAddress) => {
      // Calculate spill area (union of all spilled ranges)
      let spillArea: SpreadsheetRange;

      if (hasSpilledText && textResult.type === "spilled-values") {
        spillArea = textResult.spillArea(origin);
      } else if (hasSpilledStart && startNumResult.type === "spilled-values") {
        spillArea = startNumResult.spillArea(origin);
      } else if (hasSpilledNum && numCharsResult.type === "spilled-values") {
        spillArea = numCharsResult.spillArea(origin);
      } else {
        // This shouldn't happen since we check for spilled values at the beginning
        throw new Error("No spilled values found");
      }

      // Union with other spilled ranges if they exist
      if (
        hasSpilledText &&
        hasSpilledStart &&
        textResult.type === "spilled-values" &&
        startNumResult.type === "spilled-values"
      ) {
        spillArea = this.unionRanges(
          this.projectRange(textResult.spillArea(origin), origin),
          this.projectRange(startNumResult.spillArea(origin), origin)
        );
      }

      if (
        (hasSpilledText || hasSpilledStart) &&
        hasSpilledNum &&
        numCharsResult.type === "spilled-values"
      ) {
        const projectedSpillArea = this.projectRange(spillArea, origin);
        const numSpillArea = this.projectRange(
          numCharsResult.spillArea(origin),
          origin
        );
        spillArea = this.unionRanges(projectedSpillArea, numSpillArea);
      }
      return spillArea;
    },
    source: "MID with spilled values",
    evaluate: (
      spilledCell,
      evalContext
    ): SingleEvaluationResult | undefined => {
      // Evaluate all arguments at this spilled position
      const spillTextResult = hasSpilledText
        ? textResult.evaluate(spilledCell, evalContext)
        : textResult;
      const spillStartResult = hasSpilledStart
        ? startNumResult.evaluate(spilledCell, evalContext)
        : startNumResult;
      const spillNumResult = hasSpilledNum
        ? numCharsResult.evaluate(spilledCell, evalContext)
        : numCharsResult;

      if (!spillTextResult || spillTextResult.type === "error") {
        return spillTextResult;
      }
      if (!spillStartResult || spillStartResult.type === "error") {
        return spillStartResult;
      }
      if (!spillNumResult || spillNumResult.type === "error") {
        return spillNumResult;
      }

      return midOperation(spillTextResult, spillStartResult, spillNumResult);
    },
    evaluateAllCells: (intersectingRange) => {
      throw new Error("WIP: evaluateAllCells for MID is not implemented");
    },
  };
}
