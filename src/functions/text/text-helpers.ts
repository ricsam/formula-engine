import {
  FormulaError,
  type CellString,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type CellAddress,
  type SpreadsheetRange,
  type ErrorEvaluationResult,
  type SingleEvaluationResult,
} from "src/core/types";
import type { FormulaEngine } from "src/core/engine";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { EvaluationContext } from "src/evaluator/evaluation-context";

/**
 * Strictly extracts string value without type coercion
 */
export function convertToString(result: FunctionEvaluationResult): string | ErrorEvaluationResult {
  if (result.type === "awaiting-evaluation") {
    return result;
  }
  
  if (result.type !== "value") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Expected a value result",
    };
  }

  if (result.result.type === "string") {
    return result.result.value;
  } else {
    // No type coercion - only strings are valid
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Expected a string value",
    };
  }
}

/**
 * Strictly extracts numeric value without type coercion
 */
export function extractNumericValue(result: FunctionEvaluationResult): number | ErrorEvaluationResult {
  if (result.type === "awaiting-evaluation") {
    return result;
  }
  
  if (result.type !== "value") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Expected a value result",
    };
  }

  if (result.result.type === "number") {
    return result.result.value;
  } else {
    // No type coercion - only numbers are valid
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Expected a number value",
    };
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

  // Check if any of the results are awaiting evaluation or errors
  if (typeof textStr === "object" && (textStr.type === "awaiting-evaluation" || textStr.type === "error")) {
    return textStr;
  }
  if (typeof startNum === "object" && (startNum.type === "awaiting-evaluation" || startNum.type === "error")) {
    return startNum;
  }
  if (typeof numChars === "object" && (numChars.type === "awaiting-evaluation" || numChars.type === "error")) {
    return numChars;
  }

  // At this point, all values should be primitive types
  const textValue = textStr as string;
  const startNumValue = startNum as number;
  const numCharsValue = numChars as number;

  // Validate startNum and numChars
  if (startNumValue < 1) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "StartNum argument must be a positive number",
    };
  }
  if (numCharsValue < 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "NumChars argument must be a positive number",
    };
  }

  // Convert to 0-based index (Excel uses 1-based)
  const startIndex = Math.floor(startNumValue) - 1;
  const length = Math.floor(numCharsValue);

  // Extract substring
  const result = textValue.substring(startIndex, startIndex + length);

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
    ): SingleEvaluationResult => {
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
