import {
  FormulaError,
  type CellString,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type EvaluationContext,
  type CellAddress,
} from "src/core/types";
import type { FormulaEngine } from "src/core/engine";

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
 * Generic substring operation for LEFT and RIGHT functions
 */
export function substringOperation(
  textResult: FunctionEvaluationResult,
  numCharsResult: FunctionEvaluationResult,
  direction: "left" | "right"
): CellString {
  const textStr = convertToString(textResult);
  const numChars = extractNumericValue(numCharsResult);

  // Validate numChars
  if (numChars < 0) {
    throw new Error(FormulaError.VALUE);
  }

  let result: string;
  if (direction === "left") {
    result = textStr.substring(0, Math.floor(numChars));
  } else {
    const start = Math.max(0, textStr.length - Math.floor(numChars));
    result = textStr.substring(start);
  }

  return { type: "string", value: result };
}

/**
 * Helper for creating spilled-values result for text functions
 */
export function createTextSpilledResult(
  this: FormulaEngine,
  {
    operation,
    textResult,
    numCharsResult,
    context,
    functionName,
  }: {
    operation: (
      textResult: FunctionEvaluationResult,
      numCharsResult: FunctionEvaluationResult
    ) => CellString;
    textResult: FunctionEvaluationResult;
    numCharsResult: FunctionEvaluationResult;
    context: EvaluationContext;
    functionName: string;
  }
): FunctionEvaluationResult {
  // Handle spilled-values input - return spilled-values for spilling
  if (textResult.type === "spilled-values") {
    // If both arguments are spilled-values, we need to zip them together
    if (numCharsResult.type === "spilled-values") {
      if (numCharsResult.originResult.type !== "number") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Invalid numChars argument",
        };
      }

      // Calculate origin result using origin values from both spilled arrays
      const textValue: FunctionEvaluationResult = {
        type: "value",
        result: textResult.originResult,
      };
      const numCharsValue: FunctionEvaluationResult = {
        type: "value",
        result: numCharsResult.originResult,
      };
      
      let originCellValue: CellString;
      try {
        originCellValue = operation(textValue, numCharsValue);
      } catch (error) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: `${functionName} operation failed`,
        };
      }

      // Create unified spill area (union of both ranges)
      const spillArea = this.unionRanges(
        this.projectRange(textResult.spillArea, context.currentCell),
        this.projectRange(numCharsResult.spillArea, context.currentCell)
      );

      return {
        type: "spilled-values",
        spillArea,
        spillOrigin: context.currentCell,
        source: `${functionName} with zipped spilled text and numChars values`,
        originResult: originCellValue,
        evaluate: (spilledCell: { address: CellAddress; spillOffset: { x: number; y: number } }, evalContext: EvaluationContext) => {
          // Evaluate both spilled arrays at this position
          const spillTextResult = textResult.evaluate(spilledCell, evalContext);
          const spillNumCharsResult = numCharsResult.evaluate(spilledCell, evalContext);

          if (!spillTextResult || spillTextResult.type === "error") {
            return spillTextResult;
          }
          if (!spillNumCharsResult || spillNumCharsResult.type === "error") {
            return spillNumCharsResult;
          }

          try {
            return {
              type: "value",
              result: operation(spillTextResult, spillNumCharsResult),
            };
          } catch (error) {
            return {
              type: "error",
              err: FormulaError.VALUE,
              message: `${functionName} operation failed`,
            };
          }
        },
      };
    }

    // Single numChars value with spilled text values
    if (
      numCharsResult.type !== "value" ||
      numCharsResult.result.type !== "number"
    ) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid numChars argument",
      };
    }

          const textValue: ValueEvaluationResult = {
        type: "value",
        result: textResult.originResult,
      };
    
    let originCellValue: CellString;
    try {
      originCellValue = operation(textValue, numCharsResult);
    } catch (error) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `${functionName} operation failed`,
      };
    }

    return {
      type: "spilled-values",
      spillArea: textResult.spillArea,
      spillOrigin: context.currentCell,
      source: `${functionName} with spilled text values`,
      originResult: originCellValue,
      evaluate: (spilledCell: { address: CellAddress; spillOffset: { x: number; y: number } }, evalContext: EvaluationContext) => {
        const spillResult = textResult.evaluate(spilledCell, evalContext);
        if (!spillResult || spillResult.type === "error") {
          return spillResult;
        }
        try {
          return {
            type: "value",
            result: operation(spillResult, numCharsResult),
          };
        } catch (error) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: `${functionName} operation failed`,
          };
        }
      },
    };
  }

  // Single text value with spilled numChars
  if (numCharsResult.type === "spilled-values") {
    if (textResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid text argument",
      };
    }

    const numCharsValue: ValueEvaluationResult = {
      type: "value",
      result: numCharsResult.originResult,
    };
    
    let originCellValue: CellString;
    try {
      originCellValue = operation(textResult, numCharsValue);
    } catch (error) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `${functionName} operation failed`,
      };
    }

    return {
      type: "spilled-values",
      spillArea: numCharsResult.spillArea,
      spillOrigin: context.currentCell,
      source: `${functionName} with spilled numChars values`,
      originResult: originCellValue,
      evaluate: (spilledCell: { address: CellAddress; spillOffset: { x: number; y: number } }, evalContext: EvaluationContext) => {
        const spillResult = numCharsResult.evaluate(spilledCell, evalContext);
        if (!spillResult || spillResult.type === "error") {
          return spillResult;
        }
        try {
          return {
            type: "value",
            result: operation(textResult, spillResult),
          };
        } catch (error) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: `${functionName} operation failed`,
          };
        }
      },
    };
  }

  // No spilled values - this shouldn't be called in this case
  throw new Error("createTextSpilledResult called without spilled values");
}
