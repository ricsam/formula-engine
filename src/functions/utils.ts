/**
 * Shared utility functions for function implementations
 * This file breaks circular imports by providing utilities without importing from index.ts
 */

import type { CellValue, FormulaError } from "../core/types";
import type { EvaluationResult } from "../evaluator/evaluator";

// Helper function to coerce values to numbers
export function coerceToNumber(value: CellValue): number {
  if (typeof value === "number") {
    return value;
  }
  if (typeof value === "string") {
    const num = parseFloat(value);
    if (isNaN(num)) {
      throw new Error("#VALUE!");
    }
    return num;
  }
  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }
  if (value === undefined || value === null) {
    return 0;
  }
  throw new Error("#VALUE!");
}

// Helper function to coerce values to strings
export function coerceToString(value: CellValue): string {
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "number") {
    return value.toString();
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (value === undefined || value === null) {
    return "";
  }
  if (isFormulaError(value)) {
    return value;
  }
  return String(value);
}

// Helper function to check if value is a formula error
export function isFormulaError(value: CellValue): value is FormulaError {
  return (
    typeof value === "string" && value.startsWith("#") && value.endsWith("!")
  );
}

// Helper function to propagate errors
export function propagateError(
  args: (CellValue | CellValue[][])[]
): FormulaError | undefined {
  for (const arg of args) {
    if (typeof arg === "string" && isFormulaError(arg)) {
      return arg;
    }
    if (Array.isArray(arg) && arg.length > 0) {
      for (const item of arg) {
        if (typeof item === "string" && isFormulaError(item)) {
          return item;
        }
      }
    }
  }
  return undefined;
}

// Helper function to propagate errors from evaluation results
export function propagateErrorFromEvalResults(
  args: EvaluationResult[]
): FormulaError | undefined {
  for (const arg of args) {
    if (arg.type === "value" && typeof arg.value === "string" && isFormulaError(arg.value)) {
      return arg.value;
    }
    if (arg.type === "2d-array") {
      const flatValues = arg.value.flat();
      for (const value of flatValues) {
        if (typeof value === "string" && isFormulaError(value)) {
          return value;
        }
      }
    }
  }
  return undefined;
}

// Helper function to assert argument is a scalar value (not array)
export function assertScalarArg(
  evalResult: EvaluationResult,
  argIndex: number
): CellValue {
  if (evalResult.type === "2d-array") {
    throw new Error(
      `Argument ${argIndex + 1} must be a scalar value, not an array`
    );
  }
  return evalResult.value;
}

// Helper function to assert argument is a 2D array
export function assertArrayArg(
  evalResult: EvaluationResult,
  argIndex: number
): CellValue[][] {
  if (evalResult.type === "2d-array") {
    if (
      !Array.isArray(evalResult.value) ||
      !Array.isArray(evalResult.value[0])
    ) {
      throw new Error(`Argument ${argIndex + 1} must be an array`);
    }
    return evalResult.value as CellValue[][];
  }
  // Convert scalar to 2D array
  return [[evalResult.value]];
}

// Helper to get scalar value from evaluation result
export function getScalarValue(evalResult: EvaluationResult): CellValue {
  if (evalResult.type === "value") {
    return evalResult.value;
  }
  // For arrays, return the first value
  const array = evalResult.value as CellValue[][];
  return array[0]?.[0];
}

// Helper to safely get scalar value without throwing (overloaded)
export function safeGetScalarValue(
  argEvaluatedValues: EvaluationResult[],
  index: number,
  defaultValue?: CellValue
): CellValue;
export function safeGetScalarValue(evalResult: EvaluationResult): CellValue;
export function safeGetScalarValue(
  arg1: EvaluationResult | EvaluationResult[],
  index?: number,
  defaultValue: CellValue = 0
): CellValue {
  if (Array.isArray(arg1)) {
    const result = arg1[index!];
    if (!result) return defaultValue;
    return getScalarValue(result);
  } else {
    try {
      return getScalarValue(arg1);
    } catch {
      return undefined;
    }
  }
}

// Helper to get array from evaluation result
export function getArrayFromEvalResult(evalResult: EvaluationResult): CellValue[][] {
  if (evalResult.type === "2d-array") {
    return evalResult.value as CellValue[][];
  }
  // Convert scalar to 2D array
  return [[evalResult.value]];
}

// Helper function to validate argument counts
export function validateArgCount(
  funcName: string,
  args: CellValue[],
  minArgs?: number,
  maxArgs?: number
): void {
  if (minArgs !== undefined && args.length < minArgs) {
    throw new Error(`#VALUE!`);
  }
  if (maxArgs !== undefined && args.length > maxArgs) {
    throw new Error(`#VALUE!`);
  }
}
