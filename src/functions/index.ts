import type { CellValue, FormulaError } from "../core/types";
import type { ASTNode } from "../parser/ast";
import type {
  EvaluationContext,
  FunctionDefinition,
  EvaluationResult,
} from "../evaluator/evaluator";
import { basicMathFunctions } from "./math/basic";
import { advancedMathFunctions } from "./math/advanced";
import { statisticalFunctions } from "./math/statistical";
import { lookupFunctions } from "./lookup/lookup-functions";
import { logicalConditionFunctions } from "./logical/conditions";
import { logicalComparisonFunctions } from "./logical/comparisons";
import { infoFunctions } from "./info/info-functions";
import { arrayFunctions } from "./array/array-functions";
import { textFunctions } from "./text/string-functions";

// Import and re-export utility functions to maintain backward compatibility
import {
  coerceToNumber,
  coerceToString,
  isFormulaError,
  propagateError,
  propagateErrorFromEvalResults,
  assertScalarArg,
  assertArrayArg,
  getScalarValue,
  getArrayFromEvalResult,
  validateArgCount,
} from "./utils";

// Re-exported for internal use only
// These are available through the imports above



class FunctionRegistry {
  private functions = new Map<string, FunctionDefinition>();

  constructor() {
    this.registerBuiltinFunctions();
  }

  register(definition: FunctionDefinition): void {
    this.functions.set(definition.name.toUpperCase(), definition);
  }

  get(name: string): FunctionDefinition | undefined {
    return this.functions.get(name.toUpperCase());
  }

  has(name: string): boolean {
    return this.functions.has(name.toUpperCase());
  }

  getAllFunctions(): Map<string, FunctionDefinition> {
    return new Map(this.functions);
  }

  private registerBuiltinFunctions(): void {
    // Basic math functions
    this.registerMathFunctions();
    // Statistical functions
    this.registerStatisticalFunctions();
    // Logical functions
    this.registerLogicalFunctions();
    // Text functions
    this.registerTextFunctions();
    // Lookup functions
    this.registerLookupFunctions();
    // Info functions
    this.registerInfoFunctions();
    // Array functions
    this.registerArrayFunctions();
  }

  private registerMathFunctions(): void {
    // Register basic math functions
    basicMathFunctions.forEach((func) => this.register(func));
    // Register advanced math functions
    advancedMathFunctions.forEach((func) => this.register(func));
  }

  private registerStatisticalFunctions(): void {
    // Register statistical functions
    statisticalFunctions.forEach((func) => this.register(func));
  }

  private registerLogicalFunctions(): void {
    // Register logical condition functions
    logicalConditionFunctions.forEach((func) => this.register(func));
    // Register logical comparison functions
    logicalComparisonFunctions.forEach((func) => this.register(func));
  }

  private registerTextFunctions(): void {
    // Register text functions
    textFunctions.forEach((func) => this.register(func));
  }

  private registerLookupFunctions(): void {
    // Register lookup functions
    lookupFunctions.forEach((func) => this.register(func));
  }

  private registerInfoFunctions(): void {
    // Register info functions
    infoFunctions.forEach((func) => this.register(func));
  }

  private registerArrayFunctions(): void {
    // Register array functions
    arrayFunctions.forEach((func) => this.register(func));
  }
}

// Singleton instance
export const functionRegistry = new FunctionRegistry();

// All utility functions are now re-exported from ./utils to avoid circular imports
// This prevents the circular dependency: index.ts -> functions -> index.ts



// Overloaded version of safeGetScalarValue for backward compatibility
function safeGetScalarValue(
  argEvaluatedValues: EvaluationResult[],
  index: number,
  defaultValue?: CellValue
): CellValue;
function safeGetScalarValue(
  evalResult: EvaluationResult
): CellValue;
function safeGetScalarValue(
  arg1: EvaluationResult | EvaluationResult[],
  index?: number,
  defaultValue: CellValue = 0
): CellValue {
  if (Array.isArray(arg1)) {
    const result = arg1[index!];
    if (!result) return defaultValue;
    return getScalarValue(result);
  } else {
    return getScalarValue(arg1);
  }
}
