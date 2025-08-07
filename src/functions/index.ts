import type { CellValue, FormulaError } from '../core/types';
import type { ASTNode } from '../parser/ast';
import type { EvaluationContext, FunctionDefinition } from '../evaluator/evaluator';
import { basicMathFunctions } from './math/basic';
import { advancedMathFunctions } from './math/advanced';
import { statisticalFunctions } from './math/statistical';
import { lookupFunctions } from './lookup/lookup-functions';
import { logicalConditionFunctions } from './logical/conditions';
import { logicalComparisonFunctions } from './logical/comparisons';
import { infoFunctions } from './info/info-functions';
import { arrayFunctions } from './array/array-functions';
import { textFunctions } from './text/string-functions';

export class FunctionRegistry {
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
    basicMathFunctions.forEach(func => this.register(func));
    // Register advanced math functions
    advancedMathFunctions.forEach(func => this.register(func));
  }

  private registerStatisticalFunctions(): void {
    // Register statistical functions
    statisticalFunctions.forEach(func => this.register(func));
  }

  private registerLogicalFunctions(): void {
    // Register logical condition functions
    logicalConditionFunctions.forEach(func => this.register(func));
    // Register logical comparison functions
    logicalComparisonFunctions.forEach(func => this.register(func));
  }

  private registerTextFunctions(): void {
    // Register text functions
    textFunctions.forEach(func => this.register(func));
  }

  private registerLookupFunctions(): void {
    // Register lookup functions
    lookupFunctions.forEach(func => this.register(func));
  }

  private registerInfoFunctions(): void {
    // Register info functions
    infoFunctions.forEach(func => this.register(func));
  }

  private registerArrayFunctions(): void {
    // Register array functions
    arrayFunctions.forEach(func => this.register(func));
  }
}

// Singleton instance
export const functionRegistry = new FunctionRegistry();

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

// Helper function to coerce values to numbers
export function coerceToNumber(value: CellValue): number {
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value === 'string') {
    const num = parseFloat(value);
    if (isNaN(num)) {
      throw new Error('#VALUE!');
    }
    return num;
  }
  if (typeof value === 'boolean') {
    return value ? 1 : 0;
  }
  if (value === undefined || value === null) {
    return 0;
  }
  throw new Error('#VALUE!');
}

// Helper function to coerce values to strings
export function coerceToString(value: CellValue): string {
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'number') {
    return value.toString();
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  if (value === undefined || value === null) {
    return '';
  }
  if (isFormulaError(value)) {
    return value;
  }
  return String(value);
}

// Helper function to check if value is a formula error
export function isFormulaError(value: CellValue): value is FormulaError {
  return typeof value === 'string' && value.startsWith('#') && value.endsWith('!');
}

// Helper function to propagate errors
export function propagateError(args: CellValue[]): FormulaError | undefined {
  for (const arg of args) {
    if (isFormulaError(arg)) {
      return arg;
    }
  }
  return undefined;
}
