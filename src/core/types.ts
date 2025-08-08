/**
 * Core type definitions for FormulaEngine
 * This file contains all fundamental types used throughout the engine
 */

// Cell addressing types
export interface SimpleCellAddress {
  sheet: number;
  col: number;
  row: number;
}

export interface SimpleCellRange {
  start: SimpleCellAddress;
  end: SimpleCellAddress;
}

// Cell value types
export type CellValue = number | string | boolean | FormulaError | undefined;
export type RawCellContent = CellValue;
export type CellType = "FORMULA" | "VALUE" | "ARRAY" | "EMPTY";
export type CellValueType = "NUMBER" | "STRING" | "BOOLEAN" | "ERROR" | "EMPTY";
export type CellValueDetailedType = CellValueType;

// Format information placeholder
export interface FormatInfo {
  // Placeholder for future formatting support
}

// Bounding rectangle for sheet dimensions
export interface BoundingRect {
  minCol: number;
  maxCol: number;
  minRow: number;
  maxRow: number;
  width: number; // maxCol - minCol + 1
  height: number; // maxRow - minRow + 1
}

// Change tracking for undo/redo and events
export interface ExportedChange {
  address?: SimpleCellAddress;
  oldValue?: CellValue;
  newValue?: CellValue;
  type: "cell-change" | "sheet-change" | "structure-change";
}

// Named expressions
export interface NamedExpression {
  name: string;
  expression: string;
  scope?: number; // undefined = global scope, number = sheet ID
}

export interface SerializedNamedExpression extends NamedExpression {
  id: string;
}

export interface NamedExpressionOptions {
  overwrite?: boolean;
}

// Formula errors
export type FormulaError =
  | "#DIV/0!" // Division by zero
  | "#N/A" // Value not available
  | "#NAME?" // Invalid name/function
  | "#NUM!" // Invalid numeric value
  | "#REF!" // Invalid reference
  | "#VALUE!" // Wrong argument type
  | "#CYCLE!" // Circular reference detected
  | "#ERROR!" // General error
  | "#SPILL!"; // Array spill blocked

// Internal cell structure
export interface Cell {
  value: CellValue;
  formula?: string;
  type: CellType;
  dependencies?: Set<string>;
  dependents?: Set<string>;
  arrayFormula?: ArrayFormula;
}

// Sheet structure
export interface Sheet {
  id: number;
  name: string;
  cells: Map<string, Cell>;
  dimensions?: SheetDimensions;
}

export interface SheetDimensions {
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
}

// Array formula types
export interface ArrayFormula {
  formula: string;
  originAddress: SimpleCellAddress;
  spillRange: SimpleCellRange;
}

// Function definition types
export interface FunctionDefinition {
  name: string;
  evaluate: (...args: any[]) => CellValue;
  validateArgs?: (args: unknown[]) => boolean;
  minArgs?: number;
  maxArgs?: number;
  isVolatile?: boolean; // Functions like NOW(), RAND() that should recalculate on every change
}

// Parser and AST types
export interface ASTNode {
  type: "function" | "reference" | "value" | "operator" | "array" | "range";
  value?: any;
  children?: ASTNode[];
  functionName?: string;
  operator?: string;
  reference?: SimpleCellAddress | SimpleCellRange;
}

// Evaluation context
export interface EvaluationContext {
  currentSheet: number;
  currentAddress: SimpleCellAddress;
  engine: any; // Will be FormulaEngine type when circular deps are resolved
  visitedCells?: Set<string>; // For cycle detection
}

// Command pattern for undo/redo
export interface Command {
  execute(): ExportedChange[];
  undo(): ExportedChange[];
  redo(): ExportedChange[];
  description?: string;
}

export interface CellChangeEvent {
  address: SimpleCellAddress;
  oldValue: CellValue;
  newValue: CellValue;
}

// Preferred name for cell update event in the public API
export type CellUpdateEvent = CellChangeEvent;

// Event types
export interface FormulaEngineEvents {
  "sheet-added": {
    sheetId: number;
    sheetName: string;
  };
  "sheet-removed": {
    sheetId: number;
    sheetName: string;
  };
  "sheet-renamed": {
    sheetId: number;
    oldName: string;
    newName: string;
  };
}

// Utility types
export type AddressKey = string; // Format: "sheet:col:row" e.g., "0:1:2"

// Helper function to create address keys
export function addressToKey(address: SimpleCellAddress): AddressKey {
  return `${address.sheet}:${address.col}:${address.row}`;
}

// Helper function to parse address keys
export function keyToAddress(key: AddressKey): SimpleCellAddress {
  const parts = key.split(":").map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) {
    throw new Error(`Invalid address key: ${key}`);
  }
  const [sheet, col, row] = parts;
  return { sheet: sheet!, col: col!, row: row! };
}

// A1 notation helpers
export function colNumberToLetter(col: number): string {
  let result = "";
  let n = col;
  while (n >= 0) {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

export function letterToColNumber(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result - 1;
}

// Check if a value is a formula error
export function isFormulaError(value: CellValue): value is FormulaError {
  return (
    typeof value === "string" && value.startsWith("#") && value.endsWith("!")
  );
}

// Check if a cell is empty
export function isCellEmpty(value: CellValue): value is undefined {
  return value === undefined;
}

// Type guards for cell value types
export function isNumber(value: CellValue): value is number {
  return typeof value === "number" && !isNaN(value);
}

export function isString(value: CellValue): value is string {
  return typeof value === "string" && !isFormulaError(value);
}

export function isBoolean(value: CellValue): value is boolean {
  return typeof value === "boolean";
}

// Get the type of a cell value
export function getCellValueType(value: CellValue): CellValueType {
  if (value === undefined) return "EMPTY";
  if (isFormulaError(value)) return "ERROR";
  if (isNumber(value)) return "NUMBER";
  if (isBoolean(value)) return "BOOLEAN";
  if (isString(value)) return "STRING";
  return "EMPTY";
}

// Result type for error handling
export type Result<T, E = FormulaError> =
  | { success: true; data: T }
  | { success: false; error: E };

// Configuration options
export interface FormulaEngineOptions {
  evaluationMode?: "lazy" | "eager";
  maxIterations?: number;
  cacheSize?: number;
  enableArrayFormulas?: boolean;
  enableNamedExpressions?: boolean;
  locale?: string;
}

// Export everything for easy access
export * from "./types";
