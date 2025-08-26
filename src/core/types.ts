/**
 * Core type definitions for FormulaEngine
 * This file contains all fundamental types used throughout the engine
 */

import type { ASTNode, FunctionNode } from "src/parser/ast";
import type { FormulaEngine } from "./engine";

// Cell addressing types
export interface CellAddress {
  sheetName: string;
  colIndex: number;
  rowIndex: number;
}

export interface LocalCellAddress {
  colIndex: number;
  rowIndex: number;
}

export type ArethmeticEvaluator = (
  left: CellValue,
  right: CellValue
) => CellValue | { type: "error"; err: FormulaError; message: string };

export type PositiveInfinity = {
  type: "infinity";
  sign: "positive";
};

export type CellInfinity = {
  type: "infinity";
  sign: "positive" | "negative";
};

export type CellNumber = {
  type: "number";
  value: number;
};

export type SpreadsheetRangeEnd = CellNumber | PositiveInfinity;

export type SpreadsheetRange = {
  start: {
    col: number;
    row: number;
  };
  end: {
    col: SpreadsheetRangeEnd;
    row: SpreadsheetRangeEnd;
  };
};

export type CellString = {
  type: "string";
  value: string;
};

export type CellBoolean = {
  type: "boolean";
  value: boolean;
};

// Cell value types
export type CellValue = CellNumber | CellString | CellBoolean | CellInfinity;
export type SerializedCellValue = string | number | boolean | undefined;

// Named expressions
export interface NamedExpression {
  name: string;
  expression: string;
  sheetName?: string;
}

export interface TableDefinition {
  name: string;
  start: {
    rowIndex: number;
    colIndex: number;
  };
  headers: Map<string, { name: string; index: number }>;
  endRow: SpreadsheetRangeEnd;
  sheetName: string;
}

// Formula errors
export enum FormulaError {
  DIV0 = "#DIV/0!",
  NA = "#N/A",
  NAME = "#NAME?",
  NUM = "#NUM!",
  REF = "#REF!",
  VALUE = "#VALUE!",
  CYCLE = "#CYCLE!",
  ERROR = "#ERROR!",
  SPILL = "#SPILL!",
}

// Sheet structure
export interface Sheet {
  name: string;
  index: number; // 0-based index of the sheet
  content: Map<string, SerializedCellValue>;
}

// Event types
export interface FormulaEngineEvents {
  "sheet-added": {
    sheetName: string;
  };
  "sheet-removed": {
    sheetName: string;
  };
  "sheet-renamed": {
    oldName: string;
    newName: string;
  };
  "global-named-expressions-updated": Map<string, NamedExpression>;
  "tables-updated": Map<string, TableDefinition>;
}

/**
 * All dependency nodes are evaluated in the context of the current cell, so therefore it will always have a sheetName
 */
export type DependencyNode =
  | {
      type: "cell";
      address: LocalCellAddress;
      sheetName: string;
    }
  | {
      type: "range";
      range: SpreadsheetRange;
      sheetName: string;
    }
  | {
      type: "multi-spreadsheet-range";
      ranges: SpreadsheetRange;
      sheetNames:
        | { type: "list"; list: string[] }
        | {
            type: "range";
            startSpreadsheetName: string;
            endSpreadsheetName: string;
          };
    }
  | {
      type: "named-expression";
      name: string;
      sheetName: string;
    }
  | {
      type: "table";
      tableName: string;
      sheetName: string;
      area:
        | { kind: "Headers" | "All" | "AllData" }
        | { kind: "Data"; columns: string[]; isCurrentRow: boolean };
    };

/**
 * Evaluation context containing necessary information
 */
export interface EvaluationContext {
  currentSheet: string;
  currentCell: CellAddress;
  evaluationStack: Set<string>; // For cycle detection
  dependencies: Set<string>;
}

export type ValueEvaluationResult = {
  type: "value";
  result: CellValue;
};

export type ErrorEvaluationResult = {
  type: "error";
  err: FormulaError;
  message: string;
};

export type SpilledValuesEvaluationResult = {
  type: "spilled-values";
  spillOrigin: CellAddress;
  spillArea: SpreadsheetRange;
  originResult: CellValue;
  /**
   * for debugging we add a source string to denote where the spilled values were created
   */
  source: string;
  evaluate: (
    spilledCell: {
      address: CellAddress;
      spillOffset: { x: number; y: number };
    },
    context: EvaluationContext
  ) => FunctionEvaluationResult | undefined;
};

export type FunctionEvaluationResult =
  | ValueEvaluationResult
  | ErrorEvaluationResult
  | SpilledValuesEvaluationResult;

export type SpilledValue = {
  /**
   * spillOnto is the range that the spilled value is spilled onto
   */
  spillOnto: SpreadsheetRange;
  /**
   * origin is the cell address that the spilled value is spilled from
   */
  origin: CellAddress;
};

/**
 * Function definition
 */
export interface FunctionDefinition {
  name: string;
  evaluate: (
    this: FormulaEngine,
    node: FunctionNode,
    context: EvaluationContext
  ) => FunctionEvaluationResult;
}

/**
 * Evaluation result
 */
export type EvaluationResult = {
  dependencies: Set<string>;
} & FunctionEvaluationResult;
