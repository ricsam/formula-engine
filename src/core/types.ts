/**
 * Core type definitions for FormulaEngine
 * This file contains all fundamental types used throughout the engine
 */

import type { ASTNode, FunctionNode } from "src/parser/ast";
import type { FormulaEngine } from "./engine";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

// Cell addressing types
export interface CellAddress {
  sheetName: string;
  workbookName: string;
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

export type FiniteSpreadsheetRange = {
  start: {
    col: number;
    row: number;
  };
  end: {
    col: number;
    row: number;
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
  workbookName?: string;
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
  workbookName: string;
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

export interface Workbook {
  name: string;
  sheets: Map<string, Sheet>;
}

// Event types
export interface FormulaEngineEvents {
  "workbook-added": { workbookName: string };
  "workbook-removed": { workbookName: string };
  "workbook-renamed": { oldName: string; newName: string };
  "sheet-added": {
    sheetName: string;
    workbookName: string;
  };
  "sheet-removed": {
    sheetName: string;
    workbookName: string;
  };
  "sheet-renamed": {
    oldSheetName: string;
    newSheetName: string;
    workbookName: string;
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
      workbookName: string;
    }
  | {
      type: "range";
      range: SpreadsheetRange;
      sheetName: string;
      workbookName: string;
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
      workbookName: string;
    }
  | {
      type: "table";
      tableName: string;
      sheetName: string;
      workbookName: string;
      area:
        | { kind: "Headers" | "All" | "AllData" }
        | { kind: "Data"; columns: string[]; isCurrentRow: boolean };
    };

/**
 * Evaluation context containing necessary information
 */
export interface EvaluationContext {
  currentSheet: string;
  currentWorkbook: string;
  currentCell: CellAddress;
  evaluationStack: Set<string>; // For cycle detection
  dependencies: Set<string>;
  /**
   * candidates for frontier dependencies that are in the intersection of the spilled range and the target range
   */
  frontierDependencies: Set<string>;
  /**
   * Frontier dependency candidates that were discarded because they are not in the intersection of the spilled range and the target range
   */
  discardedFrontierDependencies: Set<string>;
}

export type EvaluatedDependencyNode = {
  /**
   * deps is the set of dependency node keys
   */
  deps?: Set<string>;
  /**
   * frontierDependencies is the set of dependency node keys that are frontier dependencies
   */
  frontierDependencies?: Set<string>;
  /**
   * discardedFrontierDependencies is the set of dependency node keys that were discarded as frontier dependencies
   */
  discardedFrontierDependencies?: Set<string>;
  /**
   * evaluationResult is the evaluation result
   */
  evaluationResult?: FunctionEvaluationResult;
};

export type ValueEvaluationResult = {
  type: "value";
  result: CellValue;
};

export type ErrorEvaluationResult = {
  type: "error";
  err: FormulaError;
  message: string;
};

export type SingleEvaluationResult =
  | ValueEvaluationResult
  | ErrorEvaluationResult;

export type SpilledValuesEvaluator = (
  spillOffset: { x: number; y: number },
  context: EvaluationContext
) => SingleEvaluationResult | undefined;

export type SpilledValuesEvaluationResult = {
  type: "spilled-values";
  spillArea: (origin: CellAddress) => SpreadsheetRange;
  /**
   * for debugging we add a source string to denote where the spilled values were created
   */
  source: string;
  evaluate: SpilledValuesEvaluator;
  /**
   * evaluateAllCells is a generator function that evaluates all non-empty cells in the spilled range.
   * Because a spilled range can be open-ended, we need to have logic for which cells we should evaluate.
   * e.g. when evaluating a range such as D:D only the cells in the current sheet residing in
   * column D should be evaluated and cells producing spilled values that spill onto D:D.
   *
   * In order to evaluate spilled cells in D:D the range evaluateAllCells need to get all cells in the
   * the intersection of the spilled range and D:D, for that reason evaluateAllCells gets an intersection parameter.
   *
   * #### Producers:
   * In e.g. SEQUENCE and evaluateRange we have logic for which cells in a spilled range we should evaluate,
   *
   * #### Nesting:
   * e.g. evaluation of scalar operators where we want to nest e.g. `5 * right.evaluate()`
   * can be implemented by calling
   * ```ts
   * for (const val of child.evaluateAllCells.call(this, options))
   *   yield 5 * val;
   * ```
   *
   * #### Consumers:
   * Only functions that need access to all spilled values in a range end up calling evaluateAllCells, e.g.
   * SUM, MIN, MAX, MATCH. Other types of functions like INDEX doesn't need to evaluate all cells in a range,
   * but does a lookup into a spilled range using the evaluate method.
   *
   */
  evaluateAllCells: (
    this: FormulaEvaluator,
    options: {
      intersection?: SpreadsheetRange;
      evaluate: SpilledValuesEvaluator;
      context: EvaluationContext;
      origin: CellAddress;
    }
  ) => IterableIterator<
    SingleEvaluationResult,
    undefined | void,
    SingleEvaluationResult | undefined
  >;
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
    this: FormulaEvaluator,
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
