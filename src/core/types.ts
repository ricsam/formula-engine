/**
 * Core type definitions for FormulaEngine
 * This file contains all fundamental types used throughout the engine
 */

import type { EvaluationContext } from "../evaluator/evaluation-context";
import type { FormulaEvaluator } from "../evaluator/formula-evaluator";
import type { FunctionNode } from "../parser/ast";
import type { DependencyNode } from "./managers/dependency-node";
import type { LookupOrder } from "./managers/range-eval-order-builder";

// Cell addressing types
export interface CellAddress {
  sheetName: string;
  workbookName: string;
  colIndex: number;
  rowIndex: number;
}

export interface RangeAddress {
  sheetName: string;
  workbookName: string;
  range: SpreadsheetRange;
}

export interface LocalCellAddress {
  colIndex: number;
  rowIndex: number;
}

export type ArethmeticEvaluator = (
  left: CellValue,
  right: CellValue,
  context: EvaluationContext
) => CellValue | ErrorEvaluationResult;

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

export type RelativeRange = {
  start: {
    col: number;
    row: number;
  };
  width: SpreadsheetRangeEnd;
  height: SpreadsheetRangeEnd;
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
/**
 * undefined and "" are considered empty values
 * undefineds are converted to "" in the engine
 *
 * any empty values are deleted from the sheet content
 */
export type SerializedCellValue = string | number | boolean | undefined;

// Named expressions
export interface NamedExpression {
  name: string;
  expression: string;
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

export type ValueEvaluationResult = {
  type: "value";
  result: CellValue;
  /**
   * If the terminating evaluation result is a reference (see evaluateReference)
   * then we store information about the source cell for context dependent functions like CELL
   */
  sourceCell?: CellAddress;
};

export type AwaitingEvaluationResult = {
  type: "awaiting-evaluation";
  waitingFor: DependencyNode;
  errAddress: DependencyNode;
  /**
   * If the terminating evaluation result is a reference (see evaluateReference)
   * then we store information about the source cell for context dependent functions like CELL
   */
  sourceCell?: CellAddress;
};

export type DoesNotSpillResult = {
  type: "does-not-spill";
};

export type ErrorEvaluationResult =
  | {
      type: "error";
      err: FormulaError;
      errAddress: DependencyNode;
      message: string;
      /**
       * If the terminating evaluation result is a reference (see evaluateReference)
       * then we store information about the source cell for context dependent functions like CELL
       */
      sourceCell?: CellAddress;
    }
  | AwaitingEvaluationResult;

export type SingleEvaluationResult =
  | ValueEvaluationResult
  | ErrorEvaluationResult;

export type SpilledValuesEvaluator = (
  spillOffset: { x: number; y: number },
  context: EvaluationContext
) => SingleEvaluationResult;

export type SpilledValuesEvaluationResult = {
  type: "spilled-values";

  /**
   * When a raw range is evaluated, we will add it to the sourceRange so it can be used e.g. for context dependent functions
   */
  sourceRange?: RangeAddress;

  /**
   * If the terminating evaluation result is a reference (see evaluateReference)
   * then we store information about the source cell for context dependent functions like CELL
   * sourceCell will only be defined on a spilledValue when a single value is looked up,
   */
  sourceCell?: CellAddress;

  spillArea: (origin: CellAddress) => SpreadsheetRange;
  /**
   * for debugging we add a source string to denote where the spilled values were created
   */
  source: string;
  evaluate: SpilledValuesEvaluator;
  /**
   * evaluateAllCells evaluates all non-empty cells in the spilled range.
   * Because a spilled range can be open-ended, we need to have logic for which cells we should evaluate.
   * e.g. when evaluating a range such as D:D only the cells in the current sheet residing in
   * column D should be evaluated and cells producing spilled values that spill onto D:D.
   *
   * In order to evaluate spilled cells in D:D the range evaluateAllCells need to get all cells in the
   * the intersection of the spilled range and D:D, for that reason evaluateAllCells gets an intersection parameter,
   * where the intersection is relative to the origin.
   *
   * #### Producers:
   * In e.g. SEQUENCE and evaluateRange we have logic for which cells in a spilled range we should evaluate,
   *
   * #### Nesting:
   * e.g. evaluation of scalar operators where we want to nest e.g. `5 * right.evaluate()`
   * can be implemented by calling
   * ```ts
   * const vals = child.evaluateAllCells.call(this, options);
   * return vals.map(val => ({ ...val, result: 5 * val.result }));
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
      /**
       * an intersection relative to the origin
       */
      intersection?: SpreadsheetRange;
      evaluate: SpilledValuesEvaluator;
      context: EvaluationContext;
      /**
       * origin is the cell address that the spilled range is spilled from
       * e.g. in A3=B2:B4 the origin is A3
       */
      origin: CellAddress;

      lookupOrder: LookupOrder;
    }
  ) => EvaluateAllCellsResult;
};

export type EvaluateAllCellsResult =
  | ErrorEvaluationResult
  | {
      type: "values";
      values: CellInRangeResult[];
    };

export type CellInRangeResult = {
  result: SingleEvaluationResult;
  relativePos: { x: number; y: number };
};

export type FunctionEvaluationResult =
  | SingleEvaluationResult
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
  aliases?: string[];
}

/**
 * Evaluation result
 */
export type EvaluationResult = {
  dependencies: Set<string>;
} & FunctionEvaluationResult;

export type SCC = {
  id: number;
  nodes: Set<DependencyNode>; // All nodes considering soft + hard edges
  evaluationOrder: DependencyNode[]; // Flat topologically ordered list
  resolved: boolean;
  hardEdgeSCCs: Set<DependencyNode>[]; // SCCs formed by only hard edges (regular dependencies)
};

export type SCCDAG = {
  sccList: SCC[];
  sccGraph: Map<number, Set<number>>; // Adjacency list of SCC dependencies
};

export type EvaluationOrder = {
  evaluationOrder: Set<DependencyNode>;
  hasCycle: boolean;
  cycleNodes?: Set<DependencyNode>;
  hash: string;
  sccDAG?: SCCDAG;
};

// Conditional Styling types
export interface LCHColor {
  l: number; // Lightness: 0-100
  c: number; // Chroma: 0-150+
  h: number; // Hue: 0-360
}

export interface FormulaStyleCondition {
  type: "formula";
  formula: string;
  color: LCHColor;
}

export interface GradientStyleCondition {
  type: "gradient";
  min:
    | { type: "lowest_value"; color: LCHColor }
    | { type: "number"; color: LCHColor; valueFormula: string };
  max:
    | { type: "highest_value"; color: LCHColor }
    | { type: "number"; color: LCHColor; valueFormula: string };
}

export type StyleCondition = FormulaStyleCondition | GradientStyleCondition;

export interface ConditionalStyle {
  area: RangeAddress;
  condition: StyleCondition;
}

export interface DirectCellStyle {
  area: RangeAddress;
  style: CellStyle
}

export interface CellStyle {
  backgroundColor?: string; // Hex color format
  color?: string; // Text color in hex format
}
