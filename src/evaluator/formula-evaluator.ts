import { parseFormula } from "src/parser/parser";
import type { SingleEvaluationResult } from "../core/types";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type FunctionEvaluationResult,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
} from "../core/types";

import { type EvaluationContext } from "../core/types";

import type { StoreManager } from "src/core/managers/store-manager";
import type { TableManager } from "src/core/managers/table-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import {
  evaluateScalarOperator,
  type EvaluateScalarOperatorOptions,
} from "src/evaluator/evaluate-scalar-operator";
import { functions } from "src/functions";
import {
  getRangeIntersection,
  OpenRangeEvaluator,
} from "src/functions/math/open-range-evaluator";
import type {
  ArrayNode,
  ASTNode,
  BinaryOpNode,
  FunctionNode,
  NamedExpressionNode,
  RangeNode,
  ReferenceNode,
  StructuredReferenceNode,
  ThreeDRangeNode,
  UnaryOpNode,
  ValueNode,
} from "src/parser/ast";
import { getCellReference, isRangeOneCell } from "../core/utils";
import { add } from "./arithmetic/add/add";
import { divide } from "./arithmetic/divide/divide";
import { multiply } from "./arithmetic/multiply/multiply";
import { power } from "./arithmetic/power/power";
import { subtract } from "./arithmetic/subtract/subtract";
import { equals } from "./comparison/equals";
import { greaterThan } from "./comparison/greater-than";
import { greaterThanOrEqual } from "./comparison/greater-than-or-equal";
import { lessThan } from "./comparison/less-than";
import { lessThanOrEqual } from "./comparison/less-than-or-equal";
import { notEquals } from "./comparison/not-equals";
import { concatenate } from "./concatenation/concatenate";

function isFormulaError(value: string): value is FormulaError {
  if (typeof value !== "string") return false;

  // Check for all known formula errors
  const errors: FormulaError[] = Object.values(FormulaError);

  return errors.includes(value as FormulaError);
}

/**
 * Maps JavaScript errors to formula errors
 */
function mapJSErrorToFormulaError(error: Error): FormulaError {
  const message = error.message.toLowerCase();

  if (isFormulaError(error.message)) {
    return error.message;
  }

  if (
    message.includes("division by zero") ||
    message.includes("divide by zero")
  ) {
    return FormulaError.DIV0;
  }
  if (message.includes("circular") || message.includes("cycle")) {
    return FormulaError.CYCLE;
  }
  if (
    message.includes("invalid reference") ||
    (message.includes("reference") && !message.includes("circular"))
  ) {
    return FormulaError.REF;
  }
  if (
    message.includes("invalid name") ||
    message.includes("unknown function")
  ) {
    return FormulaError.NAME;
  }
  if (
    message.includes("invalid number") ||
    message.includes("nan") ||
    message.includes("infinity")
  ) {
    return FormulaError.NUM;
  }
  if (message.includes("type") || message.includes("invalid argument")) {
    return FormulaError.VALUE;
  }
  if (message.includes("not available") || message.includes("n/a")) {
    return FormulaError.NA;
  }

  return FormulaError.ERROR;
}

export class FormulaEvaluator {
  private openRangeEvaluator: OpenRangeEvaluator;
  constructor(
    private tableManager: TableManager,
    private storeManager: StoreManager,
    workbookManager: WorkbookManager
  ) {
    this.openRangeEvaluator = new OpenRangeEvaluator(
      storeManager,
      workbookManager,
      this
    );
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    const { rowIndex, colIndex } = cellAddress;

    // Get all tables for this sheet

    for (const table of this.tableManager
      .getTables(cellAddress.workbookName)
      .values()) {
      // Check each table to see if the cell is within its bounds
      if (table.sheetName !== cellAddress.sheetName) {
        continue;
      }

      const { start, endRow, headers } = table;

      // Check row bounds
      const isInRowRange =
        endRow.type === "infinity"
          ? rowIndex >= start.rowIndex
          : rowIndex >= start.rowIndex && rowIndex <= endRow.value;

      // Check column bounds
      const endColIndex = start.colIndex + headers.size - 1;
      const isInColRange =
        colIndex >= start.colIndex && colIndex <= endColIndex;

      if (isInRowRange && isInColRange) {
        return table;
      }
    }

    return undefined;
  }

  // evaluator methods
  evaluateFormula(
    /**
     * formula is the formula to evaluate, without the leading =
     */
    formula: string,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const ast = parseFormula(formula);

    try {
      const result = this.evaluateNode(ast, context);

      return result;
    } catch (error) {
      // Convert JavaScript errors to formula errors
      const formulaError =
        error instanceof Error
          ? mapJSErrorToFormulaError(error)
          : FormulaError.ERROR;

      return {
        type: "error",
        err: formulaError,
        message: (error as any)?.stack || "An error was thrown",
      };
    }
  }

  evaluateNode(
    node: ASTNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    switch (node.type) {
      case "value":
        return {
          type: "value",
          result: this.evaluateValue(node),
        };
      case "infinity":
        return {
          type: "value",
          result: {
            type: "infinity",
            sign: "positive",
          },
        };
      case "binary-op":
        return this.evaluateBinaryOp(node, context);

      case "reference":
        return this.evaluateReference(node, context);

      case "named-expression":
        return this.evaluateNamedExpression(node, context);

      case "structured-reference":
        return this.evaluateStructuredReference(node, context);

      case "function":
        return this.evaluateFunction(node, context);

      case "range":
        return this.evaluateRange(node, context);

      case "unary-op":
        return this.evaluateUnaryOp(node, context);

      case "3d-range":
        return this.evaluate3DRange(node, context);

      case "array":
        return this.evaluateArray(node, context);

      default:
        return {
          type: "error",
          err: FormulaError.ERROR,
          message: "WIP: unimplemented support for " + node.type,
        };
    }
  }

  evaluateStructuredReference(
    node: StructuredReferenceNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    let table: TableDefinition | undefined;
    if (node.tableName && node.workbookName) {
      table = this.tableManager
        .getTables(node.workbookName)
        .get(node.tableName);
    } else if (node.tableName) {
      table = this.tableManager
        .getTables(context.currentWorkbook)
        .get(node.tableName);
    } else {
      table = this.isCellInTable(context.currentCell);
    }
    if (!table) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: `Table ${node.tableName} not found`,
      };
    }

    const rowIndex = context.currentCell.rowIndex;
    const tableStart = table.start;
    if (node.cols) {
      const startCol = table.headers.get(node.cols.startCol);
      const endCol = table.headers.get(node.cols.endCol);
      if (!startCol || !endCol) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: `Column ${node.cols.startCol} or ${node.cols.endCol} not found in table ${node.tableName}`,
        };
      }
      const startColIndex = tableStart.colIndex + startCol.index;
      const endColIndex = tableStart.colIndex + endCol.index;
      const range: SpreadsheetRange = {
        start: {
          row: node.isCurrentRow ? rowIndex : table.start.rowIndex + 1,
          col: startColIndex,
        },
        end: {
          row: node.isCurrentRow
            ? { type: "number", value: rowIndex }
            : table.endRow,
          col: { type: "number", value: endColIndex },
        },
      };

      return this.evaluateRange(
        {
          type: "range",
          range,
          isAbsolute: {
            start: {
              col: true,
              row: true,
            },
            end: {
              col: true,
              row: true,
            },
          },
          sheetName: table.sheetName,
        },
        context
      );
    }
    return {
      type: "error",
      err: FormulaError.REF,
      message: "WIP: unimplemented support for structured reference",
    };
  }

  /**
   * Evaluates a value node
   */
  evaluateValue(node: ValueNode): CellValue {
    return node.value;
  }

  evaluate3DRange(
    node: ThreeDRangeNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    throw new Error("WIP: 3d range is not implemented");
    /*
    const startSheet = this.sheets.get(node.startSheet);
    const endSheet = this.sheets.get(node.endSheet);
    if (!startSheet || !endSheet) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: `Sheet ${node.startSheet} or ${node.endSheet} not found`,
      };
    }

    let numCols = 0;
    let numRows = 0;
    for (let i = startSheet.index; i <= endSheet.index; i++) {
      if (node.reference.type === "reference") {
        numCols += 1;
      } else {
        numCols += node.reference.range.end.col.value - node.reference.range.start.col.value + 1;
      }

      return {
        type: "error",
        err: FormulaError.REF,
        message: `Sheet ${i} not found`,
      };
    }

    return {
      type: "spilled-values",
      spillArea: (origin: CellAddress) => ({
        start: {
          col: origin.colIndex,
          row: origin.rowIndex,
        },
        end: {
          col: { type: "number", value: origin.colIndex },
          row: {
            type: "number",
            value: origin.rowIndex + numSheets - 1,
          },
        },
      }),
      source: `range`,
      originResult:
        originResult.type === "value"
          ? originResult.result
          : originResult.originResult,
      evaluate: (spilledCell, context) => {
        const colIndex = range.start.col + spilledCell.spillOffset.x;
        const rowIndex = range.start.row + spilledCell.spillOffset.y;
        const sheetName = node.sheetName ?? context.currentSheet;
        return this.evalTimeSafeEvaluateCell(
          {
            colIndex,
            rowIndex,
            sheetName,
          },
          context
        );
      },
    };
    */
  }

  evaluateRange(
    node: RangeNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    if (isRangeOneCell(node.range)) {
      return this.evaluateReference(
        {
          type: "reference",
          address: {
            colIndex: node.range.start.col,
            rowIndex: node.range.start.row,
          },
          isAbsolute: {
            col: node.isAbsolute.start.col || node.isAbsolute.end.col,
            row: node.isAbsolute.start.row || node.isAbsolute.end.row,
          },
          sheetName: node.sheetName,
        },
        context
      );
    }

    return {
      type: "spilled-values",
      spillArea: (origin) => this.projectRange(node.range, origin),
      source: `range`,
      evaluate: (spillOffset, context) => {
        const originSheetName = node.sheetName ?? context.currentSheet;
        const originWorkbookName = node.workbookName ?? context.currentWorkbook;
        const colIndex = node.range.start.col + spillOffset.x;
        const rowIndex = node.range.start.row + spillOffset.y;
        const result = this.storeManager.evalTimeSafeEvaluateCell(
          {
            colIndex,
            rowIndex,
            sheetName: originSheetName,
            workbookName: originWorkbookName,
          },
          context
        );

        if (result) {
          if (result.type === "spilled-values") {
            const originResult = result.evaluate({ x: 0, y: 0 }, context);
            return originResult;
          }
          return result;
        }
      },
      evaluateAllCells: function* ({ evaluate, intersection, context }) {
        let range = node.range;
        if (intersection) {
          const calculateIntersection = getRangeIntersection(
            node.range,
            intersection
          );
          if (calculateIntersection) {
            range = calculateIntersection;
          }
        }

        return yield* this.openRangeEvaluator.evaluateCellsInRange({
          evaluate,
          context,
          origin: {
            range,
            sheetName: node.sheetName ?? context.currentSheet,
            workbookName: node.workbookName ?? context.currentWorkbook,
          },
        });
      },
    };
  }

  evaluateArray(
    node: ArrayNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const firstRow = node.elements[0];
    if (!firstRow) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: "Array is empty",
      };
    }
    const firstCell = firstRow[0];
    if (!firstCell) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: "Array is empty",
      };
    }
    const originResult = this.evaluateNode(firstCell, context);
    if (originResult.type === "error") {
      return originResult;
    }
    return {
      type: "spilled-values",
      spillArea: (origin) => ({
        start: {
          col: origin.colIndex,
          row: origin.rowIndex,
        },
        end: {
          col: {
            type: "number",
            value: origin.colIndex + firstRow.length - 1,
          },
          row: {
            type: "number",
            value: origin.rowIndex + node.elements.length - 1,
          },
        },
      }),
      source: `array`,
      evaluate: (spillOffset, context) => {
        const row = node.elements[spillOffset.y];
        if (!row) {
          return {
            type: "error",
            err: FormulaError.REF,
            message: "Array is empty",
          };
        }
        const cell = row[spillOffset.x];
        if (!cell) {
          return {
            type: "error",
            err: FormulaError.REF,
            message: "Array is empty",
          };
        }
        const result = this.evaluateNode(cell, context);
        if (result.type === "spilled-values") {
          throw new Error("Arrays cannot contain spilled values");
        }
        return result;
      },
      evaluateAllCells: (intersectingRange) => {
        throw new Error("WIP: evaluateAllCells for array is not implemented");
      },
    };
  }

  /**
   * Evaluates a unary operation
   */
  evaluateUnaryOp(
    node: UnaryOpNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const operandResult = this.evaluateNode(node.operand, context);

    if (operandResult.type === "error") {
      return operandResult;
    }

    if (operandResult.type === "spilled-values") {
      return {
        type: "spilled-values",
        spillArea: (origin) => operandResult.spillArea(origin),
        source: `unary ${node.operator} operation`,
        evaluate: (spilledCell, context) => {
          const spillResult = operandResult.evaluate(spilledCell, context);
          if (!spillResult || spillResult.type === "error") {
            return spillResult;
          }
          if (spillResult.type !== "value") {
            return {
              type: "error",
              err: FormulaError.VALUE,
              message: "Invalid spilled result for unary operation",
            };
          }
          return this.evaluateUnaryScalar(node.operator, spillResult.result);
        },
        evaluateAllCells: function* (options) {
          for (const cellValue of operandResult.evaluateAllCells.call(
            this,
            options
          )) {
            if (cellValue.type === "error") {
              yield cellValue;
            } else {
              yield this.evaluateUnaryScalar(node.operator, cellValue.result);
            }
          }
        },
      };
    }

    if (operandResult.type === "value") {
      return this.evaluateUnaryScalar(node.operator, operandResult.result);
    }

    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Invalid operand for unary operation",
    };
  }

  /**
   * Evaluates a unary scalar operation
   */
  private evaluateUnaryScalar(
    operator: "+" | "-" | "%",
    operand: CellValue
  ): SingleEvaluationResult {
    if (operand.type !== "number" && operand.type !== "infinity") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `Cannot apply unary ${operator} to non-number`,
      };
    }
    if (operand.type === "infinity") {
      if (operator === "%") {
        return {
          type: "error",
          err: FormulaError.NUM,
          message: "Cannot apply % to infinity",
        };
      }
      return {
        type: "value",
        result: {
          type: "infinity",
          sign: operator === "+" ? "positive" : "negative",
        },
      };
    }
    switch (operator) {
      case "+":
        return { type: "value", result: operand };

      case "-":
        return {
          type: "value",
          result: {
            type: "number",
            value: -operand.value,
          },
        };

      case "%":
        return {
          type: "value",
          result: { type: "number", value: operand.value / 100 },
        };

      default:
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: `Unknown unary operator: ${operator}`,
        };
    }
  }

  /**
   * Evaluates a binary operation
   */
  evaluateBinaryOp(
    node: BinaryOpNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const left = this.evaluateNode(node.left, context);
    const right = this.evaluateNode(node.right, context);

    if (left.type === "error") {
      return left;
    }
    if (right.type === "error") {
      return right;
    }

    // Scalar operation
    return this.evaluateBinaryScalar(node.operator, left, right, context);
  }

  /**
   * Evaluates a reference node
   */
  evaluateReference(
    node: ReferenceNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const cellAddress: CellAddress = {
      ...node.address,
      sheetName: node.sheetName ?? context.currentSheet,
      workbookName: node.workbookName ?? context.currentWorkbook,
    };
    const result = this.storeManager.evalTimeSafeEvaluateCell(
      cellAddress,
      context
    );
    if (!result) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: `Cell ${getCellReference(cellAddress)} not found`,
      };
    }
    return result;
  }

  /**
   * Evaluates a named expression node
   */
  evaluateNamedExpression(
    node: NamedExpressionNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const result = this.storeManager.evalTimeSafeEvaluateNamedExpression(
      node,
      context
    );
    if (!result) {
      return {
        type: "error",
        err: FormulaError.NAME,
        message: `Named expression ${node.name} not found`,
      };
    }

    return result;
  }

  /**
   * Binary scalar operations
   */
  evaluateBinaryScalar(
    operator: string,
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    switch (operator) {
      case "+":
        return this.add(left, right, context);
      case "-":
        return this.subtract(left, right, context);
      case "*":
        return this.multiply(left, right, context);
      case "/":
        return this.divide(left, right, context);
      case "^":
        return this.power(left, right, context);

      case "&":
        return this.concatenateOp(left, right, context);
      case "=":
        return this.equalsOp(left, right, context);
      case "<>":
        return this.notEqualsOp(left, right, context);
      case "<":
        return this.lessThanOp(left, right, context);
      case "<=":
        return this.lessThanOrEqualOp(left, right, context);
      case ">":
        return this.greaterThanOp(left, right, context);
      case ">=":
        return this.greaterThanOrEqualOp(left, right, context);
      default:
        throw new Error(FormulaError.ERROR);
    }
  }

  evaluateFunction(
    node: FunctionNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const func = functions[node.name];
    if (!func) {
      throw new Error(FormulaError.NAME);
    }
    return func.evaluate.call(this, node, context);
  }

  // Arithmetic operations
  add(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: add,
      context,
      name: "add",
    });
  }

  multiply(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: multiply,
      context,
      name: "multiply",
    });
  }

  divide(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: divide,
      context,
      name: "divide",
    });
  }

  evaluateScalarOperator(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    options: EvaluateScalarOperatorOptions
  ): FunctionEvaluationResult {
    return evaluateScalarOperator.call(this, left, right, options);
  }

  subtract(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: subtract,
      context,
      name: "subtract",
    });
  }

  power(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: power,
      context,
      name: "power",
    });
  }

  // Comparison operations
  equalsOp(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: equals,
      context,
      name: "equals",
    });
  }

  notEqualsOp(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: notEquals,
      context,
      name: "notEquals",
    });
  }

  lessThanOp(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: lessThan,
      context,
      name: "lessThan",
    });
  }

  lessThanOrEqualOp(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: lessThanOrEqual,
      context,
      name: "lessThanOrEqual",
    });
  }

  greaterThanOp(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: greaterThan,
      context,
      name: "greaterThan",
    });
  }

  greaterThanOrEqualOp(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: greaterThanOrEqual,
      context,
      name: "greaterThanOrEqual",
    });
  }

  // Concatenation operation
  concatenateOp(
    left: FunctionEvaluationResult,
    right: FunctionEvaluationResult,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    return this.evaluateScalarOperator(left, right, {
      evaluateScalar: concatenate,
      context,
      name: "concatenate",
    });
  }

  projectRange(
    range: SpreadsheetRange,
    originCellAddress: CellAddress
  ): SpreadsheetRange {
    const offsetLeft = originCellAddress.colIndex - range.start.col;
    const offsetTop = originCellAddress.rowIndex - range.start.row;
    return {
      start: {
        col: range.start.col + offsetLeft,
        row: range.start.row + offsetTop,
      },
      end: {
        col:
          range.end.col.type === "number"
            ? { type: "number", value: range.end.col.value + offsetLeft }
            : range.end.col,
        row:
          range.end.row.type === "number"
            ? { type: "number", value: range.end.row.value + offsetTop }
            : range.end.row,
      },
    };
  }

  unionRanges(
    range1: SpreadsheetRange,
    range2: SpreadsheetRange
  ): SpreadsheetRange {
    const endCol = ((): SpreadsheetRangeEnd => {
      if (
        range1.end.col.type === "infinity" ||
        range2.end.col.type === "infinity"
      ) {
        return { type: "infinity", sign: "positive" };
      }
      return {
        type: "number",
        value: Math.max(range1.end.col.value, range2.end.col.value),
      };
    })();

    const endRow = ((): SpreadsheetRangeEnd => {
      if (
        range1.end.row.type === "infinity" ||
        range2.end.row.type === "infinity"
      ) {
        return { type: "infinity", sign: "positive" };
      }
      return {
        type: "number",
        value: Math.max(range1.end.row.value, range2.end.row.value),
      };
    })();

    return {
      start: {
        col: Math.min(range1.start.col, range2.start.col),
        row: Math.min(range1.start.row, range2.start.row),
      },
      end: {
        col: endCol,
        row: endRow,
      },
    };
  }
}
