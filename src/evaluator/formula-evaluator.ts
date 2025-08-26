import { parseFormula } from "src/parser/parser";
import { SheetHandler } from "../core/sheet-handler";
import {
  FormulaError,
  type CellAddress,
  type CellNumber,
  type CellValue,
  type FunctionEvaluationResult,
  type SpilledValue,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
} from "../core/types";
import { dependencyNodeToKey } from "../core/utils/dependency-node-key";

import { type EvaluationContext } from "../core/types";

import {
  evaluateScalarOperator,
  type EvaluateScalarOperatorOptions,
} from "src/evaluator/evaluate-scalar-operator";
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
import { getCellReference } from "../core/utils";
import { add } from "./arithmetic/add/add";
import { subtract } from "./arithmetic/subtract/subtract";
import { multiply } from "./arithmetic/multiply/multiply";
import { divide } from "./arithmetic/divide/divide";
import { power } from "./arithmetic/power/power";

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

export class FormulaEvaluator extends SheetHandler {
  evaluatedNodes: Map<
    /**
     * key is the dependency node key, from dependencyNodeToKey
     */
    string,
    {
      /**
       * deps is the set of dependency node keys
       */
      deps: Set<string>;
      /**
       * evaluationResult is the evaluation result
       */
      evaluationResult?: FunctionEvaluationResult;
    }
  > = new Map();

  spilledValues: Map<
    /**
     * key is the dependency node key, from dependencyNodeToKey for the origin cell
     */
    string,
    SpilledValue
  > = new Map();

  isCellInRange(cellAddress: CellAddress, range: SpreadsheetRange) {
    const endCol = range.end.col;
    const endRow = range.end.row;
    if (endCol.type === "number" && endRow.type === "number") {
      return (
        cellAddress.colIndex >= range.start.col &&
        cellAddress.colIndex <= endCol.value &&
        cellAddress.rowIndex >= range.start.row &&
        cellAddress.rowIndex <= endRow.value
      );
    } else if (endCol.type === "infinity" && endRow.type === "number") {
      return (
        cellAddress.colIndex >= range.start.col &&
        cellAddress.rowIndex <= endRow.value
      );
    } else if (endCol.type === "number" && endRow.type === "infinity") {
      return (
        cellAddress.rowIndex >= range.start.row &&
        cellAddress.colIndex <= endCol.value
      );
    } else if (endCol.type === "infinity" && endRow.type === "infinity") {
      return (
        cellAddress.colIndex >= range.start.col &&
        cellAddress.rowIndex >= range.start.row
      );
    }
    return false;
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    const { rowIndex, colIndex } = cellAddress;

    // Get all tables for this sheet

    for (const table of this.tables.values()) {
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

  isSpilled(cellAddress: CellAddress): SpilledValue | undefined {
    for (const spilledValue of this.spilledValues.values()) {
      if (
        spilledValue.origin.sheetName === cellAddress.sheetName &&
        spilledValue.origin.colIndex === cellAddress.colIndex &&
        spilledValue.origin.rowIndex === cellAddress.rowIndex
      ) {
        return undefined;
      }
      if (this.isCellInRange(cellAddress, spilledValue.spillOnto)) {
        return spilledValue;
      }
    }
    return undefined;
  }

  evaluateSpilled(
    cellAddress: CellAddress,
    context: EvaluationContext
  ):
    | { isSpilled: true; result: FunctionEvaluationResult | undefined }
    | { isSpilled: false } {
    const spilled = this.isSpilled(cellAddress);
    if (spilled) {
      const spillSource = this.getSpilledAddress(cellAddress, spilled);
      const spillOrigin = this.runtimeSafeEvaluatedNode(
        spilled.origin,
        context
      );
      if (spillOrigin && spillOrigin.type === "spilled-values") {
        return {
          isSpilled: true,
          result: spillOrigin.evaluate(spillSource, context),
        };
      }
    }
    return { isSpilled: false };
  }

  getSpilledAddress(
    cellAddress: CellAddress,
    /**
     * if the spilled value is already available, we can use it to get the source address
     */
    passedSpilledValue?: SpilledValue
  ): { address: CellAddress; spillOffset: { x: number; y: number } } {
    const spilledValue = passedSpilledValue ?? this.isSpilled(cellAddress);
    if (!spilledValue) {
      throw new Error("Cell is not spilled");
    }
    const offsetLeft = cellAddress.colIndex - spilledValue.origin.colIndex;
    const offsetTop = cellAddress.rowIndex - spilledValue.origin.rowIndex;
    const address: CellAddress = {
      ...cellAddress,
      colIndex: spilledValue.origin.colIndex + offsetLeft,
      rowIndex: spilledValue.origin.rowIndex + offsetTop,
    };
    if (offsetLeft === 0 && offsetTop === 0) {
      throw new Error(
        "Spilled value is the same as the cell address! The origin has a pre-calculated value that can be used"
      );
    }
    return { address, spillOffset: { x: offsetLeft, y: offsetTop } };
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
    if (node.tableName) {
      table = this.tables.get(node.tableName);
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

      const originCellValue = this.runtimeSafeEvaluatedNode(
        {
          colIndex: range.start.col,
          rowIndex: range.start.row,
          sheetName: table.sheetName,
        },
        context
      );
      if (!originCellValue) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: `Origin cell not found`,
        };
      }

      if (originCellValue.type === "error") {
        return originCellValue;
      }

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
      spillArea: {
        start: {
          col: context.currentCell.colIndex,
          row: context.currentCell.rowIndex,
        },
        end: {
          col: { type: "number", value: context.currentCell.colIndex },
          row: {
            type: "number",
            value: context.currentCell.rowIndex + numSheets - 1,
          },
        },
      },
      spillOrigin: context.currentCell,
      source: `range`,
      originResult:
        originResult.type === "value"
          ? originResult.result
          : originResult.originResult,
      evaluate: (spilledCell, context) => {
        const colIndex = range.start.col + spilledCell.spillOffset.x;
        const rowIndex = range.start.row + spilledCell.spillOffset.y;
        const sheetName = node.sheetName ?? context.currentSheet;
        return this.runtimeSafeEvaluatedNode(
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
    const range = node.range;

    if (this.isRangeOneCell(range)) {
      return this.evaluateReference(
        {
          type: "reference",
          address: {
            colIndex: range.start.col,
            rowIndex: range.start.row,
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

    const originSheetName = node.sheetName ?? context.currentSheet;
    const originCellAddress: CellAddress = {
      sheetName: originSheetName,
      colIndex: range.start.col,
      rowIndex: range.start.row,
    };
    const originResult = this.runtimeSafeEvaluatedNode(
      originCellAddress,
      context
    );
    if (!originResult) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: "Range is empty",
      };
    }
    if (originResult.type === "error") {
      return originResult;
    }
    return {
      type: "spilled-values",
      spillArea: this.projectRange(range, context.currentCell),
      spillOrigin: context.currentCell,
      source: `range`,
      originResult:
        originResult.type === "value"
          ? originResult.result
          : originResult.originResult,
      evaluate: (spilledCell, context) => {
        const colIndex = range.start.col + spilledCell.spillOffset.x;
        const rowIndex = range.start.row + spilledCell.spillOffset.y;
        const sheetName = node.sheetName ?? context.currentSheet;
        return this.runtimeSafeEvaluatedNode(
          {
            colIndex,
            rowIndex,
            sheetName,
          },
          context
        );
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
      spillArea: {
        start: {
          col: context.currentCell.colIndex,
          row: context.currentCell.rowIndex,
        },
        end: {
          col: {
            type: "number",
            value: context.currentCell.colIndex + firstRow.length - 1,
          },
          row: {
            type: "number",
            value: context.currentCell.rowIndex + node.elements.length - 1,
          },
        },
      },
      spillOrigin: context.currentCell,
      source: `array`,
      originResult:
        originResult.type === "value"
          ? originResult.result
          : originResult.originResult,
      evaluate: (spilledCell, context) => {
        const row = node.elements[spilledCell.spillOffset.y];
        if (!row) {
          return {
            type: "error",
            err: FormulaError.REF,
            message: "Array is empty",
          };
        }
        const cell = row[spilledCell.spillOffset.x];
        if (!cell) {
          return {
            type: "error",
            err: FormulaError.REF,
            message: "Array is empty",
          };
        }
        return this.evaluateNode(cell, context);
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
      // Apply unary operation to each spilled value
      const originResult = this.evaluateUnaryScalar(
        node.operator,
        operandResult.originResult
      );

      if (originResult.type === "error") {
        return originResult;
      }

      if (originResult.type !== "value") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Invalid origin result for unary operation",
        };
      }

      return {
        type: "spilled-values",
        spillArea: operandResult.spillArea,
        spillOrigin: context.currentCell,
        source: `unary ${node.operator} operation`,
        originResult: originResult.result,
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
  ): FunctionEvaluationResult {
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

  getCellEvaluationResult(
    cellAddress: CellAddress
  ): FunctionEvaluationResult | undefined {
    throw new Error("WIP: unimplemented");
  }

  runtimeSafeEvaluatedNode(
    cellAddress: CellAddress,
    context: EvaluationContext
  ): FunctionEvaluationResult | undefined {
    const spilled = this.isSpilled(cellAddress);
    if (spilled) {
      const spillSource = this.getSpilledAddress(cellAddress, spilled);
      const spillOrigin = this.runtimeSafeEvaluatedNode(
        spilled.origin,
        context
      );
      if (spillOrigin && spillOrigin.type === "spilled-values") {
        return spillOrigin.evaluate(spillSource, context);
      }
    }
    const key = dependencyNodeToKey({
      type: "cell",
      address: cellAddress,
      sheetName: cellAddress.sheetName,
    });
    context.dependencies.add(key);
    return this.evaluatedNodes.get(key)?.evaluationResult;
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
    };
    const result = this.runtimeSafeEvaluatedNode(cellAddress, context);
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
    const nodeKey = dependencyNodeToKey({
      type: "named-expression",
      name: node.name,
      sheetName: node.sheetName ?? context.currentSheet,
    });
    context.dependencies.add(nodeKey);

    const value = this.evaluatedNodes.get(nodeKey);
    if (!value || !value.evaluationResult) {
      throw new Error(FormulaError.REF);
    }

    return value.evaluationResult;
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

      // case "&":
      //   return this.concatenate(left, right);
      // case "=":
      //   return this.equals(left, right);
      // case "<>":
      //   return this.notEquals(left, right);
      // case "<":
      //   return this.lessThan(left, right);
      // case "<=":
      //   return this.lessThanOrEqual(left, right);
      // case ">":
      //   return this.greaterThan(left, right);
      // case ">=":
      //   return this.greaterThanOrEqual(left, right);
      default:
        throw new Error(FormulaError.ERROR);
    }
  }

  evaluateFunction(
    node: FunctionNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    throw new Error("Must be implemented by the consuming engine");
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

  projectRange(
    range: SpreadsheetRange,
    originCellAddress: CellAddress
  ): SpreadsheetRange {
    const sheet = this.sheets.get(originCellAddress.sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }
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
