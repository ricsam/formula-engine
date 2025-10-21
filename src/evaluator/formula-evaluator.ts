import { parseFormula } from "src/parser/parser";
import type {
  LocalCellAddress,
  RangeAddress,
  SingleEvaluationResult,
} from "../core/types";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type FunctionEvaluationResult,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
} from "../core/types";

import type { DependencyManager } from "src/core/managers/dependency-manager";
import type { TableManager } from "src/core/managers/table-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import {
  evaluateScalarOperator,
  type EvaluateScalarOperatorOptions,
} from "src/evaluator/evaluate-scalar-operator";
import { functions } from "src/functions";
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
import {
  captureEvaluationErrors,
  cellAddressToKey,
  getAbsoluteRange,
  getCellReference,
  getRangeIntersection,
  getRangeKey,
  getRelativeRange,
  getRelativeRangeKey,
  isRangeOneCell,
  rangeAddressToKey,
} from "../core/utils";
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
import type { NamedExpressionManager } from "src/core/managers/named-expression-manager";
import { EvaluationContext } from "./evaluation-context";
import { flags } from "src/debug/flags";
import { AwaitingEvaluationError, EvaluationError } from "./evaluation-error";
import { formatFormula } from "src/parser/formatter";
import { CellValueNode } from "./dependency-nodes/cell-value-node";

export class FormulaEvaluator {
  constructor(
    private tableManager: TableManager,
    private dependencyManager: DependencyManager,
    private namedExpressionManager: NamedExpressionManager
  ) {}

  // evaluator methods
  evaluateFormula(
    /**
     * formula is the formula to evaluate, without the leading =
     */
    formula: string,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const ast = parseFormula(formula);

    return captureEvaluationErrors(context.dependencyNode, () => {
      const result = this.evaluateNode(ast, context);
      return result;
    });
  }

  evaluateNode(
    node: ASTNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const currentContext = {
      ...context.cellAddress,
      tableName: context.tableName,
    };

    const astNode = this.dependencyManager.getAstNode(node, currentContext);
    context.dependencyNode.addDependency(astNode);

    if (astNode.resolved) {
      const astContextDependency = astNode.getContextDependency();
      context.appendContextDependency(astContextDependency);
      return astNode.evaluationResult;
    }

    astNode.resetDirectDepsUpdated();

    function runEvaluation(
      this: FormulaEvaluator,
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
            errAddress: context.dependencyNode,
          };
      }
    }
    const newContext = new EvaluationContext(
      this.tableManager,
      astNode,
      context.cellAddress
    );
    const result = runEvaluation.call(this, newContext);
    astNode.setEvaluationResult(result);
    const astContextDependency = newContext.getContextDependency();

    astNode.setContextDependency(astContextDependency);

    context.appendContextDependency(astContextDependency);

    return result;
  }

  evaluateStructuredReference(
    node: StructuredReferenceNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    // the tables are never dependent on the sheet, interesetingly enough
    let table: TableDefinition | undefined;
    if (node.tableName && node.workbookName) {
      // this expression will evaluate to the same value regardless of where in the workbooks we are evaluating it in
      table = this.tableManager
        .getTables(node.workbookName)
        .get(node.tableName);
    } else if (node.tableName) {
      // different workbooks could have different tables with the same name
      // so it can differ based on the workbook we are evaluating it in
      context.addContextDependency("workbook");
      table = this.tableManager
        .getTables(context.cellAddress.workbookName)
        .get(node.tableName);
    } else {
      // if no table nor workbook name is provided, we need to find the table in the current workbook
      // and the formula will be evaluated differently based on the table (and workbook) we are evaluating it in
      context.addContextDependency("workbook", "table");
      table = this.tableManager.isCellInTable(context.cellAddress);
    }

    if (node.isCurrentRow) {
      context.addContextDependency("row");
    }

    if (!table) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: `Table ${node.tableName} not found`,
        errAddress: context.dependencyNode,
      };
    }

    const rowIndex = context.cellAddress.rowIndex;
    const tableStart = table.start;

    // Handle selector-based references
    if (node.selector) {
      let startRow: number;
      let endRow: SpreadsheetRangeEnd;

      switch (node.selector) {
        case "#Headers":
          startRow = table.start.rowIndex;
          endRow = { type: "number", value: table.start.rowIndex };
          break;
        case "#Data":
          startRow = table.start.rowIndex + 1;
          endRow = table.endRow;
          break;
        case "#All":
          startRow = table.start.rowIndex;
          endRow = table.endRow;
          break;
        default:
          return {
            type: "error",
            err: FormulaError.REF,
            message: `Unknown table selector: ${node.selector}`,
            errAddress: context.dependencyNode,
          };
      }

      // If we also have column specification, use those columns
      if (node.cols) {
        const startCol = table.headers.get(node.cols.startCol);
        const endCol = table.headers.get(node.cols.endCol);
        if (!startCol || !endCol) {
          return {
            type: "error",
            err: FormulaError.REF,
            message: `Column ${node.cols.startCol} or ${node.cols.endCol} not found in table ${table.name}`,
            errAddress: context.dependencyNode,
          };
        }
        const startColIndex = tableStart.colIndex + startCol.index;
        const endColIndex = tableStart.colIndex + endCol.index;

        const range: SpreadsheetRange = {
          start: {
            row: startRow,
            col: startColIndex,
          },
          end: {
            row: endRow,
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
      } else {
        // No column specification, return entire row(s) for the selector
        const range: SpreadsheetRange = {
          start: {
            row: startRow,
            col: tableStart.colIndex,
          },
          end: {
            row: endRow,
            col: {
              type: "number",
              value: tableStart.colIndex + table.headers.size - 1,
            },
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
    }

    // Handle column-only references (no selector)
    if (node.cols) {
      const startCol = table.headers.get(node.cols.startCol);
      const endCol = table.headers.get(node.cols.endCol);
      if (!startCol || !endCol) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: `Column ${node.cols.startCol} or ${node.cols.endCol} not found in table ${table.name}`,
          errAddress: context.dependencyNode,
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
      message: "Structured reference must specify either a selector or columns",
      errAddress: context.dependencyNode,
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

    const debugRange = getRangeKey(node.range);

    const rangeAddress: RangeAddress = {
      sheetName: node.sheetName ?? context.cellAddress.sheetName,
      workbookName: node.workbookName ?? context.cellAddress.workbookName,
      range: node.range,
    };

    return {
      type: "spilled-values",
      spillArea: (origin) => this.projectRange(node.range, origin),
      source: `range ${debugRange}`,
      sourceRange: rangeAddress,
      evaluate: (spillOffset, context) => {
        if (!node.sheetName && !node.workbookName) {
          // e.g. the result from A4:A9 will depend in which sheet and workbook we are evaluating it in
          context.addContextDependency("workbook", "sheet");
        } else if (!node.sheetName) {
          // e.g. the result from [Workbook1]A4:A9 will depend in which sheet we are evaluating it in
          context.addContextDependency("sheet");
        } else if (!node.workbookName) {
          // e.g. the result from Sheet1!A4:A9 will depend in which workbook we are evaluating it in
          context.addContextDependency("workbook");
        } else {
          // if we have both sheetName and workbookName, we don't need to add any context dependencies
        }

        const originSheetName = node.sheetName ?? context.cellAddress.sheetName;
        const originWorkbookName =
          node.workbookName ?? context.cellAddress.workbookName;
        const colIndex = node.range.start.col + spillOffset.x;
        const rowIndex = node.range.start.row + spillOffset.y;

        const cellAddress: CellAddress = {
          colIndex,
          rowIndex,
          sheetName: originSheetName,
          workbookName: originWorkbookName,
        };

        const evalNode = this.dependencyManager.getCellValueOrEmptyCellNode(
          cellAddressToKey(cellAddress)
        );
        context.dependencyNode.addDependency(evalNode);

        const result = evalNode.evaluationResult;

        return result;
      },
      evaluateAllCells: function ({
        evaluate,
        intersection,
        context,
        origin,
        lookupOrder,
      }) {
        let range = node.range;
        if (intersection) {
          // When we have an intersection, it's defined relative to where the spilled range
          // will appear (the origin). However, we need to evaluate cells from the source
          // range (node.range). So we must translate the intersection coordinates back
          // to the source range's coordinate system.
          //
          // Example: If source range A1:C3 spills to D5:F7, and we want intersection E6:F7,
          // we need to translate E6:F7 (relative to D5) to B2:C3 (relative to A1).

          // Calculate the offset of the intersection from the spill origin
          const relativeRange = getRelativeRange(intersection, origin);
          const start: LocalCellAddress = {
            colIndex: node.range.start.col,
            rowIndex: node.range.start.row,
          };
          const projectedIntersection = getAbsoluteRange(relativeRange, start);
          const calculateIntersection = getRangeIntersection(
            node.range,
            projectedIntersection
          );
          if (calculateIntersection) {
            range = calculateIntersection;
          }
        }

        const address: RangeAddress = {
          range,
          sheetName: node.sheetName ?? context.cellAddress.sheetName,
          workbookName: node.workbookName ?? context.cellAddress.workbookName,
        };

        const rangeNode = this.dependencyManager.getRangeNode(
          rangeAddressToKey(address)
        );

        context.dependencyNode.addDependency(rangeNode);

        return this.dependencyManager.getRangeNode(rangeAddressToKey(address))
          .result;
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
        errAddress: context.dependencyNode,
      };
    }
    const firstCell = firstRow[0];
    if (!firstCell) {
      return {
        type: "error",
        err: FormulaError.REF,
        message: "Array is empty",
        errAddress: context.dependencyNode,
      };
    }
    const originResult = this.evaluateNode(firstCell, context);
    if (
      originResult.type === "error" ||
      originResult.type === "awaiting-evaluation"
    ) {
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
            errAddress: context.dependencyNode,
          };
        }
        const cell = row[spillOffset.x];
        if (!cell) {
          return {
            type: "error",
            err: FormulaError.REF,
            message: "Array is empty",
            errAddress: context.dependencyNode,
          };
        }
        const result = this.evaluateNode(cell, context);
        if (result.type === "spilled-values") {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Arrays cannot contain spilled values",
            errAddress: context.dependencyNode,
          };
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

    if (
      operandResult.type === "error" ||
      operandResult.type === "awaiting-evaluation"
    ) {
      return operandResult;
    }

    if (operandResult.type === "spilled-values") {
      return {
        type: "spilled-values",
        spillArea: (origin) => operandResult.spillArea(origin),
        source: `unary ${node.operator} operation`,
        evaluate: (spilledCell, context) => {
          const spillResult = operandResult.evaluate(spilledCell, context);
          if (
            !spillResult ||
            spillResult.type === "error" ||
            spillResult.type === "awaiting-evaluation"
          ) {
            return spillResult;
          }
          if (spillResult.type !== "value") {
            return {
              type: "error",
              err: FormulaError.VALUE,
              message: "Invalid spilled result for unary operation",
              errAddress: context.dependencyNode,
            };
          }
          return this.evaluateUnaryScalar(
            node.operator,
            spillResult.result,
            context
          );
        },
        evaluateAllCells: function (options) {
          const cellValues = operandResult.evaluateAllCells.call(this, options);
          if (cellValues.type !== "values") {
            return cellValues;
          }
          return {
            type: "values",
            values: cellValues.values.map((cellValue) => {
              if (
                cellValue.result.type === "error" ||
                cellValue.result.type === "awaiting-evaluation"
              ) {
                return cellValue;
              } else {
                return {
                  result: this.evaluateUnaryScalar(
                    node.operator,
                    cellValue.result.result,
                    context
                  ),
                  relativePos: cellValue.relativePos,
                };
              }
            }),
          };
        },
      };
    }

    if (operandResult.type === "value") {
      return this.evaluateUnaryScalar(
        node.operator,
        operandResult.result,
        context
      );
    }

    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Invalid operand for unary operation",
      errAddress: context.dependencyNode,
    };
  }

  /**
   * Evaluates a unary scalar operation
   */
  private evaluateUnaryScalar(
    operator: "+" | "-" | "%",
    operand: CellValue,
    context: EvaluationContext
  ): SingleEvaluationResult {
    if (operand.type !== "number" && operand.type !== "infinity") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `Cannot apply unary ${operator} to non-number`,
        errAddress: context.dependencyNode,
      };
    }
    if (operand.type === "infinity") {
      if (operator === "%") {
        return {
          type: "error",
          err: FormulaError.NUM,
          message: "Cannot apply % to infinity",
          errAddress: context.dependencyNode,
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
          errAddress: context.dependencyNode,
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

    if (left.type === "error" || left.type === "awaiting-evaluation") {
      return left;
    }
    if (right.type === "error" || right.type === "awaiting-evaluation") {
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
      sheetName: node.sheetName ?? context.cellAddress.sheetName,
      workbookName: node.workbookName ?? context.cellAddress.workbookName,
    };

    const key = cellAddressToKey(cellAddress);
    const evalNode = this.dependencyManager.getCellValueOrEmptyCellNode(key);
    context.dependencyNode.addDependency(evalNode);

    if (!node.sheetName && !node.workbookName) {
      // e.g. the result from A4:A9 will depend in which sheet and workbook we are evaluating it in
      context.addContextDependency("workbook", "sheet");
    } else if (!node.sheetName) {
      // e.g. the result from [Workbook1]A4:A9 will depend in which sheet we are evaluating it in
      context.addContextDependency("sheet");
    } else if (!node.workbookName) {
      // e.g. the result from Sheet1!A4:A9 will depend in which workbook we are evaluating it in
      context.addContextDependency("workbook");
    } else {
      // if we have both sheetName and workbookName, we don't need to add any context dependencies
    }

    if (evalNode instanceof CellValueNode && evalNode.spillMeta) {
      if (evalNode.spillMeta.evaluationResult.type === "spilled-values") {
        return {
          ...evalNode.spillMeta.evaluationResult,
          sourceCell: cellAddress,
          sourceRange: undefined,
        };
      }
    }

    return { ...evalNode.evaluationResult, sourceCell: cellAddress };
  }

  /**
   * Evaluates a named expression node
   */
  evaluateNamedExpression(
    node: NamedExpressionNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const expression = this.namedExpressionManager.resolveNamedExpression(
      node,
      context
    );

    if (!expression) {
      return {
        type: "error",
        err: FormulaError.NAME,
        message: `Named expression ${node.name} not found`,
        errAddress: context.dependencyNode,
      };
    }

    return this.evaluateFormula(expression, context);
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
        return {
          type: "error",
          err: FormulaError.ERROR,
          message: `Unknown binary operator: ${operator}`,
          errAddress: context.dependencyNode,
        };
    }
  }

  evaluateFunction(
    node: FunctionNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const func = functions[node.name];
    if (!func) {
      return {
        type: "error",
        err: FormulaError.NAME,
        message: `Function ${node.name} not found`,
        errAddress: context.dependencyNode,
      };
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
