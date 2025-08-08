/**
 * Core evaluation engine for FormulaEngine
 * Handles AST evaluation, context management, value coercion, and evaluation strategies
 */

import {
  type CellValue,
  type SimpleCellAddress,
  type SimpleCellRange,
  FormulaError,
  type NamedExpression,
} from "../core/types";
import type {
  ASTNode,
  ValueNode,
  ReferenceNode,
  RangeNode,
  FunctionNode,
  BinaryOpNode,
  UnaryOpNode,
  ArrayNode,
  NamedExpressionNode,
  ErrorNode,
} from "../parser/ast";
import { DependencyGraph } from "./dependency-graph";
import {
  isFormulaError,
  createStandardError,
  validateNumericArgument,
  ErrorHandler,
  mapJSErrorToFormulaError,
} from "./error-handler";
import {
  to2DArray,
  elementWiseBinaryOp,
  elementWiseUnaryOp,
  getArrayDimensions,
  broadcastToSize,
} from "./array-evaluator";
import { parseFormula, type SheetResolver } from "../parser/parser";

/**
 * Evaluation context containing necessary information
 */
export interface EvaluationContext {
  currentSheet: number;
  currentCell?: SimpleCellAddress;
  namedExpressions: Map<string, NamedExpression>;
  getCellValue: (address: SimpleCellAddress) => CellValue;
  getRangeValues: (
    range: SimpleCellRange,
    evaluationStack?: Set<string>
  ) => CellValue[][];
  getFunction: (name: string) => FunctionDefinition | undefined;
  errorHandler: ErrorHandler;
  evaluationStack: Set<string>; // For cycle detection
  sheetResolver?: SheetResolver; // For resolving sheet names in formulas
}

export type FunctionEvaluationResult =
  | {
      type: "value";
      value: CellValue;
    }
  | {
      type: "2d-array";
      value: CellValue[][];
      dimensions: { rows: number; cols: number };
    };

/**
 * Function definition
 */
export interface FunctionDefinition {
  name: string;
  minArgs?: number;
  maxArgs?: number;
  evaluate: (args: {
    /**
     * Flattened array of argument values, if two ranges are provided E.g. A1:A3, B1:B3, the result will be [A1, A2, A3, B1, B2, B3]
     */
    flatArgValues: CellValue[];
    /**
     * Evaluated argument values, E.g. A1:A3, B1:B3, C4, the result could be:
     * ```
     * [{ type: '2d-array', value: [[1, 2, 3]], dimensions: { rows: 3, cols: 1 } },
     *  { type: '2d-array', value: [[4, 5, 6]], dimensions: { rows: 3, cols: 1 } },
     *  { type: 'value', value: 7 }]
     * ```
     */
    argEvaluatedValues: EvaluationResult[];
    /**
     * The raw argument nodes
     */
    argNodes: ASTNode[];
    context: EvaluationContext;
  }) => FunctionEvaluationResult;
}

/**
 * Evaluation result
 */
export type EvaluationResult = {
  dependencies: Set<string>;
} & FunctionEvaluationResult;
/**
 * Main evaluator class
 */
export class Evaluator {
  private dependencyGraph: DependencyGraph;
  private functions: Map<string, FunctionDefinition>;
  private errorHandler: ErrorHandler;
  private evaluationCache: Map<string, CellValue>;

  constructor(
    dependencyGraph: DependencyGraph,
    functions: Map<string, FunctionDefinition>,
    errorHandler: ErrorHandler
  ) {
    this.dependencyGraph = dependencyGraph;
    this.functions = functions;
    this.errorHandler = errorHandler;
    this.evaluationCache = new Map();
  }

  /**
   * Evaluates an AST node
   */
  evaluate(node: ASTNode, context: EvaluationContext): EvaluationResult {
    const dependencies = new Set<string>();

    try {
      const result = this.evaluateNode(node, context, dependencies);

      return result;
    } catch (error) {
      // Convert JavaScript errors to formula errors
      const formulaError =
        error instanceof Error
          ? mapJSErrorToFormulaError(error)
          : FormulaError.ERROR;

      return {
        type: "value",
        value: formulaError,
        dependencies,
      };
    }
  }

  /**
   * Evaluates a single node
   */
  private evaluateNode(
    node: ASTNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): EvaluationResult {
    switch (node.type) {
      case "value":
        return {
          type: "value",
          value: this.evaluateValue(node),
          dependencies,
        };

      case "reference":
        return this.evaluateReference(node, context, dependencies);

      case "range": {
        const value = this.evaluateRange(node, context, dependencies);
        return {
          type: "2d-array",
          value,
          dimensions: getArrayDimensions(value),
          dependencies,
        };
      }

      case "function": {
        const result = this.evaluateFunction(node, context, dependencies);

        // Check if the result is a 2D array (for array functions like FILTER, SORT, etc.)
        if (Array.isArray(result) && Array.isArray(result[0])) {
          return {
            type: "2d-array",
            value: result as CellValue[][],
            dimensions: getArrayDimensions(result),
            dependencies,
          };
        }

        return result;
      }

      case "binary-op":
        return this.evaluateBinaryOp(node, context, dependencies);

      case "unary-op":
        return this.evaluateUnaryOp(node, context, dependencies);

      case "array": {
        const result = this.evaluateArray(node, context, dependencies);
        return {
          type: "2d-array",
          value: result,
          dimensions: getArrayDimensions(result),
          dependencies,
        };
      }

      case "named-expression": {
        return this.evaluateNamedExpression(node, context, dependencies);
      }

      case "error":
        return {
          type: "value",
          value: this.evaluateError(node),
          dependencies,
        };

      default:
        return {
          type: "value",
          value: FormulaError.ERROR,
          dependencies,
        };
    }
  }

  /**
   * Evaluates a value node
   */
  private evaluateValue(node: ValueNode): CellValue {
    return node.value;
  }

  /**
   * Evaluates an error node
   */
  private evaluateError(node: ErrorNode): FormulaError {
    return node.error;
  }

  /**
   * Evaluates a named expression node
   */
  private evaluateNamedExpression(
    node: NamedExpressionNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): EvaluationResult {
    const namedExpr = this.resolveNamedExpression(node.name, context);
    if (!namedExpr) {
      return {
        type: "value",
        value: "#NAME?",
        dependencies,
      };
    }

    const key = DependencyGraph.getNamedExpressionKey(
      namedExpr.name,
      namedExpr.scope
    );
    dependencies.add(key);

    // Check for circular reference
    if (context.evaluationStack.has(key)) {
      return {
        type: "value",
        value: "#CYCLE!",
        dependencies,
      };
    }

    // Parse and evaluate named expression
    if (typeof namedExpr.expression === "string") {
      if (!namedExpr.expression.startsWith("=")) {
        // Simple value
        const num = parseFloat(namedExpr.expression);
        if (!isNaN(num))
          return {
            type: "value",
            value: num,
            dependencies,
          };
        if (namedExpr.expression === "TRUE")
          return {
            type: "value",
            value: true,
            dependencies,
          };
      } else {
        // Formula expression - need to evaluate it
        try {
          // Add this named expression to the evaluation stack to detect cycles
          context.evaluationStack.add(key);

          // Parse the formula (remove the = prefix)
          const formula = namedExpr.expression.substring(1);
          const ast = parseFormula(
            formula,
            context.currentSheet,
            context.sheetResolver
          );

          // Create a new context for evaluating the named expression
          const nestedContext: EvaluationContext = {
            ...context,
            evaluationStack: new Set(context.evaluationStack),
          };

          const result = this.evaluate(ast, nestedContext);

          // Merge dependencies from the nested evaluation
          for (const dep of result.dependencies) {
            dependencies.add(dep);
          }

          // Remove from evaluation stack
          context.evaluationStack.delete(key);

          return result;
        } catch (error) {
          context.evaluationStack.delete(key);
          return {
            type: "value",
            value: "#NAME?",
            dependencies,
          };
        }
      }
    }

    return {
      type: "value",
      value: "#NAME?",
      dependencies,
    };
  }

  /**
   * Helper to check if a cell is within a range
   */
  private isCellInRange(
    cell: SimpleCellAddress,
    range: SimpleCellRange
  ): boolean {
    return (
      cell.sheet === range.start.sheet &&
      cell.row >= range.start.row &&
      cell.row <= range.end.row &&
      cell.col >= range.start.col &&
      cell.col <= range.end.col
    );
  }

  /**
   * Evaluates a reference node
   */
  private evaluateReference(
    node: ReferenceNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): EvaluationResult {
    const address = node.address;
    const key = DependencyGraph.getCellKey(address);
    dependencies.add(key);

    // Check for circular reference
    if (context.evaluationStack.has(key)) {
      return {
        type: "value",
        value: "#CYCLE!",
        dependencies,
      };
    }

    return {
      type: "value",
      value: context.getCellValue(address),
      dependencies,
    };
  }

  /**
   * Evaluates a range node
   */
  private evaluateRange(
    node: RangeNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): CellValue[][] {
    const range = node.range;

    // Add range to dependencies
    const key = DependencyGraph.getRangeKey(range);
    dependencies.add(key);

    // Check if this is an infinite range
    const isInfiniteColumn = range.end.row === Number.MAX_SAFE_INTEGER;
    const isInfiniteRow = range.end.col === Number.MAX_SAFE_INTEGER;

    // Check for circular reference: if current cell is within this range
    if (context.currentCell && this.isCellInRange(context.currentCell, range)) {
      // We're trying to evaluate a range that includes the current cell
      const cellKey = DependencyGraph.getCellKey(context.currentCell);
      if (context.evaluationStack.has(cellKey)) {
        // This creates a circular reference
        // Return an array filled with #CYCLE! errors
        const rows = isInfiniteColumn
          ? 1
          : Math.min(range.end.row - range.start.row + 1, 1000);
        const cols = isInfiniteRow
          ? 1
          : Math.min(range.end.col - range.start.col + 1, 1000);
        const result: CellValue[][] = [];
        for (let r = 0; r < rows; r++) {
          const row: CellValue[] = [];
          for (let c = 0; c < cols; c++) {
            row.push("#CYCLE!");
          }
          result.push(row);
        }
        return result;
      }
    }

    if (!isInfiniteColumn && !isInfiniteRow) {
      // Normal range - add individual cells to dependencies
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.col; col <= range.end.col; col++) {
          const cellKey = DependencyGraph.getCellKey({
            sheet: range.start.sheet,
            col,
            row,
          });
          dependencies.add(cellKey);
        }
      }
    } else {
      // For infinite ranges, we'll let getRangeValues handle the sparse iteration
      // and add dependencies dynamically based on actual populated cells
      // This is handled by the getRangeValues implementation
    }

    return context.getRangeValues(range, context.evaluationStack);
  }

  /**
   * Evaluates a function call
   */
  private evaluateFunction(
    node: FunctionNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): EvaluationResult {
    const func = context.getFunction(node.name.toUpperCase());
    if (!func) {
      return {
        type: "value",
        value: FormulaError.NAME,
        dependencies,
      };
    }

    // Evaluate arguments
    const argValues: (CellValue | CellValue[][])[] = [];
    const argEvaluatedValues: EvaluationResult[] = [];

    // IS* functions should receive errors as arguments without propagation
    const upperName = node.name.toUpperCase();
    const isErrorCheckingFunction = upperName.startsWith("IS");

    // IF function needs special handling - only propagate errors from condition
    const isIfFunction = upperName === "IF";

    // Error-handling functions like IFERROR must also receive error args
    const isErrorHandlingFunction =
      upperName === "IFERROR" || upperName === "IFNA";

    for (let i = 0; i < node.args.length; i++) {
      const argNode = node.args[i];
      if (!argNode) continue;
      const argResult = this.evaluateNode(argNode, context, dependencies);

      // Check for errors in arguments
      if (
        argResult.type === "value" &&
        isFormulaError(argResult.value) &&
        !isErrorCheckingFunction &&
        !isErrorHandlingFunction
      ) {
        // For IF, only propagate errors from the condition (first argument)
        if (!isIfFunction || i === 0) {
          return argResult;
        }
      }
      argValues.push(argResult.value);
      argEvaluatedValues.push(argResult);
    }

    // Check argument count
    if (func.minArgs !== undefined && argValues.length < func.minArgs) {
      return {
        type: "value",
        value: FormulaError.VALUE,
        dependencies,
      };
    }

    if (func.maxArgs !== undefined && argValues.length > func.maxArgs) {
      return {
        type: "value",
        value: FormulaError.VALUE,
        dependencies,
      };
    }

    // Execute function
    try {
      const result = func.evaluate({
        flatArgValues: argValues.flatMap((arg) =>
          Array.isArray(arg) ? arg.flat() : [arg]
        ),
        argEvaluatedValues,
        argNodes: node.args,
        context,
      });

      return {
        ...result,
        dependencies,
      };
    } catch (error) {
      if (typeof error === "string" && isFormulaError(error)) {
        return {
          type: "value",
          value: error,
          dependencies,
        };
      }

      return {
        type: "value",
        value: FormulaError.ERROR,
        dependencies,
      };
    }
  }

  /**
   * Evaluates a binary operation
   */
  private evaluateBinaryOp(
    node: BinaryOpNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): EvaluationResult {
    const left = this.evaluateNode(node.left, context, dependencies);
    const right = this.evaluateNode(node.right, context, dependencies);

    // Handle array operations
    if (left.type === "2d-array" || right.type === "2d-array") {
      const leftArray = to2DArray(left.value);
      const rightArray = to2DArray(right.value);

      const operation = this.getBinaryOperation(node.operator);
      const result = elementWiseBinaryOp(leftArray, rightArray, operation);

      if (typeof result === "string" && isFormulaError(result)) {
        return {
          type: "value",
          value: result,
          dependencies,
        };
      }

      // Return as single value if 1x1 array
      if (result.length === 1 && result[0]?.length === 1) {
        return {
          type: "value",
          value: result[0][0],
          dependencies,
        };
      }

      return {
        type: "2d-array",
        value: result,
        dimensions: getArrayDimensions(result),
        dependencies,
      };
    }

    // Scalar operation
    return {
      type: "value",
      value: this.evaluateBinaryScalar(node.operator, left.value, right.value),
      dependencies,
    };
  }

  /**
   * Evaluates a unary operation
   */
  private evaluateUnaryOp(
    node: UnaryOpNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): EvaluationResult {
    const operand = this.evaluateNode(node.operand, context, dependencies);

    // Handle array operations
    if (operand.type === "2d-array") {
      const array = to2DArray(operand.value);
      const operation = this.getUnaryOperation(node.operator);
      const result = elementWiseUnaryOp(array, operation);

      // Return as single value if 1x1 array
      if (result.length === 1 && result[0]?.length === 1) {
        return {
          type: "value",
          value: result[0][0],
          dependencies,
        };
      }

      return {
        type: "2d-array",
        value: result,
        dimensions: getArrayDimensions(result),
        dependencies,
      };
    }

    // Scalar operation
    return {
      type: "value",
      value: this.evaluateUnaryScalar(node.operator, operand.value),
      dependencies,
    };
  }

  /**
   * Evaluates an array node
   */
  private evaluateArray(
    node: ArrayNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): CellValue[][] {
    const result: CellValue[][] = [];

    for (const row of node.elements) {
      const evaluatedRow: CellValue[] = [];

      for (const element of row) {
        const result = this.evaluateNode(element, context, dependencies);

        const value = result.value;

        // Flatten nested arrays in array literals
        if (Array.isArray(value) && Array.isArray(value[0])) {
          // This is a 2D array - take first element
          evaluatedRow.push(value[0][0]);
        } else if (Array.isArray(value) && !Array.isArray(value[0])) {
          // 1D array - take first element
          evaluatedRow.push(value[0]);
        } else if (!Array.isArray(value)) {
          evaluatedRow.push(value);
        }
      }

      result.push(evaluatedRow);
    }

    return result;
  }

  /**
   * Binary scalar operations
   */
  private evaluateBinaryScalar(
    operator: string,
    left: CellValue,
    right: CellValue
  ): CellValue {
    // Check for errors
    if (isFormulaError(left)) return left;
    if (isFormulaError(right)) return right;

    switch (operator) {
      case "+":
        return this.add(left, right);
      case "-":
        return this.subtract(left, right);
      case "*":
        return this.multiply(left, right);
      case "/":
        return this.divide(left, right);
      case "^":
        return this.power(left, right);
      case "&":
        return this.concatenate(left, right);
      case "=":
        return this.equals(left, right);
      case "<>":
        return this.notEquals(left, right);
      case "<":
        return this.lessThan(left, right);
      case "<=":
        return this.lessThanOrEqual(left, right);
      case ">":
        return this.greaterThan(left, right);
      case ">=":
        return this.greaterThanOrEqual(left, right);
      default:
        return "#ERROR!";
    }
  }

  /**
   * Unary scalar operations
   */
  private evaluateUnaryScalar(operator: string, operand: CellValue): CellValue {
    if (isFormulaError(operand)) return operand;

    switch (operator) {
      case "-":
        return this.negate(operand);
      case "+":
        return this.unaryPlus(operand);
      case "%":
        return this.percent(operand);
      default:
        return "#ERROR!";
    }
  }

  /**
   * Gets binary operation function
   */
  private getBinaryOperation(
    operator: string
  ): (a: CellValue, b: CellValue) => CellValue {
    return (a, b) => this.evaluateBinaryScalar(operator, a, b);
  }

  /**
   * Gets unary operation function
   */
  private getUnaryOperation(operator: string): (value: CellValue) => CellValue {
    return (value) => this.evaluateUnaryScalar(operator, value);
  }

  // Arithmetic operations

  private add(left: CellValue, right: CellValue): CellValue {
    const leftNum = this.coerceToNumber(left);
    const rightNum = this.coerceToNumber(right);

    if (isFormulaError(leftNum)) return leftNum;
    if (isFormulaError(rightNum)) return rightNum;

    return leftNum + rightNum;
  }

  private subtract(left: CellValue, right: CellValue): CellValue {
    const leftNum = this.coerceToNumber(left);
    const rightNum = this.coerceToNumber(right);

    if (isFormulaError(leftNum)) return leftNum;
    if (isFormulaError(rightNum)) return rightNum;

    return leftNum - rightNum;
  }

  private multiply(left: CellValue, right: CellValue): CellValue {
    const leftNum = this.coerceToNumber(left);
    const rightNum = this.coerceToNumber(right);

    if (isFormulaError(leftNum)) return leftNum;
    if (isFormulaError(rightNum)) return rightNum;

    return leftNum * rightNum;
  }

  private divide(left: CellValue, right: CellValue): CellValue {
    const leftNum = this.coerceToNumber(left);
    const rightNum = this.coerceToNumber(right);

    if (isFormulaError(leftNum)) return leftNum;
    if (isFormulaError(rightNum)) return rightNum;

    if (rightNum === 0) {
      return "#DIV/0!";
    }

    return leftNum / rightNum;
  }

  private power(left: CellValue, right: CellValue): CellValue {
    const leftNum = this.coerceToNumber(left);
    const rightNum = this.coerceToNumber(right);

    if (isFormulaError(leftNum)) return leftNum;
    if (isFormulaError(rightNum)) return rightNum;

    const result = Math.pow(leftNum, rightNum);

    if (!isFinite(result)) {
      return "#NUM!";
    }

    return result;
  }

  private negate(operand: CellValue): CellValue {
    const num = this.coerceToNumber(operand);
    if (isFormulaError(num)) return num;
    return -num;
  }

  private unaryPlus(operand: CellValue): CellValue {
    const num = this.coerceToNumber(operand);
    if (isFormulaError(num)) return num;
    return num;
  }

  private percent(operand: CellValue): CellValue {
    const num = this.coerceToNumber(operand);
    if (isFormulaError(num)) return num;
    return num / 100;
  }

  // String operations

  private concatenate(left: CellValue, right: CellValue): CellValue {
    const leftStr = this.coerceToString(left);
    const rightStr = this.coerceToString(right);

    if (isFormulaError(leftStr)) return leftStr;
    if (isFormulaError(rightStr)) return rightStr;

    return leftStr + rightStr;
  }

  // Comparison operations

  private equals(left: CellValue, right: CellValue): boolean {
    // Handle errors
    if (isFormulaError(left) || isFormulaError(right)) {
      return left === right;
    }

    // Handle undefined (empty cells)
    if (left === undefined && right === undefined) return true;
    if (left === undefined || right === undefined) return false;

    // Same type comparison
    if (typeof left === typeof right) {
      return left === right;
    }

    // Different types are not equal in Excel
    return false;
  }

  private notEquals(left: CellValue, right: CellValue): boolean {
    return !this.equals(left, right);
  }

  private lessThan(left: CellValue, right: CellValue): CellValue {
    const result = this.compare(left, right);
    if (isFormulaError(result)) return result;
    return result < 0;
  }

  private lessThanOrEqual(left: CellValue, right: CellValue): CellValue {
    const result = this.compare(left, right);
    if (isFormulaError(result)) return result;
    return result <= 0;
  }

  private greaterThan(left: CellValue, right: CellValue): CellValue {
    const result = this.compare(left, right);
    if (isFormulaError(result)) return result;
    return result > 0;
  }

  private greaterThanOrEqual(left: CellValue, right: CellValue): CellValue {
    const result = this.compare(left, right);
    if (isFormulaError(result)) return result;
    return result >= 0;
  }

  /**
   * Compares two values according to Excel rules
   */
  private compare(left: CellValue, right: CellValue): number | FormulaError {
    // Handle errors
    if (isFormulaError(left)) return left;
    if (isFormulaError(right)) return right;

    // Handle empty cells
    if (left === undefined) left = 0;
    if (right === undefined) right = 0;

    // Numbers
    if (typeof left === "number" && typeof right === "number") {
      return left - right;
    }

    // Strings
    if (typeof left === "string" && typeof right === "string") {
      return left.localeCompare(right);
    }

    // Booleans
    if (typeof left === "boolean" && typeof right === "boolean") {
      return (left ? 1 : 0) - (right ? 1 : 0);
    }

    // Mixed types - Excel type hierarchy: number < string < boolean
    const typeOrder = { number: 0, string: 1, boolean: 2 };
    const leftType = typeof left as keyof typeof typeOrder;
    const rightType = typeof right as keyof typeof typeOrder;

    return typeOrder[leftType] - typeOrder[rightType];
  }

  // Type coercion

  /**
   * Coerces a value to number
   */
  coerceToNumber(value: CellValue): number | FormulaError {
    if (isFormulaError(value)) return value;

    if (typeof value === "number") return value;
    if (typeof value === "boolean") return value ? 1 : 0;
    if (value === undefined) return 0;

    if (typeof value === "string") {
      if (value === "") return 0;
      const num = parseFloat(value);
      if (isNaN(num)) return FormulaError.VALUE;
      return num;
    }

    return FormulaError.VALUE;
  }

  /**
   * Coerces a value to string
   */
  coerceToString(value: CellValue): string | FormulaError {
    if (isFormulaError(value)) return value;

    if (typeof value === "string") return value;
    if (typeof value === "number") return value.toString();
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    if (value === undefined) return "";

    return "#VALUE!";
  }

  /**
   * Coerces a value to boolean
   */
  coerceToBoolean(value: CellValue): boolean | FormulaError {
    if (isFormulaError(value)) return value;

    if (typeof value === "boolean") return value;
    if (typeof value === "number") return value !== 0;
    if (typeof value === "string") {
      const upper = value.toUpperCase();
      if (upper === "TRUE") return true;
      if (upper === "FALSE") return false;
      return value.length > 0;
    }
    if (value === undefined) return false;

    return FormulaError.VALUE;
  }

  /**
   * Resolves a named expression
   */
  private resolveNamedExpression(
    name: string,
    context: EvaluationContext
  ): NamedExpression | undefined {
    // Check sheet-scoped names first
    if (context.currentSheet !== undefined) {
      const sheetScoped = `${context.currentSheet}:${name}`;
      if (context.namedExpressions.has(sheetScoped)) {
        return context.namedExpressions.get(sheetScoped);
      }
    }

    // Check global names
    return context.namedExpressions.get(name);
  }

  /**
   * Clears evaluation cache
   */
  clearCache(): void {
    this.evaluationCache.clear();
  }
}
