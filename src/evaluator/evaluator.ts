/**
 * Core evaluation engine for FormulaEngine
 * Handles AST evaluation, context management, value coercion, and evaluation strategies
 */

import type { 
  CellValue, 
  SimpleCellAddress, 
  SimpleCellRange,
  FormulaError,
  NamedExpression
} from '../core/types';
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
  ErrorNode
} from '../parser/ast';
import { DependencyGraph } from './dependency-graph';
import { 
  isFormulaError, 
  createStandardError,
  validateNumericArgument,
  ErrorHandler,
  mapJSErrorToFormulaError
} from './error-handler';
import {
  to2DArray,
  elementWiseBinaryOp,
  elementWiseUnaryOp,
  getArrayDimensions,
  broadcastToSize
} from './array-evaluator';
import { parseFormula } from '../parser/parser';

/**
 * Evaluation context containing necessary information
 */
export interface EvaluationContext {
  currentSheet: number;
  currentCell?: SimpleCellAddress;
  namedExpressions: Map<string, NamedExpression>;
  getCellValue: (address: SimpleCellAddress) => CellValue;
  getRangeValues: (range: SimpleCellRange) => CellValue[][];
  getFunction: (name: string) => FunctionDefinition | undefined;
  errorHandler: ErrorHandler;
  evaluationStack: Set<string>; // For cycle detection
  arrayContext?: ArrayEvaluationContext;
}

/**
 * Array evaluation context
 */
export interface ArrayEvaluationContext {
  isArrayFormula: boolean;
  targetAddress?: SimpleCellAddress;
  spillRange?: SimpleCellRange;
}

/**
 * Function definition
 */
export interface FunctionDefinition {
  name: string;
  minArgs?: number;
  maxArgs?: number;
  evaluate: (args: CellValue[], context: EvaluationContext) => CellValue;
  isVolatile?: boolean;
  acceptsArrays?: boolean;
  returnsArray?: boolean;
}

/**
 * Evaluation result
 */
export interface EvaluationResult {
  value: CellValue;
  dependencies: Set<string>;
  isArrayResult: boolean;
  arrayDimensions?: { rows: number; cols: number };
}

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
  evaluate(
    node: ASTNode,
    context: EvaluationContext
  ): EvaluationResult {
    const dependencies = new Set<string>();
    
    try {
      const value = this.evaluateNode(node, context, dependencies);
      
      // Check if result is an array
      const isArrayResult = Array.isArray(value);
      const arrayDimensions = isArrayResult ? getArrayDimensions(value) : undefined;
      
      return {
        value,
        dependencies,
        isArrayResult,
        arrayDimensions
      };
    } catch (error) {
      // Convert JavaScript errors to formula errors
      const formulaError = error instanceof Error
        ? mapJSErrorToFormulaError(error)
        : '#ERROR!';
        
      return {
        value: formulaError,
        dependencies,
        isArrayResult: false
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
  ): CellValue {
    switch (node.type) {
      case 'value':
        return this.evaluateValue(node as ValueNode);
        
      case 'reference':
        return this.evaluateReference(node as ReferenceNode, context, dependencies);
        
      case 'range': {
        const result = this.evaluateRange(node as RangeNode, context, dependencies);
        return result as unknown as CellValue;
      }
        
      case 'function':
        return this.evaluateFunction(node as FunctionNode, context, dependencies);
        
      case 'binary-op':
        return this.evaluateBinaryOp(node as BinaryOpNode, context, dependencies);
        
      case 'unary-op':
        return this.evaluateUnaryOp(node as UnaryOpNode, context, dependencies);
        
      case 'array': {
        const result = this.evaluateArray(node as ArrayNode, context, dependencies);
        return result as unknown as CellValue;
      }
        
      case 'named-expression':
        return this.evaluateNamedExpression(node as NamedExpressionNode, context, dependencies);
        
      case 'error':
        return this.evaluateError(node as ErrorNode);
        
      default:
        return createStandardError('#ERROR!', {
          message: `Unknown node type: ${(node as any).type}`
        }).type;
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
  ): CellValue {
    const namedExpr = this.resolveNamedExpression(node.name, context);
    if (!namedExpr) {
      return '#NAME?';
    }
    
    const key = DependencyGraph.getNamedExpressionKey(
      namedExpr.name,
      namedExpr.scope
    );
    dependencies.add(key);
    
    // Check for circular reference
    if (context.evaluationStack.has(key)) {
      return '#CYCLE!';
    }
    
    // Parse and evaluate named expression
    if (typeof namedExpr.expression === 'string') {
      if (!namedExpr.expression.startsWith('=')) {
        // Simple value
        const num = parseFloat(namedExpr.expression);
        if (!isNaN(num)) return num;
        if (namedExpr.expression === 'TRUE') return true;
        if (namedExpr.expression === 'FALSE') return false;
        return namedExpr.expression;
      } else {
        // Formula expression - need to evaluate it
        try {
          // Add this named expression to the evaluation stack to detect cycles
          context.evaluationStack.add(key);
          
          // Parse the formula (remove the = prefix)
          const formula = namedExpr.expression.substring(1);
          const ast = parseFormula(formula, context.currentSheet);
          
          // Create a new context for evaluating the named expression
          const nestedContext: EvaluationContext = {
            ...context,
            evaluationStack: new Set(context.evaluationStack)
          };
          
          const result = this.evaluate(ast, nestedContext);
          
          // Merge dependencies from the nested evaluation
          for (const dep of result.dependencies) {
            dependencies.add(dep);
          }
          
          // Remove from evaluation stack
          context.evaluationStack.delete(key);
          
          return result.value;
        } catch (error) {
          context.evaluationStack.delete(key);
          return '#NAME?';
        }
      }
    }
    
    return '#NAME?';
  }
  
  /**
   * Evaluates a reference node
   */
  private evaluateReference(
    node: ReferenceNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): CellValue {
    const address = node.address;
    const key = DependencyGraph.getCellKey(address);
    dependencies.add(key);
    
    // Check for circular reference
    if (context.evaluationStack.has(key)) {
      return '#CYCLE!';
    }
    
    return context.getCellValue(address);
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
    
    // Add individual cells to dependencies
    for (let row = range.start.row; row <= range.end.row; row++) {
      for (let col = range.start.col; col <= range.end.col; col++) {
        const cellKey = DependencyGraph.getCellKey({
          sheet: range.start.sheet,
          col,
          row
        });
        dependencies.add(cellKey);
      }
    }
    
    return context.getRangeValues(range);
  }
  
  /**
   * Evaluates a function call
   */
  private evaluateFunction(
    node: FunctionNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): CellValue {
    const func = context.getFunction(node.name.toUpperCase());
    if (!func) {
      return createStandardError('#NAME?', {
        functionName: node.name,
        message: `Unknown function: ${node.name}`
      }).type;
    }
    
    // Evaluate arguments
    const args: CellValue[] = [];
    
    // IS* functions should receive errors as arguments without propagation
    const isErrorCheckingFunction = node.name.toUpperCase().startsWith('IS');
    
    // IF function needs special handling - only propagate errors from condition
    const isIfFunction = node.name.toUpperCase() === 'IF';
    
    for (let i = 0; i < node.args.length; i++) {
      const argNode = node.args[i];
      if (!argNode) continue;
      const argValue = this.evaluateNode(argNode, context, dependencies);
      
      // Check for errors in arguments
      if (isFormulaError(argValue) && !isErrorCheckingFunction) {
        // For IF, only propagate errors from the condition (first argument)
        if (!isIfFunction || i === 0) {
          return argValue;
        }
      }
      args.push(argValue);
    }
    
    // Check argument count
    if (func.minArgs !== undefined && args.length < func.minArgs) {
      return createStandardError('#VALUE!', {
        functionName: node.name,
        message: `Too few arguments for ${node.name}`
      }).type;
    }
    
    if (func.maxArgs !== undefined && args.length > func.maxArgs) {
      return createStandardError('#VALUE!', {
        functionName: node.name,
        message: `Too many arguments for ${node.name}`
      }).type;
    }
    
    // Execute function
    try {
      return func.evaluate(args, context);
    } catch (error) {
      if (typeof error === 'string' && isFormulaError(error)) {
        return error;
      }
      
      return createStandardError('#ERROR!', {
        functionName: node.name,
        message: error instanceof Error ? error.message : 'Function evaluation failed'
      }).type;
    }
  }
  
  /**
   * Evaluates a binary operation
   */
  private evaluateBinaryOp(
    node: BinaryOpNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): CellValue {
    const left = this.evaluateNode(node.left, context, dependencies);
    const right = this.evaluateNode(node.right, context, dependencies);
    
    // Handle array operations
    if (Array.isArray(left) || Array.isArray(right)) {
      const leftArray = to2DArray(left);
      const rightArray = to2DArray(right);
      
      const operation = this.getBinaryOperation(node.operator);
      const result = elementWiseBinaryOp(leftArray, rightArray, operation);
      
      if (typeof result === 'string' && isFormulaError(result)) return result;
      
      // Return as single value if 1x1 array
      if (result.length === 1 && result[0]?.length === 1) {
        return result[0][0];
      }
      
      return result as unknown as CellValue;
    }
    
    // Scalar operation
    return this.evaluateBinaryScalar(node.operator, left, right);
  }
  
  /**
   * Evaluates a unary operation
   */
  private evaluateUnaryOp(
    node: UnaryOpNode,
    context: EvaluationContext,
    dependencies: Set<string>
  ): CellValue {
    const operand = this.evaluateNode(node.operand, context, dependencies);
    
    // Handle array operations
    if (Array.isArray(operand)) {
      const array = to2DArray(operand);
      const operation = this.getUnaryOperation(node.operator);
      const result = elementWiseUnaryOp(array, operation);
      
      // Return as single value if 1x1 array
      if (result.length === 1 && result[0]?.length === 1) {
        return result[0][0];
      }
      
      return result as unknown as CellValue;
    }
    
    // Scalar operation
    return this.evaluateUnaryScalar(node.operator, operand);
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
        const value = this.evaluateNode(element, context, dependencies);
        
        // Flatten nested arrays in array literals
        if (Array.isArray(value) && Array.isArray(value[0])) {
          // This is a 2D array - take first element
          evaluatedRow.push(value[0][0]);
        } else if (Array.isArray(value)) {
          // 1D array - take first element
          evaluatedRow.push(value[0]);
        } else {
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
  private evaluateBinaryScalar(operator: string, left: CellValue, right: CellValue): CellValue {
    // Check for errors
    if (isFormulaError(left)) return left;
    if (isFormulaError(right)) return right;
    
    switch (operator) {
      case '+':
        return this.add(left, right);
      case '-':
        return this.subtract(left, right);
      case '*':
        return this.multiply(left, right);
      case '/':
        return this.divide(left, right);
      case '^':
        return this.power(left, right);
      case '&':
        return this.concatenate(left, right);
      case '=':
        return this.equals(left, right);
      case '<>':
        return this.notEquals(left, right);
      case '<':
        return this.lessThan(left, right);
      case '<=':
        return this.lessThanOrEqual(left, right);
      case '>':
        return this.greaterThan(left, right);
      case '>=':
        return this.greaterThanOrEqual(left, right);
      default:
        return '#ERROR!';
    }
  }
  
  /**
   * Unary scalar operations
   */
  private evaluateUnaryScalar(operator: string, operand: CellValue): CellValue {
    if (isFormulaError(operand)) return operand;
    
    switch (operator) {
      case '-':
        return this.negate(operand);
      case '+':
        return this.unaryPlus(operand);
      case '%':
        return this.percent(operand);
      default:
        return '#ERROR!';
    }
  }
  
  /**
   * Gets binary operation function
   */
  private getBinaryOperation(operator: string): (a: CellValue, b: CellValue) => CellValue {
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
      return '#DIV/0!';
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
      return '#NUM!';
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
    if (typeof left === 'number' && typeof right === 'number') {
      return left - right;
    }
    
    // Strings
    if (typeof left === 'string' && typeof right === 'string') {
      return left.localeCompare(right);
    }
    
    // Booleans
    if (typeof left === 'boolean' && typeof right === 'boolean') {
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
    
    if (typeof value === 'number') return value;
    if (typeof value === 'boolean') return value ? 1 : 0;
    if (value === undefined) return 0;
    
    if (typeof value === 'string') {
      if (value === '') return 0;
      const num = parseFloat(value);
      if (isNaN(num)) return '#VALUE!';
      return num;
    }
    
    return '#VALUE!';
  }
  
  /**
   * Coerces a value to string
   */
  coerceToString(value: CellValue): string | FormulaError {
    if (isFormulaError(value)) return value;
    
    if (typeof value === 'string') return value;
    if (typeof value === 'number') return value.toString();
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (value === undefined) return '';
    
    return '#VALUE!';
  }
  
  /**
   * Coerces a value to boolean
   */
  coerceToBoolean(value: CellValue): boolean | FormulaError {
    if (isFormulaError(value)) return value;
    
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value !== 0;
    if (typeof value === 'string') {
      const upper = value.toUpperCase();
      if (upper === 'TRUE') return true;
      if (upper === 'FALSE') return false;
      return value.length > 0;
    }
    if (value === undefined) return false;
    
    return '#VALUE!';
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
