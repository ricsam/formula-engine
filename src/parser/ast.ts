/**
 * Abstract Syntax Tree (AST) node definitions for formula parsing
 */

import type { SimpleCellAddress, SimpleCellRange, CellValue, FormulaError } from '../core/types';

/**
 * Base AST node type
 */
export type ASTNodeType = 
  | 'value'
  | 'reference'
  | 'range'
  | 'function'
  | 'unary-op'
  | 'binary-op'
  | 'array'
  | 'named-expression'
  | 'error';

/**
 * Base interface for all AST nodes
 */
export interface ASTNode {
  type: ASTNodeType;
  position?: {
    start: number;
    end: number;
  };
}

/**
 * Literal value node (number, string, boolean, error)
 */
export interface ValueNode extends ASTNode {
  type: 'value';
  value: CellValue;
  valueType: 'number' | 'string' | 'boolean' | 'error';
}

/**
 * Cell reference node (e.g., A1, Sheet1!B2)
 */
export interface ReferenceNode extends ASTNode {
  type: 'reference';
  address: SimpleCellAddress;
  isAbsolute: {
    col: boolean;
    row: boolean;
  };
}

/**
 * Range reference node (e.g., A1:B10)
 */
export interface RangeNode extends ASTNode {
  type: 'range';
  range: SimpleCellRange;
  isAbsolute: {
    start: {
      col: boolean;
      row: boolean;
    };
    end: {
      col: boolean;
      row: boolean;
    };
  };
}

/**
 * Function call node (e.g., SUM(A1:A10))
 */
export interface FunctionNode extends ASTNode {
  type: 'function';
  name: string;
  args: ASTNode[];
}

/**
 * Unary operator node (e.g., -A1, +B2)
 */
export interface UnaryOpNode extends ASTNode {
  type: 'unary-op';
  operator: '+' | '-' | '%';
  operand: ASTNode;
}

/**
 * Binary operator node (e.g., A1+B1, C1*D1)
 */
export interface BinaryOpNode extends ASTNode {
  type: 'binary-op';
  operator: '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=';
  left: ASTNode;
  right: ASTNode;
}

/**
 * Array literal node (e.g., {1,2,3;4,5,6})
 */
export interface ArrayNode extends ASTNode {
  type: 'array';
  elements: ASTNode[][];  // 2D array of elements
}

/**
 * Named expression reference node
 */
export interface NamedExpressionNode extends ASTNode {
  type: 'named-expression';
  name: string;
  scope?: number;  // Sheet scope if specified
}

/**
 * Error node for parsing errors
 */
export interface ErrorNode extends ASTNode {
  type: 'error';
  error: FormulaError;
  message: string;
}

/**
 * Union type for all AST nodes
 */
export type FormulaAST = 
  | ValueNode
  | ReferenceNode
  | RangeNode
  | FunctionNode
  | UnaryOpNode
  | BinaryOpNode
  | ArrayNode
  | NamedExpressionNode
  | ErrorNode;

/**
 * Helper function to create a value node
 */
export function createValueNode(value: CellValue, start?: number, end?: number): ValueNode {
  let valueType: ValueNode['valueType'];
  
  if (typeof value === 'number') {
    valueType = 'number';
  } else if (typeof value === 'string') {
    // Check if it's an error string
    if (value.startsWith('#') && value.endsWith('!')) {
      valueType = 'error';
    } else {
      valueType = 'string';
    }
  } else if (typeof value === 'boolean') {
    valueType = 'boolean';
  } else {
    valueType = 'error';
  }
  
  return {
    type: 'value',
    value,
    valueType,
    ...(start !== undefined && end !== undefined && { position: { start, end } })
  };
}

/**
 * Helper function to create a reference node
 */
export function createReferenceNode(
  address: SimpleCellAddress, 
  isAbsolute: { col: boolean; row: boolean },
  start?: number,
  end?: number
): ReferenceNode {
  return {
    type: 'reference',
    address,
    isAbsolute,
    ...(start !== undefined && end !== undefined && { position: { start, end } })
  };
}

/**
 * Helper function to create a range node
 */
export function createRangeNode(
  range: SimpleCellRange,
  isAbsolute: RangeNode['isAbsolute'],
  start?: number,
  end?: number
): RangeNode {
  return {
    type: 'range',
    range,
    isAbsolute,
    ...(start !== undefined && end !== undefined && { position: { start, end } })
  };
}

/**
 * Helper function to create a function node
 */
export function createFunctionNode(
  name: string,
  args: ASTNode[],
  start?: number,
  end?: number
): FunctionNode {
  return {
    type: 'function',
    name: name.toUpperCase(),  // Normalize function names to uppercase
    args,
    ...(start !== undefined && end !== undefined && { position: { start, end } })
  };
}

/**
 * Helper function to create a unary operator node
 */
export function createUnaryOpNode(
  operator: UnaryOpNode['operator'],
  operand: ASTNode,
  start?: number,
  end?: number
): UnaryOpNode {
  return {
    type: 'unary-op',
    operator,
    operand,
    ...(start !== undefined && end !== undefined && { position: { start, end } })
  };
}

/**
 * Helper function to create a binary operator node
 */
export function createBinaryOpNode(
  operator: BinaryOpNode['operator'],
  left: ASTNode,
  right: ASTNode,
  start?: number,
  end?: number
): BinaryOpNode {
  return {
    type: 'binary-op',
    operator,
    left,
    right,
    ...(start !== undefined && end !== undefined && { position: { start, end } })
  };
}

/**
 * Helper function to create an array node
 */
export function createArrayNode(
  elements: ASTNode[][],
  start?: number,
  end?: number
): ArrayNode {
  return {
    type: 'array',
    elements,
    ...(start !== undefined && end !== undefined && { position: { start, end } })
  };
}

/**
 * Helper function to create a named expression node
 */
export function createNamedExpressionNode(
  name: string,
  scope?: number,
  start?: number,
  end?: number
): NamedExpressionNode {
  return {
    type: 'named-expression',
    name,
    scope,
    ...(start !== undefined && end !== undefined && { position: { start, end } })
  };
}

/**
 * Helper function to create an error node
 */
export function createErrorNode(
  error: FormulaError,
  message: string,
  start?: number,
  end?: number
): ErrorNode {
  return {
    type: 'error',
    error,
    message,
    ...(start !== undefined && end !== undefined && { position: { start, end } })
  };
}

/**
 * AST visitor interface for traversing the tree
 */
export interface ASTVisitor<T = void> {
  visitValue?(node: ValueNode): T;
  visitReference?(node: ReferenceNode): T;
  visitRange?(node: RangeNode): T;
  visitFunction?(node: FunctionNode): T;
  visitUnaryOp?(node: UnaryOpNode): T;
  visitBinaryOp?(node: BinaryOpNode): T;
  visitArray?(node: ArrayNode): T;
  visitNamedExpression?(node: NamedExpressionNode): T;
  visitError?(node: ErrorNode): T;
}

/**
 * Visit an AST node with a visitor
 */
export function visitNode<T>(node: ASTNode, visitor: ASTVisitor<T>): T | undefined {
  switch (node.type) {
    case 'value':
      return visitor.visitValue?.(node as ValueNode);
    case 'reference':
      return visitor.visitReference?.(node as ReferenceNode);
    case 'range':
      return visitor.visitRange?.(node as RangeNode);
    case 'function':
      return visitor.visitFunction?.(node as FunctionNode);
    case 'unary-op':
      return visitor.visitUnaryOp?.(node as UnaryOpNode);
    case 'binary-op':
      return visitor.visitBinaryOp?.(node as BinaryOpNode);
    case 'array':
      return visitor.visitArray?.(node as ArrayNode);
    case 'named-expression':
      return visitor.visitNamedExpression?.(node as NamedExpressionNode);
    case 'error':
      return visitor.visitError?.(node as ErrorNode);
  }
}

/**
 * Traverse an AST tree depth-first
 */
export function traverseAST(node: ASTNode, visitor: ASTVisitor<void>): void {
  visitNode(node, visitor);
  
  switch (node.type) {
    case 'function':
      (node as FunctionNode).args.forEach(arg => traverseAST(arg, visitor));
      break;
    case 'unary-op':
      traverseAST((node as UnaryOpNode).operand, visitor);
      break;
    case 'binary-op':
      traverseAST((node as BinaryOpNode).left, visitor);
      traverseAST((node as BinaryOpNode).right, visitor);
      break;
    case 'array':
      (node as ArrayNode).elements.forEach(row => 
        row.forEach(element => traverseAST(element, visitor))
      );
      break;
  }
}

/**
 * Check if an AST node is a constant value (no references)
 */
export function isConstantNode(node: ASTNode): boolean {
  switch (node.type) {
    case 'value':
      return true;
    case 'reference':
    case 'range':
    case 'named-expression':
      return false;
    case 'function':
      return (node as FunctionNode).args.every(isConstantNode);
    case 'unary-op':
      return isConstantNode((node as UnaryOpNode).operand);
    case 'binary-op':
      return isConstantNode((node as BinaryOpNode).left) && isConstantNode((node as BinaryOpNode).right);
    case 'array':
      return (node as ArrayNode).elements.every(row => row.every(isConstantNode));
    case 'error':
      return true;
  }
}

/**
 * Get all cell references from an AST
 */
export function getCellReferences(node: ASTNode): Array<SimpleCellAddress | SimpleCellRange> {
  const references: Array<SimpleCellAddress | SimpleCellRange> = [];
  
  traverseAST(node, {
    visitReference(node) {
      references.push(node.address);
    },
    visitRange(node) {
      references.push(node.range);
    }
  });
  
  return references;
}

/**
 * Get all named expression references from an AST
 */
export function getNamedExpressionReferences(node: ASTNode): string[] {
  const names: string[] = [];
  
  traverseAST(node, {
    visitNamedExpression(node) {
      names.push(node.name);
    }
  });
  
  return names;
}

/**
 * Clone an AST node (deep copy)
 */
export function cloneAST(node: ASTNode): ASTNode {
  switch (node.type) {
    case 'value':
      return { ...(node as ValueNode) } as ValueNode;
    case 'reference': {
      const refNode = node as ReferenceNode;
      return { ...refNode, address: { ...refNode.address } } as ReferenceNode;
    }
    case 'range': {
      const rangeNode = node as RangeNode;
      return {
        ...rangeNode,
        range: {
          start: { ...rangeNode.range.start },
          end: { ...rangeNode.range.end }
        }
      } as RangeNode;
    }
    case 'function': {
      const funcNode = node as FunctionNode;
      return {
        ...funcNode,
        args: funcNode.args.map(cloneAST)
      } as FunctionNode;
    }
    case 'unary-op': {
      const unaryNode = node as UnaryOpNode;
      return {
        ...unaryNode,
        operand: cloneAST(unaryNode.operand)
      } as UnaryOpNode;
    }
    case 'binary-op': {
      const binaryNode = node as BinaryOpNode;
      return {
        ...binaryNode,
        left: cloneAST(binaryNode.left),
        right: cloneAST(binaryNode.right)
      } as BinaryOpNode;
    }
    case 'array': {
      const arrayNode = node as ArrayNode;
      return {
        ...arrayNode,
        elements: arrayNode.elements.map(row => row.map(cloneAST))
      } as ArrayNode;
    }
    case 'named-expression':
      return { ...(node as NamedExpressionNode) } as NamedExpressionNode;
    case 'error':
      return { ...(node as ErrorNode) } as ErrorNode;
  }
}
