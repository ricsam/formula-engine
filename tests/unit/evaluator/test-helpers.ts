/**
 * Test helpers for creating AST nodes
 */

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
} from '../../../src/parser/ast';
import type { CellValue, SimpleCellAddress, SimpleCellRange } from '../../../src/core/types';

export function createTestValueNode(
  value: CellValue, 
  valueType: 'number' | 'string' | 'boolean' | 'error' = 'number'
): ValueNode {
  return {
    type: 'value',
    value,
    valueType,
    position: { start: 0, end: 1 }
  };
}

export function createTestReferenceNode(
  address: SimpleCellAddress,
  isAbsolute: { col: boolean; row: boolean } = { col: false, row: false }
): ReferenceNode {
  return {
    type: 'reference',
    address,
    isAbsolute,
    position: { start: 0, end: 1 }
  };
}

export function createTestRangeNode(
  range: SimpleCellRange,
  isAbsolute?: {
    start: { col: boolean; row: boolean };
    end: { col: boolean; row: boolean };
  }
): RangeNode {
  return {
    type: 'range',
    range,
    isAbsolute: isAbsolute || {
      start: { col: false, row: false },
      end: { col: false, row: false }
    },
    position: { start: 0, end: 1 }
  };
}

export function createTestFunctionNode(
  name: string,
  args: ASTNode[] = []
): FunctionNode {
  return {
    type: 'function',
    name,
    args,
    position: { start: 0, end: 1 }
  };
}

export function createTestBinaryOpNode(
  operator: BinaryOpNode['operator'],
  left: ASTNode,
  right: ASTNode
): BinaryOpNode {
  return {
    type: 'binary-op',
    operator,
    left,
    right,
    position: { start: 0, end: 1 }
  };
}

export function createTestUnaryOpNode(
  operator: UnaryOpNode['operator'],
  operand: ASTNode
): UnaryOpNode {
  return {
    type: 'unary-op',
    operator,
    operand,
    position: { start: 0, end: 1 }
  };
}

export function createTestArrayNode(
  elements: ASTNode[][]
): ArrayNode {
  return {
    type: 'array',
    elements,
    position: { start: 0, end: 1 }
  };
}

export function createTestNamedExpressionNode(
  name: string,
  scope?: number
): NamedExpressionNode {
  return {
    type: 'named-expression',
    name,
    scope,
    position: { start: 0, end: 1 }
  };
}

export function createTestErrorNode(
  error: ErrorNode['error'],
  message: string = 'Test error'
): ErrorNode {
  return {
    type: 'error',
    error,
    message,
    position: { start: 0, end: 1 }
  };
}