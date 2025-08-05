import { test, expect, describe } from "bun:test";
import {
  createValueNode,
  createReferenceNode,
  createRangeNode,
  createFunctionNode,
  createUnaryOpNode,
  createBinaryOpNode,
  createArrayNode,
  createNamedExpressionNode,
  createErrorNode,
  visitNode,
  traverseAST,
  isConstantNode,
  getCellReferences,
  getNamedExpressionReferences,
  cloneAST,
  type ASTNode,
  type ValueNode,
  type ReferenceNode,
  type FunctionNode,
  type BinaryOpNode,
  type ASTVisitor
} from '../../../src/parser/ast';

describe('AST Node Creation', () => {
  test('should create value nodes correctly', () => {
    const numberNode = createValueNode(42, 0, 2);
    expect(numberNode.type).toBe('value');
    expect(numberNode.value).toBe(42);
    expect(numberNode.valueType).toBe('number');
    expect(numberNode.position).toEqual({ start: 0, end: 2 });

    const stringNode = createValueNode('hello');
    expect(stringNode.valueType).toBe('string');
    expect(stringNode.value).toBe('hello');

    const boolNode = createValueNode(true);
    expect(boolNode.valueType).toBe('boolean');
    expect(boolNode.value).toBe(true);

    const errorNode = createValueNode('#DIV/0!');
    expect(errorNode.valueType).toBe('error');
    expect(errorNode.value).toBe('#DIV/0!');
  });

  test('should create reference nodes correctly', () => {
    const refNode = createReferenceNode(
      { sheet: 0, col: 1, row: 2 },
      { col: false, row: true },
      0, 4
    );
    expect(refNode.type).toBe('reference');
    expect(refNode.address).toEqual({ sheet: 0, col: 1, row: 2 });
    expect(refNode.isAbsolute).toEqual({ col: false, row: true });
  });

  test('should create range nodes correctly', () => {
    const rangeNode = createRangeNode(
      {
        start: { sheet: 0, col: 0, row: 0 },
        end: { sheet: 0, col: 2, row: 9 }
      },
      {
        start: { col: true, row: true },
        end: { col: false, row: false }
      },
      0, 8
    );
    expect(rangeNode.type).toBe('range');
    expect(rangeNode.range.start).toEqual({ sheet: 0, col: 0, row: 0 });
    expect(rangeNode.range.end).toEqual({ sheet: 0, col: 2, row: 9 });
  });

  test('should create function nodes correctly', () => {
    const sumNode = createFunctionNode(
      'SUM',
      [createValueNode(1), createValueNode(2)],
      0, 10
    );
    expect(sumNode.type).toBe('function');
    expect(sumNode.name).toBe('SUM'); // Should be uppercase
    expect(sumNode.args).toHaveLength(2);
  });

  test('should create unary operator nodes correctly', () => {
    const negNode = createUnaryOpNode(
      '-',
      createValueNode(5),
      0, 2
    );
    expect(negNode.type).toBe('unary-op');
    expect(negNode.operator).toBe('-');
    expect((negNode.operand as ValueNode).value).toBe(5);
  });

  test('should create binary operator nodes correctly', () => {
    const addNode = createBinaryOpNode(
      '+',
      createValueNode(1),
      createValueNode(2),
      0, 5
    );
    expect(addNode.type).toBe('binary-op');
    expect(addNode.operator).toBe('+');
    expect((addNode.left as ValueNode).value).toBe(1);
    expect((addNode.right as ValueNode).value).toBe(2);
  });

  test('should create array nodes correctly', () => {
    const arrayNode = createArrayNode(
      [
        [createValueNode(1), createValueNode(2)],
        [createValueNode(3), createValueNode(4)]
      ],
      0, 15
    );
    expect(arrayNode.type).toBe('array');
    expect(arrayNode.elements).toHaveLength(2);
    expect(arrayNode.elements[0]).toHaveLength(2);
  });

  test('should create named expression nodes correctly', () => {
    const namedNode = createNamedExpressionNode('TaxRate', undefined, 0, 7);
    expect(namedNode.type).toBe('named-expression');
    expect(namedNode.name).toBe('TaxRate');
    expect(namedNode.scope).toBeUndefined();

    const scopedNode = createNamedExpressionNode('LocalRate', 2);
    expect(scopedNode.scope).toBe(2);
  });

  test('should create error nodes correctly', () => {
    const errorNode = createErrorNode('#REF!', 'Invalid reference', 0, 5);
    expect(errorNode.type).toBe('error');
    expect(errorNode.error).toBe('#REF!');
    expect(errorNode.message).toBe('Invalid reference');
  });
});

describe('AST Visitor Pattern', () => {
  test('should visit nodes with visitor pattern', () => {
    const visited: string[] = [];
    const visitor: ASTVisitor<void> = {
      visitValue: (node) => { visited.push(`value:${node.value}`); },
      visitFunction: (node) => { visited.push(`function:${node.name}`); },
      visitBinaryOp: (node) => { visited.push(`binary:${node.operator}`); }
    };

    const addNode = createBinaryOpNode(
      '+',
      createValueNode(1),
      createValueNode(2)
    );

    visitNode(addNode, visitor);
    expect(visited).toEqual(['binary:+']);

    visited.length = 0;
    visitNode(createValueNode(42), visitor);
    expect(visited).toEqual(['value:42']);
  });

  test('should traverse AST depth-first', () => {
    const visited: string[] = [];
    const visitor: ASTVisitor<void> = {
      visitValue: (node) => { visited.push(`${node.value}`); },
      visitBinaryOp: (node) => { visited.push(node.operator); }
    };

    // Create expression: (1 + 2) * 3
    const addNode = createBinaryOpNode('+', createValueNode(1), createValueNode(2));
    const mulNode = createBinaryOpNode('*', addNode, createValueNode(3));

    traverseAST(mulNode, visitor);
    expect(visited).toEqual(['*', '+', '1', '2', '3']);
  });
});

describe('AST Analysis Functions', () => {
  test('should identify constant nodes correctly', () => {
    expect(isConstantNode(createValueNode(42))).toBe(true);
    expect(isConstantNode(createErrorNode('#DIV/0!', 'Division by zero'))).toBe(true);
    
    const constExpr = createBinaryOpNode(
      '+',
      createValueNode(1),
      createValueNode(2)
    );
    expect(isConstantNode(constExpr)).toBe(true);

    const nonConstExpr = createBinaryOpNode(
      '+',
      createReferenceNode({ sheet: 0, col: 0, row: 0 }, { col: false, row: false }),
      createValueNode(2)
    );
    expect(isConstantNode(nonConstExpr)).toBe(false);
  });

  test('should extract cell references from AST', () => {
    // Create expression: A1 + SUM(B1:B10)
    const a1Ref = createReferenceNode(
      { sheet: 0, col: 0, row: 0 },
      { col: false, row: false }
    );
    const b1b10Range = createRangeNode(
      {
        start: { sheet: 0, col: 1, row: 0 },
        end: { sheet: 0, col: 1, row: 9 }
      },
      {
        start: { col: false, row: false },
        end: { col: false, row: false }
      }
    );
    const sumNode = createFunctionNode('SUM', [b1b10Range]);
    const addNode = createBinaryOpNode('+', a1Ref, sumNode);

    const refs = getCellReferences(addNode);
    expect(refs).toHaveLength(2);
    expect(refs[0]).toEqual({ sheet: 0, col: 0, row: 0 });
    expect(refs[1]).toEqual({
      start: { sheet: 0, col: 1, row: 0 },
      end: { sheet: 0, col: 1, row: 9 }
    });
  });

  test('should extract named expression references from AST', () => {
    // Create expression: TaxRate * (SubTotal + LocalRate)
    const taxRate = createNamedExpressionNode('TaxRate');
    const subTotal = createNamedExpressionNode('SubTotal');
    const localRate = createNamedExpressionNode('LocalRate', 1);
    
    const addNode = createBinaryOpNode('+', subTotal, localRate);
    const mulNode = createBinaryOpNode('*', taxRate, addNode);

    const names = getNamedExpressionReferences(mulNode);
    expect(names).toEqual(['TaxRate', 'SubTotal', 'LocalRate']);
  });
});

describe('AST Cloning', () => {
  test('should deep clone value nodes', () => {
    const original = createValueNode(42, 0, 2);
    const cloned = cloneAST(original) as ValueNode;
    
    expect(cloned).not.toBe(original);
    expect(cloned.value).toBe(original.value as string | number | boolean);
    expect(cloned.valueType).toBe(original.valueType);
  });

  test('should deep clone reference nodes', () => {
    const original = createReferenceNode(
      { sheet: 0, col: 1, row: 2 },
      { col: true, row: false }
    );
    const cloned = cloneAST(original) as ReferenceNode;
    
    expect(cloned).not.toBe(original);
    expect(cloned.address).not.toBe(original.address);
    expect(cloned.address).toEqual(original.address);
  });

  test('should deep clone complex expressions', () => {
    // Create expression: SUM(A1:A10) + COUNT(B1:B10)
    const a1a10 = createRangeNode(
      {
        start: { sheet: 0, col: 0, row: 0 },
        end: { sheet: 0, col: 0, row: 9 }
      },
      {
        start: { col: false, row: false },
        end: { col: false, row: false }
      }
    );
    const b1b10 = createRangeNode(
      {
        start: { sheet: 0, col: 1, row: 0 },
        end: { sheet: 0, col: 1, row: 9 }
      },
      {
        start: { col: false, row: false },
        end: { col: false, row: false }
      }
    );
    
    const sumNode = createFunctionNode('SUM', [a1a10]);
    const countNode = createFunctionNode('COUNT', [b1b10]);
    const addNode = createBinaryOpNode('+', sumNode, countNode);
    
    const cloned = cloneAST(addNode) as BinaryOpNode;
    
    expect(cloned).not.toBe(addNode);
    expect(cloned.left).not.toBe(addNode.left);
    expect(cloned.right).not.toBe(addNode.right);
    expect((cloned.left as FunctionNode).name).toBe('SUM');
    expect((cloned.right as FunctionNode).name).toBe('COUNT');
  });

  test('should clone array nodes correctly', () => {
    const original = createArrayNode([
      [createValueNode(1), createValueNode(2)],
      [createValueNode(3), createValueNode(4)]
    ]);
    
    const cloned = cloneAST(original);
    expect(cloned).not.toBe(original);
    expect(cloned.type).toBe('array');
  });
});