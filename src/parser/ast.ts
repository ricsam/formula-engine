/**
 * Abstract Syntax Tree (AST) node definitions for formula parsing
 */

import type {
  CellAddress,
  CellValue,
  FormulaError,
  SpreadsheetRange,
} from "../core/types";

/**
 * Base interface for all AST nodes
 */
type ASTNodeBase = {
  position?: {
    start: number;
    end: number;
  };
};

/**
 * Literal value node (number, string, boolean, error)
 */
export type ValueNode = ASTNodeBase & {
  type: "value";
  value: CellValue;
};

/**
 * Cell reference node (e.g., A1, Sheet1!B2)
 */
export type ReferenceNode = ASTNodeBase & {
  type: "reference";
  address: {
    colIndex: number;
    rowIndex: number;
  };
  sheetName?: string;
  workbookName?: string;
  isAbsolute: {
    col: boolean;
    row: boolean;
  };
};

/**
 * Range reference node (e.g., A1:B10)
 */
export type RangeNode = ASTNodeBase & {
  type: "range";
  sheetName?: string;
  workbookName?: string;
  range: SpreadsheetRange;
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
};

/**
 * Function call node (e.g., SUM(A1:A10))
 */
export type FunctionNode = ASTNodeBase & {
  type: "function";
  name: string;
  args: ASTNode[];
};

/**
 * Unary operator node (e.g., -A1, +B2)
 */
export type UnaryOpNode = ASTNodeBase & {
  type: "unary-op";
  operator: "+" | "-" | "%";
  operand: ASTNode;
};

/**
 * Binary operator node (e.g., A1+B1, C1*D1)
 */
export type BinaryOpNode = ASTNodeBase & {
  type: "binary-op";
  operator:
    | "+"
    | "-"
    | "*"
    | "/"
    | "^"
    | "&"
    | "="
    | "<>"
    | "<"
    | ">"
    | "<="
    | ">=";
  left: ASTNode;
  right: ASTNode;
};

/**
 * Array literal node (e.g., {1,2,3;4,5,6})
 */
export type ArrayNode = ASTNodeBase & {
  type: "array";
  elements: ASTNode[][]; // 2D array of elements
};

/**
 * 3D range node (e.g., Sheet1:Sheet3!A1)
 */
export type ThreeDRangeNode = ASTNodeBase & {
  type: "3d-range";
  startSheet: string;
  endSheet: string;
  workbookName?: string;
  reference: ReferenceNode | RangeNode;
};

/**
 * Structured reference node (e.g., Table1[Column1])
 */
export type StructuredReferenceNode = ASTNodeBase & {
  type: "structured-reference";
  tableName?: string;
  sheetName?: string;
  workbookName?: string;
  cols?: {
    startCol: string;
    endCol: string;
  };
  selector?: "#All" | "#Data" | "#Headers";
  isCurrentRow: boolean;
};

/**
 * Infinity literal node
 */
export type InfinityNode = ASTNodeBase & {
  type: "infinity";
};

export type ASTNode =
  | ValueNode
  | ReferenceNode
  | RangeNode
  | FunctionNode
  | UnaryOpNode
  | BinaryOpNode
  | ArrayNode
  | NamedExpressionNode
  | ErrorNode
  | EmptyNode
  | ThreeDRangeNode
  | StructuredReferenceNode
  | InfinityNode;

/**
 * Named expression reference node
 */
export type NamedExpressionNode = ASTNodeBase & {
  type: "named-expression";
  name: string;
  sheetName?: string;
  workbookName?: string;
};

/**
 * Error node for parsing errors
 */
export type ErrorNode = ASTNodeBase & {
  type: "error";
  error: FormulaError;
  message: string;
};

export type EmptyNode = ASTNodeBase & {
  type: "empty";
};

/**
 * Helper function to create a value node
 */
export function createValueNode(
  value: CellValue,
  position?: {
    start: number;
    end: number;
  }
): ValueNode {
  return {
    type: "value",
    value,
    position,
  };
}

export function createEmptyNode(position?: {
  start: number;
  end: number;
}): EmptyNode {
  return { type: "empty", position };
}

/**
 * Helper function to create a reference node
 */
export function createReferenceNode({
  address,
  isAbsolute,
  position,
  sheetName,
  workbookName,
}: {
  address: {
    colIndex: number;
    rowIndex: number;
  };
  isAbsolute: { col: boolean; row: boolean };
  position?: {
    start: number;
    end: number;
  };
  sheetName?: string;
  workbookName?: string;
}): ReferenceNode {
  return {
    type: "reference",
    address,
    isAbsolute,
    sheetName,
    workbookName,
    position,
  };
}

/**
 * Helper function to create a range node
 */
export function createRangeNode({
  sheetName,
  workbookName,
  range,
  isAbsolute,
  position,
}: {
  sheetName?: string;
  workbookName?: string;
  range: SpreadsheetRange;
  isAbsolute: RangeNode["isAbsolute"];
  position?: {
    start: number;
    end: number;
  };
}): RangeNode {
  return {
    type: "range",
    range,
    sheetName,
    workbookName,
    isAbsolute,
    position,
  };
}

/**
 * Helper function to create a function node
 */
export function createFunctionNode(
  name: string,
  args: ASTNode[],
  position?: {
    start: number;
    end: number;
  }
): FunctionNode {
  return {
    type: "function",
    name: name.toUpperCase(), // Normalize function names to uppercase
    args,
    position,
  };
}

/**
 * Helper function to create a unary operator node
 */
export function createUnaryOpNode(
  operator: UnaryOpNode["operator"],
  operand: ASTNode,
  position?: {
    start: number;
    end: number;
  }
): UnaryOpNode {
  return {
    type: "unary-op",
    operator,
    operand,
    position,
  };
}

/**
 * Helper function to create a binary operator node
 */
export function createBinaryOpNode(
  operator: BinaryOpNode["operator"],
  left: ASTNode,
  right: ASTNode,
  position?: {
    start: number;
    end: number;
  }
): BinaryOpNode {
  return {
    type: "binary-op",
    operator,
    left,
    right,
    position,
  };
}

/**
 * Helper function to create an array node
 */
export function createArrayNode(
  elements: ASTNode[][],
  position?: {
    start: number;
    end: number;
  }
): ArrayNode {
  return {
    type: "array",
    elements,
    position,
  };
}

/**
 * Helper function to create a named expression node
 */
export function createNamedExpressionNode(
  name: string,
  position?: {
    start: number;
    end: number;
  },
  sheetName?: string,
  workbookName?: string
): NamedExpressionNode {
  return {
    type: "named-expression",
    name,
    sheetName,
    workbookName,
    position,
  };
}

/**
 * Helper function to create an error node
 */
export function createErrorNode(
  error: FormulaError,
  message: string,
  position?: {
    start: number;
    end: number;
  }
): ErrorNode {
  return {
    type: "error",
    error,
    message,
    position,
  };
}

/**
 * Helper function to create a 3D range node
 */
export function createThreeDRangeNode(
  startSheet: string,
  endSheet: string,
  reference: ReferenceNode | RangeNode,
  position?: {
    start: number;
    end: number;
  },
  workbookName?: string
): ThreeDRangeNode {
  return {
    type: "3d-range",
    startSheet,
    endSheet,
    workbookName,
    reference,
    position,
  };
}

/**
 * Helper function to create a structured reference node
 */
export function createStructuredReferenceNode({
  tableName,
  sheetName,
  workbookName,
  cols,
  selector,
  isCurrentRow = false,
  position,
}: {
  tableName?: string;
  sheetName?: string;
  workbookName?: string;
  cols?: {
    startCol: string;
    endCol: string;
  };
  selector?: "#All" | "#Data" | "#Headers";
  isCurrentRow?: boolean;
  position?: {
    start: number;
    end: number;
  };
}): StructuredReferenceNode {
  return {
    type: "structured-reference",
    tableName,
    sheetName,
    workbookName,
    cols,
    selector,
    isCurrentRow,
    position,
  };
}

/**
 * Helper function to create an infinity node
 */
export function createInfinityNode(position?: {
  start: number;
  end: number;
}): InfinityNode {
  return {
    type: "infinity",
    position,
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
  visitEmpty?(node: EmptyNode): T;
  visitThreeDRange?(node: ThreeDRangeNode): T;
  visitStructuredReference?(node: StructuredReferenceNode): T;
  visitInfinity?(node: InfinityNode): T;
}
