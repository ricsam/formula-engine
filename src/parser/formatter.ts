import { indexToColumn } from "src/core/utils";
import { type CellValue, type SerializedCellValue } from "../core/types";
import { parseOpenEndedRange, parseInfiniteRange } from "./grammar";
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
} from "./ast";
import { getOperatorPrecedence, getOperatorAssociativity } from "./grammar";
import { parseFormula } from "./parser";

export function astToString(ast: ASTNode): string {
  switch (ast.type) {
    case "value":
      return formatValue(ast.value);

    case "reference":
      return formatReference(ast);

    case "range":
      return formatRange(ast);

    case "function":
      return formatFunction(ast);

    case "unary-op":
      return formatUnaryOp(ast);

    case "binary-op":
      return formatBinaryOp(ast);

    case "array":
      return formatArray(ast);

    case "named-expression":
      return formatNamedExpression(ast);

    case "error":
      return ast.error;

    case "empty":
      return "";

    case "3d-range":
      return formatThreeDRange(ast);

    case "structured-reference":
      return formatStructuredReference(ast);

    case "infinity":
      return "INFINITY";

    default:
      throw new Error(`Unknown AST node type: ${(ast as any).type}`);
  }
}

function formatValue(value: CellValue): string {
  if (value.type === "string") {
    // Escape quotes by doubling them and wrap in quotes
    return `"${value.value.replace(/"/g, '""')}"`;
  } else if (value.type === "boolean") {
    return value.value ? "TRUE" : "FALSE";
  } else if (value.type === "number") {
    return value.value.toString();
  }
  return String(value);
}

function formatReference(ast: ReferenceNode): string {
  const { address, isAbsolute, sheetName, workbookName } = ast;
  const colLetter = indexToColumn(address.colIndex);
  const rowNumber = address.rowIndex + 1; // Convert from 0-based to 1-based

  const colRef = isAbsolute.col ? `$${colLetter}` : colLetter;
  const rowRef = isAbsolute.row ? `$${rowNumber}` : rowNumber.toString();

  const cellRef = `${colRef}${rowRef}`;

  let result = cellRef;

  if (sheetName) {
    const quotedSheet = sheetName.includes(" ") ? `'${sheetName}'` : sheetName;
    result = `${quotedSheet}!${cellRef}`;
  }

  if (workbookName) {
    result = `[${workbookName}]${result}`;
  }

  return result;
}

function formatRange(ast: RangeNode): string {
  const { range, isAbsolute, sheetName, workbookName } = ast;

  // Handle infinite ranges
  if (range.end.col.type === "infinity" || range.end.row.type === "infinity") {
    return formatInfiniteRange(ast);
  }

  // Regular range
  const startCol = indexToColumn(range.start.col);
  const startRow = range.start.row + 1;
  const endCol = indexToColumn(range.end.col.value);
  const endRow = range.end.row.value + 1;

  const startColRef = isAbsolute.start.col ? `$${startCol}` : startCol;
  const startRowRef = isAbsolute.start.row
    ? `$${startRow}`
    : startRow.toString();
  const endColRef = isAbsolute.end.col ? `$${endCol}` : endCol;
  const endRowRef = isAbsolute.end.row ? `$${endRow}` : endRow.toString();

  const rangeRef = `${startColRef}${startRowRef}:${endColRef}${endRowRef}`;

  let result = rangeRef;

  if (sheetName) {
    const quotedSheet = sheetName.includes(" ") ? `'${sheetName}'` : sheetName;
    result = `${quotedSheet}!${rangeRef}`;
  }

  if (workbookName) {
    result = `[${workbookName}]${result}`;
  }

  return result;
}

function formatInfiniteRange(ast: RangeNode): string {
  const { range, isAbsolute, sheetName, workbookName } = ast;

  let rangeRef: string;

  // Canonical formatting rules:
  // 1. Start is always explicit cell (A5, $A$5, etc.)
  // 2. End varies by openness:
  //    - Open→: A5:10 (row-bounded)
  //    - Open↓: A5:D (column-bounded)  
  //    - Open both: A5:INFINITY
  // 3. Excel compatibility: whole rows/columns use canonical cell form
  //    - 5:5 → A5:5, A:A → A1:A

  const startCol = indexToColumn(range.start.col);
  const startRow = range.start.row + 1;

  const startColRef = isAbsolute.start.col ? `$${startCol}` : startCol;
  const startRowRef = isAbsolute.start.row ? `$${startRow}` : startRow.toString();

  if (range.end.col.type === "infinity" && range.end.row.type === "infinity") {
    // Open both: A5:INFINITY
    rangeRef = `${startColRef}${startRowRef}:INFINITY`;
  } else if (range.end.col.type === "infinity") {
    // Open→ (row-bounded): A5:10
    if (range.end.row.type !== "number") {
      throw new Error("Expected finite row for infinite column range");
    }
    
    const endRow = range.end.row.value + 1;
    const endRowRef = isAbsolute.end.row ? `$${endRow}` : endRow.toString();
    
    // Always use canonical cell form: A5:10 (not 5:10)
    rangeRef = `${startColRef}${startRowRef}:${endRowRef}`;
  } else if (range.end.row.type === "infinity") {
    // Open↓ (column-bounded): A5:D
    if (range.end.col.type !== "number") {
      throw new Error("Expected finite column for infinite row range");
    }

    const endCol = indexToColumn(range.end.col.value);
    const endColRef = isAbsolute.end.col ? `$${endCol}` : endCol;
    
    // Always use canonical cell form: A1:D (not A:D)
    rangeRef = `${startColRef}${startRowRef}:${endColRef}`;
  } else {
    throw new Error("Expected at least one infinite dimension for infinite range");
  }

  let result = rangeRef;

  if (sheetName) {
    const quotedSheet = sheetName.includes(" ") ? `'${sheetName}'` : sheetName;
    result = `${quotedSheet}!${rangeRef}`;
  }

  if (workbookName) {
    result = `[${workbookName}]${result}`;
  }

  return result;
}

function formatFunction(ast: FunctionNode): string {
  const { name, args } = ast;
  const formattedArgs = args.map((arg: ASTNode) => astToString(arg)).join(",");
  return `${name}(${formattedArgs})`;
}

function formatUnaryOp(ast: UnaryOpNode): string {
  const { operator, operand } = ast;
  const operandStr = astToString(operand);

  if (operator === "%") {
    return `${operandStr}%`;
  } else {
    return `${operator}${operandStr}`;
  }
}

function formatBinaryOp(ast: BinaryOpNode): string {
  const { operator, left, right } = ast;

  // Determine if we need parentheses for left and right operands
  const leftNeedsParens = needsParentheses(left, ast, "left");
  const rightNeedsParens = needsParentheses(right, ast, "right");

  const leftStr = leftNeedsParens
    ? `(${astToString(left)})`
    : astToString(left);
  const rightStr = rightNeedsParens
    ? `(${astToString(right)})`
    : astToString(right);

  return `${leftStr}${operator}${rightStr}`;
}

function needsParentheses(
  child: ASTNode,
  parent: ASTNode,
  position: "left" | "right"
): boolean {
  if (child.type !== "binary-op") {
    return false;
  }

  // Only binary operations need precedence checking
  if (parent.type !== "binary-op") {
    return false;
  }

  const childPrecedence = getOperatorPrecedence(child.operator);
  const parentPrecedence = getOperatorPrecedence(parent.operator);

  // Lower precedence always needs parentheses
  if (childPrecedence < parentPrecedence) {
    return true;
  }

  // For same precedence, check associativity
  if (childPrecedence === parentPrecedence) {
    const associativity = getOperatorAssociativity(parent.operator);

    // For left-associative operators, right side needs parentheses
    // For right-associative operators, left side needs parentheses
    if (associativity === "left" && position === "right") {
      return true;
    } else if (associativity === "right" && position === "left") {
      return true;
    }
  }

  return false;
}

function formatArray(ast: ArrayNode): string {
  const { elements } = ast;
  const rows = elements.map((row: ASTNode[]) =>
    row.map((cell: ASTNode) => astToString(cell)).join(",")
  );
  return `{${rows.join(";")}}`;
}

function formatNamedExpression(ast: NamedExpressionNode): string {
  const { name, sheetName, workbookName } = ast;
  
  let result = name;
  
  if (sheetName !== undefined) {
    // Sheet-scoped named expression
    const quotedSheet = sheetName.includes(" ") ? `'${sheetName}'` : sheetName;
    result = `${quotedSheet}!${name}`;
  }
  
  if (workbookName) {
    result = `[${workbookName}]${result}`;
  }
  
  return result;
}

function formatThreeDRange(ast: ThreeDRangeNode): string {
  const { startSheet, endSheet, workbookName, reference } = ast;
  const refStr = astToString(reference);

  // Remove sheet name and workbook name from reference if present (since we're adding the 3D range prefix)
  let cleanRef = refStr;
  if (cleanRef.includes("!")) {
    cleanRef = cleanRef.split("!")[1]!;
  }
  // Remove workbook prefix if present
  if (cleanRef.startsWith("[") && cleanRef.includes("]")) {
    const bracketEnd = cleanRef.indexOf("]");
    cleanRef = cleanRef.substring(bracketEnd + 1);
  }

  const quotedStartSheet = startSheet.includes(" ")
    ? `'${startSheet}'`
    : startSheet;
  const quotedEndSheet = endSheet.includes(" ") ? `'${endSheet}'` : endSheet;

  let result = `${quotedStartSheet}:${quotedEndSheet}!${cleanRef}`;

  if (workbookName) {
    result = `[${workbookName}]${result}`;
  }

  return result;
}

/**
 * Check if a column name contains special characters that require double bracket syntax
 */
function needsColumnBrackets(columnName: string): boolean {
  // Column names need extra brackets if they contain spaces or special characters
  return /[\s\[\]#@,:=]/.test(columnName);
}

function formatStructuredReference(ast: StructuredReferenceNode): string {
  const { tableName, sheetName, workbookName, cols, selector, isCurrentRow } = ast;

  if (!tableName && isCurrentRow) {
    // Current row reference like [@Column] or @Column
    if (cols && cols.startCol) {
      return `[@${cols.startCol}]`;
    }
    return "@";
  }

  if (!tableName && !isCurrentRow) {
    // Bare column reference like [Column] or [#Data]
    if (selector) {
      return `[${selector}]`;
    } else if (cols) {
      const startNeedsBrackets = needsColumnBrackets(cols.startCol);
      const endNeedsBrackets = cols.startCol !== cols.endCol && needsColumnBrackets(cols.endCol);
      const anyNeedsBrackets = startNeedsBrackets || endNeedsBrackets;
      
      if (cols.startCol === cols.endCol) {
        // Single column
        if (startNeedsBrackets) {
          return `[[${cols.startCol}]]`;
        } else {
          return `[${cols.startCol}]`;
        }
      } else {
        // Column range
        if (anyNeedsBrackets) {
          return `[[${cols.startCol}]:[${cols.endCol}]]`;
        } else {
          return `[${cols.startCol}:${cols.endCol}]`;
        }
      }
    }
    return "[]"; // Empty bare reference (shouldn't happen)
  }

  let result = "";

  if (workbookName) {
    result += `[${workbookName}]`;
  }

  if (sheetName) {
    const quotedSheet = sheetName.includes(" ") ? `'${sheetName}'` : sheetName;
    result += `${quotedSheet}!`;
  }

  result += tableName;

  if (selector && cols) {
    // Complex syntax like Table1[[#Headers],[Column1]] or Table1[[#Headers],[Col1:Col2]]
    const colRef =
      cols.startCol === cols.endCol
        ? cols.startCol
        : `${cols.startCol}:${cols.endCol}`;
    result += `[[${selector}],[${colRef}]]`;
  } else if (selector) {
    // Simple selector like Table1[#Data]
    result += `[${selector}]`;
  } else if (cols) {
    const startNeedsBrackets = needsColumnBrackets(cols.startCol);
    const endNeedsBrackets = cols.startCol !== cols.endCol && needsColumnBrackets(cols.endCol);
    const anyNeedsBrackets = startNeedsBrackets || endNeedsBrackets;
    
    if (isCurrentRow) {
      // Current row references
      if (cols.startCol === cols.endCol) {
        // Single column
        if (startNeedsBrackets) {
          result += `[@[${cols.startCol}]]`;
        } else {
          result += `[@${cols.startCol}]`;
        }
      } else {
        // Column range
        if (anyNeedsBrackets) {
          result += `[@[${cols.startCol}]:[${cols.endCol}]]`;
        } else {
          result += `[@${cols.startCol}:${cols.endCol}]`;
        }
      }
    } else {
      // Regular column references
      if (cols.startCol === cols.endCol) {
        // Single column - always use single brackets for table references
        result += `[${cols.startCol}]`;
      } else {
        // Column range
        if (anyNeedsBrackets) {
          result += `[[${cols.startCol}]:[${cols.endCol}]]`;
        } else {
          result += `[${cols.startCol}:${cols.endCol}]`;
        }
      }
    }
  }

  return result;
}

export function formatFormula(formula: string): string {
  return astToString(parseFormula(formula));
}

export function normalizeSerializedCellValue(
  value: SerializedCellValue
): SerializedCellValue {
  if (value === undefined || value === "") return undefined;

  if (typeof value === "string" && value.startsWith("=")) {
    return `=${formatFormula(value.slice(1))}`;
  }

  return value;
}

export function isSerializedCellValueEqual(
  a: SerializedCellValue,
  b: SerializedCellValue
): boolean {
  const normalizedA = normalizeSerializedCellValue(a);
  const normalizedB = normalizeSerializedCellValue(b);

  return normalizedA === normalizedB;
}
