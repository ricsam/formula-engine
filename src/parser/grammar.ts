/**
 * Grammar rules and operator precedence for formula parsing
 */

import type { BinaryOpNode } from "./ast";

/**
 * Operator precedence levels (higher number = higher precedence)
 */
export const OPERATOR_PRECEDENCE: Record<string, number> = {
  // Comparison operators (lowest precedence)
  "=": 1,
  "<>": 1,
  "<": 1,
  ">": 1,
  "<=": 1,
  ">=": 1,

  // Concatenation
  "&": 2,

  // Addition and subtraction
  "+": 3,
  "-": 3,

  // Multiplication and division
  "*": 4,
  "/": 4,

  // Exponentiation (highest precedence)
  "^": 5,
};

/**
 * Operator associativity
 */
type Associativity = "left" | "right";

export const OPERATOR_ASSOCIATIVITY: Record<string, Associativity> = {
  "=": "left",
  "<>": "left",
  "<": "left",
  ">": "left",
  "<=": "left",
  ">=": "left",
  "&": "left",
  "+": "left",
  "-": "left",
  "*": "left",
  "/": "left",
  "^": "right", // Exponentiation is right-associative
};

/**
 * Check if a string is a valid binary operator
 */
export function isBinaryOperator(op: string): op is BinaryOpNode["operator"] {
  return op in OPERATOR_PRECEDENCE;
}

/**
 * Get operator precedence
 */
export function getOperatorPrecedence(op: string): number {
  return OPERATOR_PRECEDENCE[op] || 0;
}

/**
 * Get operator associativity
 */
export function getOperatorAssociativity(op: string): Associativity {
  return OPERATOR_ASSOCIATIVITY[op] || "left";
}

/**
 * Compare operator precedence
 * Returns:
 *  - positive if op1 has higher precedence than op2
 *  - negative if op1 has lower precedence than op2
 *  - 0 if they have the same precedence
 */
export function compareOperatorPrecedence(op1: string, op2: string): number {
  return getOperatorPrecedence(op1) - getOperatorPrecedence(op2);
}

/**
 * Built-in function names that don't require parentheses in some contexts
 */
export const SPECIAL_FUNCTIONS = new Set([
  "PI", // PI() can be written as PI
  "TRUE", // TRUE() can be written as TRUE
  "FALSE", // FALSE() can be written as FALSE
  "NA", // NA() can be written as NA
]);

/**
 * Special constants
 */
export const SPECIAL_CONSTANTS = new Set([
  "INFINITY", // Infinity literal
]);

/**
 * Functions that accept variable number of arguments
 */
export const VARIADIC_FUNCTIONS = new Set([
  "SUM",
  "PRODUCT",
  "COUNT",
  "MAX",
  "MIN",
  "AND",
  "OR",
  "XOR",
  "CONCATENATE",
  "CHOOSE",
  "IFS",
  "SWITCH",
]);

/**
 * Functions that require at least one argument
 */
export const REQUIRED_ARG_FUNCTIONS = new Set([
  "SUM",
  "PRODUCT",
  "COUNT",
  "MAX",
  "MIN",
  "AVERAGE",
  "MEDIAN",
  "STDEV",
  "VAR",
  "AND",
  "OR",
  "XOR",
  "CONCATENATE",
]);

/**
 * Reserved keywords that cannot be used as names
 */
export const RESERVED_KEYWORDS = new Set([
  "TRUE",
  "FALSE",
  "NULL",
  "DIV",
  "MOD",
  "AND",
  "OR",
  "NOT",
  "XOR",
]);

/**
 * Check if a name is a reserved keyword
 */
export function isReservedKeyword(name: string): boolean {
  return RESERVED_KEYWORDS.has(name.toUpperCase());
}

/**
 * Grammar production rules (in BNF-like notation):
 *
 * Formula ::= Expression
 *
 * Expression ::= ComparisonExpr
 *
 * ComparisonExpr ::= ConcatExpr (ComparisonOp ConcatExpr)*
 * ComparisonOp ::= '=' | '<>' | '<' | '>' | '<=' | '>='
 *
 * ConcatExpr ::= AddExpr ('&' AddExpr)*
 *
 * AddExpr ::= MultExpr (AddOp MultExpr)*
 * AddOp ::= '+' | '-'
 *
 * MultExpr ::= PowerExpr (MultOp PowerExpr)*
 * MultOp ::= '*' | '/'
 *
 * PowerExpr ::= UnaryExpr ('^' PowerExpr)?  // Right associative
 *
 * UnaryExpr ::= UnaryOp UnaryExpr | PostfixExpr
 * UnaryOp ::= '+' | '-'
 *
 * PostfixExpr ::= PrimaryExpr ('%')?
 *
 * PrimaryExpr ::= Number
 *              | String
 *              | Boolean
 *              | Error
 *              | CellReference
 *              | RangeReference
 *              | NamedExpression
 *              | FunctionCall
 *              | ArrayLiteral
 *              | '(' Expression ')'
 *
 * CellReference ::= (SheetName '!')? AbsoluteIndicator? Column AbsoluteIndicator? Row
 * RangeReference ::= CellReference ':' CellReference
 *
 * SheetName ::= Identifier | '\'' AnyCharsExceptQuote '\''
 * AbsoluteIndicator ::= '$'
 *
 * NamedExpression ::= Identifier  // Not followed by '(' or ':'
 *
 * FunctionCall ::= FunctionName '(' ArgumentList? ')'
 * ArgumentList ::= Expression (',' Expression)*
 *
 * ArrayLiteral ::= '{' ArrayRows '}'
 * ArrayRows ::= ArrayRow (';' ArrayRow)*
 * ArrayRow ::= Expression (',' Expression)*
 *
 * Number ::= [+-]? [0-9]+ ('.' [0-9]+)? ([eE] [+-]? [0-9]+)?
 * String ::= '"' (AnyCharExceptQuote | '""')* '"'
 * Boolean ::= 'TRUE' | 'FALSE'
 * Error ::= '#DIV/0!' | '#N/A' | '#NAME?' | '#NUM!' | '#REF!' | '#VALUE!' | '#CYCLE!' | '#ERROR!'
 * Identifier ::= [A-Za-z_][A-Za-z0-9_]*
 * FunctionName ::= Identifier
 */

/**
 * Cell reference patterns
 */
export const CELL_REFERENCE_PATTERNS = {
  // Column patterns
  COLUMN: /^[A-Z]+$/i,
  COLUMN_WITH_ABSOLUTE: /^\$?[A-Z]+$/i,

  // Row patterns
  ROW: /^[1-9][0-9]*$/,
  ROW_WITH_ABSOLUTE: /^\$?[1-9][0-9]*$/,

  // Full cell reference (e.g., A1, $A$1)
  CELL: /^(\$)?([A-Z]+)(\$)?([1-9][0-9]*)$/i,

  // Sheet qualified reference (e.g., Sheet1!A1, 'My Sheet'!$A$1)
  SHEET_QUALIFIED: /^(?:([A-Za-z_][A-Za-z0-9_]*)|'([^']+)')!(.+)$/,

  // 3D range sheet reference (e.g., Sheet1:Sheet5!A1)
  SHEET_RANGE_QUALIFIED:
    /^(?:(?:([A-Za-z_][A-Za-z0-9_]*)|'([^']+)'):(?:([A-Za-z_][A-Za-z0-9_]*)|'([^']+)'))!(.+)$/,

  // Infinite column range (e.g., A:A, $B:$B)
  INFINITE_COLUMN: /^(\$)?([A-Z]+):(\$)?([A-Z]+)$/i,

  // Infinite row range (e.g., 5:5, $10:$10)
  INFINITE_ROW: /^(\$)?([1-9][0-9]*):(\$)?([1-9][0-9]*)$/i,

  // Open-ended range patterns
  // A5:INFINITY (both row and column unbounded)
  OPEN_ENDED_INFINITY: /^(\$)?([A-Z]+)(\$)?([1-9][0-9]*):INFINITY$/i,
  
  // A5:D (open down only - bounded columns, unbounded rows)
  OPEN_ENDED_COLUMN: /^(\$)?([A-Z]+)(\$)?([1-9][0-9]*):(\$)?([A-Z]+)$/i,
  
  // A5:15 (open right only - bounded rows, unbounded columns)
  OPEN_ENDED_ROW: /^(\$)?([A-Z]+)(\$)?([1-9][0-9]*):(\$)?([1-9][0-9]*)$/i,

  // Structured reference patterns
  TABLE_REFERENCE: /^([A-Za-z_][A-Za-z0-9_]*)\[(.+)\]$/,
  CURRENT_ROW_REFERENCE: /^@([A-Za-z_][A-Za-z0-9_]*)$/,
  TABLE_SELECTOR: /^#(All|Data|Headers)$/i,
};

/**
 * Validate a column reference
 */
export function isValidColumn(col: string): boolean {
  return CELL_REFERENCE_PATTERNS.COLUMN.test(col);
}

/**
 * Validate a row reference
 */
export function isValidRow(row: string): boolean {
  return CELL_REFERENCE_PATTERNS.ROW.test(row);
}

/**
 * Parse a cell reference into components
 */
interface ParsedCellReference {
  sheet?: string;
  colAbsolute: boolean;
  col: string;
  rowAbsolute: boolean;
  row: string;
}

interface ParsedInfiniteRange {
  sheet?: string;
  type: "column" | "row";
  startAbsolute: boolean;
  start: string;
  endAbsolute: boolean;
  end: string;
}

export function parseCellReference(ref: string): ParsedCellReference | null {
  // Check for sheet qualifier
  const sheetMatch = ref.match(CELL_REFERENCE_PATTERNS.SHEET_QUALIFIED);
  let sheet: string | undefined;
  let cellPart: string;

  if (sheetMatch && sheetMatch[3]) {
    sheet = sheetMatch[1] || sheetMatch[2]; // Either unquoted or quoted sheet name
    cellPart = sheetMatch[3];
  } else {
    cellPart = ref;
  }

  // Parse cell part
  const cellMatch = cellPart.match(CELL_REFERENCE_PATTERNS.CELL);
  if (!cellMatch || !cellMatch[2] || !cellMatch[4]) {
    return null;
  }

  return {
    sheet,
    colAbsolute: cellMatch[1] === "$",
    col: cellMatch[2].toUpperCase(),
    rowAbsolute: cellMatch[3] === "$",
    row: cellMatch[4],
  };
}

export function parseInfiniteRange(ref: string): ParsedInfiniteRange | null {
  // Check for sheet qualifier
  const sheetMatch = ref.match(CELL_REFERENCE_PATTERNS.SHEET_QUALIFIED);
  let sheet: string | undefined;
  let rangePart: string;

  if (sheetMatch && sheetMatch[3]) {
    sheet = sheetMatch[1] || sheetMatch[2]; // Either unquoted or quoted sheet name
    rangePart = sheetMatch[3];
  } else {
    rangePart = ref;
  }

  // Check for infinite column range (e.g., A:A, B:C)
  const colMatch = rangePart.match(CELL_REFERENCE_PATTERNS.INFINITE_COLUMN);
  if (colMatch && colMatch[2] && colMatch[4]) {
    return {
      sheet,
      type: "column",
      startAbsolute: colMatch[1] === "$",
      start: colMatch[2].toUpperCase(),
      endAbsolute: colMatch[3] === "$",
      end: colMatch[4].toUpperCase(),
    };
  }

  // Check for infinite row range (e.g., 5:5, 1:10)
  const rowMatch = rangePart.match(CELL_REFERENCE_PATTERNS.INFINITE_ROW);
  if (rowMatch && rowMatch[2] && rowMatch[4]) {
    return {
      sheet,
      type: "row",
      startAbsolute: rowMatch[1] === "$",
      start: rowMatch[2],
      endAbsolute: rowMatch[3] === "$",
      end: rowMatch[4],
    };
  }

  return null;
}

/**
 * Parse open-ended range patterns (A5:INFINITY, A5:D, A5:15)
 */
interface ParsedOpenEndedRange {
  sheet?: string;
  type: "infinity" | "column-bounded" | "row-bounded";
  startCol: string;
  startRow: string;
  startColAbsolute: boolean;
  startRowAbsolute: boolean;
  endCol?: string;
  endRow?: string;
  endColAbsolute?: boolean;
  endRowAbsolute?: boolean;
}

export function parseOpenEndedRange(ref: string): ParsedOpenEndedRange | null {
  // Check for sheet qualifier
  const sheetMatch = ref.match(CELL_REFERENCE_PATTERNS.SHEET_QUALIFIED);
  let sheet: string | undefined;
  let rangePart: string;

  if (sheetMatch && sheetMatch[3]) {
    sheet = sheetMatch[1] || sheetMatch[2]; // Either unquoted or quoted sheet name
    rangePart = sheetMatch[3];
  } else {
    rangePart = ref;
  }

  // Check for A5:INFINITY pattern
  const infinityMatch = rangePart.match(CELL_REFERENCE_PATTERNS.OPEN_ENDED_INFINITY);
  if (infinityMatch && infinityMatch[2] && infinityMatch[4]) {
    return {
      sheet,
      type: "infinity",
      startCol: infinityMatch[2].toUpperCase(),
      startRow: infinityMatch[4],
      startColAbsolute: infinityMatch[1] === "$",
      startRowAbsolute: infinityMatch[3] === "$",
    };
  }

  // Check for A5:D pattern (open down only)
  // We need to be more careful here to distinguish from normal cell ranges
  const colonIndex = rangePart.indexOf(':');
  if (colonIndex !== -1) {
    const startPart = rangePart.substring(0, colonIndex);
    const endPart = rangePart.substring(colonIndex + 1);
    
    // Parse start part as a cell reference
    const startMatch = startPart.match(/^(\$)?([A-Z]+)(\$)?([1-9][0-9]*)$/i);
    if (startMatch && startMatch[2] && startMatch[4]) {
      // Check if end part is just a column (A5:D pattern)
      const endColMatch = endPart.match(/^(\$)?([A-Z]+)$/i);
      if (endColMatch && endColMatch[2]) {
        return {
          sheet,
          type: "column-bounded",
          startCol: startMatch[2].toUpperCase(),
          startRow: startMatch[4],
          startColAbsolute: startMatch[1] === "$",
          startRowAbsolute: startMatch[3] === "$",
          endCol: endColMatch[2].toUpperCase(),
          endColAbsolute: endColMatch[1] === "$",
        };
      }
      
      // Check if end part is just a row number (A5:15 pattern)
      const endRowMatch = endPart.match(/^(\$)?([1-9][0-9]*)$/);
      if (endRowMatch && endRowMatch[2]) {
        return {
          sheet,
          type: "row-bounded",
          startCol: startMatch[2].toUpperCase(),
          startRow: startMatch[4],
          startColAbsolute: startMatch[1] === "$",
          startRowAbsolute: startMatch[3] === "$",
          endRow: endRowMatch[2],
          endRowAbsolute: endRowMatch[1] === "$",
        };
      }
    }
  }

  return null;
}

/**
 * Parse a 3D range reference (e.g., Sheet1:Sheet3!A1)
 */
interface Parsed3DReference {
  startSheet: string;
  endSheet: string;
  reference: string;
}

export function parse3DReference(ref: string): Parsed3DReference | null {
  const match = ref.match(CELL_REFERENCE_PATTERNS.SHEET_RANGE_QUALIFIED);
  if (!match || !match[5]) {
    return null;
  }

  const startSheet = match[1] || match[2]; // Unquoted or quoted start sheet
  const endSheet = match[3] || match[4]; // Unquoted or quoted end sheet

  if (!startSheet || !endSheet) {
    return null;
  }

  return {
    startSheet,
    endSheet,
    reference: match[5],
  };
}

/**
 * Parse a structured reference (e.g., Table1[Column1])
 */
interface ParsedStructuredReference {
  tableName: string;
  columnSpec: string;
  isCurrentRow?: boolean;
  selector?: string;
  cols?: {
    startCol: string;
    endCol: string;
  };
}

export function parseStructuredReference(
  ref: string
): ParsedStructuredReference | null {
  // Check for current row reference (e.g., [@Column])
  if (ref.startsWith("@")) {
    const colName = ref.substring(1);
    if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(colName)) {
      return {
        tableName: "",
        columnSpec: colName,
        isCurrentRow: true,
        cols: {
          startCol: colName,
          endCol: colName,
        },
      };
    }
  }

  // Check for table reference (e.g., Table1[Column1])
  const match = ref.match(CELL_REFERENCE_PATTERNS.TABLE_REFERENCE);
  if (!match || !match[1] || !match[2]) {
    return null;
  }

  const tableName = match[1];
  const columnSpec = match[2];

  // Parse column spec for selectors and column name/range
  let selector: string | undefined;
  let cols: { startCol: string; endCol: string } | undefined;

  // Check if column spec contains selector (e.g., [#Headers],[Column1])
  const selectorMatch = columnSpec.match(/^\[#(All|Data|Headers)\],\[(.+)\]$/i);
  if (selectorMatch && selectorMatch[1] && selectorMatch[2]) {
    selector = "#" + selectorMatch[1];
    const colSpec = selectorMatch[2];

    // Check if it's a column range [Column1]:[Column2]
    const rangeMatch = colSpec.match(/^(.+):(.+)$/);
    if (rangeMatch && rangeMatch[1] && rangeMatch[2]) {
      cols = {
        startCol: rangeMatch[1].trim(),
        endCol: rangeMatch[2].trim(),
      };
    } else {
      cols = {
        startCol: colSpec,
        endCol: colSpec,
      };
    }
  } else if (columnSpec.startsWith("@")) {
    // Table with current row reference (e.g., Table1[@Column1])
    const colName = columnSpec.substring(1);
    return {
      tableName,
      columnSpec: colName,
      isCurrentRow: true,
      cols: {
        startCol: colName,
        endCol: colName,
      },
    };
  } else if (columnSpec.startsWith("#")) {
    // Just a selector (e.g., Table1[#Data])
    const selectorMatch = columnSpec.match(/^#(All|Data|Headers)$/i);
    if (selectorMatch) {
      selector = columnSpec;
    }
  } else {
    // Check if it's a column range [Column1]:[Column2]
    const rangeMatch = columnSpec.match(/^(.+):(.+)$/);
    if (rangeMatch && rangeMatch[1] && rangeMatch[2]) {
      cols = {
        startCol: rangeMatch[1].trim(),
        endCol: rangeMatch[2].trim(),
      };
    } else {
      // Simple column reference
      cols = {
        startCol: columnSpec,
        endCol: columnSpec,
      };
    }
  }

  return {
    tableName,
    columnSpec,
    selector,
    cols,
  };
}
