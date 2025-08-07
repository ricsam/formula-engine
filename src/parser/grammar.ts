/**
 * Grammar rules and operator precedence for formula parsing
 */

import type { BinaryOpNode } from './ast';

/**
 * Operator precedence levels (higher number = higher precedence)
 */
export const OPERATOR_PRECEDENCE: Record<string, number> = {
  // Comparison operators (lowest precedence)
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  
  // Concatenation
  '&': 2,
  
  // Addition and subtraction
  '+': 3,
  '-': 3,
  
  // Multiplication and division
  '*': 4,
  '/': 4,
  
  // Exponentiation (highest precedence)
  '^': 5,
};

/**
 * Operator associativity
 */
export type Associativity = 'left' | 'right';

export const OPERATOR_ASSOCIATIVITY: Record<string, Associativity> = {
  '=': 'left',
  '<>': 'left',
  '<': 'left',
  '>': 'left',
  '<=': 'left',
  '>=': 'left',
  '&': 'left',
  '+': 'left',
  '-': 'left',
  '*': 'left',
  '/': 'left',
  '^': 'right',  // Exponentiation is right-associative
};

/**
 * Check if a string is a valid binary operator
 */
export function isBinaryOperator(op: string): op is BinaryOpNode['operator'] {
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
  return OPERATOR_ASSOCIATIVITY[op] || 'left';
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
  'PI',   // PI() can be written as PI
  'TRUE', // TRUE() can be written as TRUE
  'FALSE', // FALSE() can be written as FALSE
  'NA',   // NA() can be written as NA
]);

/**
 * Functions that accept variable number of arguments
 */
export const VARIADIC_FUNCTIONS = new Set([
  'SUM',
  'PRODUCT',
  'COUNT',
  'MAX',
  'MIN',
  'AND',
  'OR',
  'XOR',
  'CONCATENATE',
  'CHOOSE',
  'IFS',
  'SWITCH',
]);

/**
 * Functions that require at least one argument
 */
export const REQUIRED_ARG_FUNCTIONS = new Set([
  'SUM',
  'PRODUCT',
  'COUNT',
  'MAX',
  'MIN',
  'AVERAGE',
  'MEDIAN',
  'STDEV',
  'VAR',
  'AND',
  'OR',
  'XOR',
  'CONCATENATE',
]);

/**
 * Reserved keywords that cannot be used as names
 */
export const RESERVED_KEYWORDS = new Set([
  'TRUE',
  'FALSE',
  'NULL',
  'DIV',
  'MOD',
  'AND',
  'OR',
  'NOT',
  'XOR',
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
  
  // Infinite column range (e.g., A:A, $B:$B)
  INFINITE_COLUMN: /^(\$)?([A-Z]+):(\$)?([A-Z]+)$/i,
  
  // Infinite row range (e.g., 5:5, $10:$10)
  INFINITE_ROW: /^(\$)?([1-9][0-9]*):(\$)?([1-9][0-9]*)$/i,
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
export interface ParsedCellReference {
  sheet?: string;
  colAbsolute: boolean;
  col: string;
  rowAbsolute: boolean;
  row: string;
}

export interface ParsedInfiniteRange {
  sheet?: string;
  type: 'column' | 'row';
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
    colAbsolute: cellMatch[1] === '$',
    col: cellMatch[2].toUpperCase(),
    rowAbsolute: cellMatch[3] === '$',
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
      type: 'column',
      startAbsolute: colMatch[1] === '$',
      start: colMatch[2].toUpperCase(),
      endAbsolute: colMatch[3] === '$',
      end: colMatch[4].toUpperCase(),
    };
  }
  
  // Check for infinite row range (e.g., 5:5, 1:10)
  const rowMatch = rangePart.match(CELL_REFERENCE_PATTERNS.INFINITE_ROW);
  if (rowMatch && rowMatch[2] && rowMatch[4]) {
    return {
      sheet,
      type: 'row',
      startAbsolute: rowMatch[1] === '$',
      start: rowMatch[2],
      endAbsolute: rowMatch[3] === '$',
      end: rowMatch[4],
    };
  }
  
  return null;
}

/**
 * R1C1 reference style patterns (for future support)
 */
export const R1C1_PATTERNS = {
  // R1C1, R[1]C[1], R[-1]C[-1]
  CELL: /^R(\[?-?\d+\]?)C(\[?-?\d+\]?)$/i,
  
  // R1C1:R2C2
  RANGE: /^R(\[?-?\d+\]?)C(\[?-?\d+\]?):R(\[?-?\d+\]?)C(\[?-?\d+\]?)$/i,
};

/**
 * Function argument count constraints
 */
export interface FunctionConstraints {
  minArgs?: number;
  maxArgs?: number;
  exactArgs?: number;
}

export const FUNCTION_CONSTRAINTS: Record<string, FunctionConstraints> = {
  // Math functions
  'ABS': { exactArgs: 1 },
  'SIGN': { exactArgs: 1 },
  'SQRT': { exactArgs: 1 },
  'POWER': { exactArgs: 2 },
  'EXP': { exactArgs: 1 },
  'LN': { exactArgs: 1 },
  'LOG': { minArgs: 1, maxArgs: 2 },
  'LOG10': { exactArgs: 1 },
  'SIN': { exactArgs: 1 },
  'COS': { exactArgs: 1 },
  'TAN': { exactArgs: 1 },
  'ASIN': { exactArgs: 1 },
  'ACOS': { exactArgs: 1 },
  'ATAN': { exactArgs: 1 },
  'ATAN2': { exactArgs: 2 },
  'DEGREES': { exactArgs: 1 },
  'RADIANS': { exactArgs: 1 },
  'PI': { exactArgs: 0 },
  'ROUND': { exactArgs: 2 },
  'ROUNDUP': { exactArgs: 2 },
  'ROUNDDOWN': { exactArgs: 2 },
  'CEILING': { minArgs: 1, maxArgs: 2 },
  'FLOOR': { minArgs: 1, maxArgs: 2 },
  'INT': { exactArgs: 1 },
  'TRUNC': { minArgs: 1, maxArgs: 2 },
  'MOD': { exactArgs: 2 },
  'EVEN': { exactArgs: 1 },
  'ODD': { exactArgs: 1 },
  'FACT': { exactArgs: 1 },
  'DECIMAL': { exactArgs: 2 },
  
  // Statistical functions
  'SUM': { minArgs: 1 },
  'PRODUCT': { minArgs: 1 },
  'COUNT': { minArgs: 1 },
  'COUNTBLANK': { exactArgs: 1 },
  'COUNTIF': { exactArgs: 2 },
  'SUMIF': { minArgs: 2, maxArgs: 3 },
  'SUMIFS': { minArgs: 3 },  // sum_range, criteria_range1, criteria1, ...
  'AVERAGE': { minArgs: 1 },
  'MAX': { minArgs: 1 },
  'MIN': { minArgs: 1 },
  'MEDIAN': { minArgs: 1 },
  'STDEV': { minArgs: 1 },
  'VAR': { minArgs: 1 },
  'COVAR': { exactArgs: 2 },
  'GAMMA': { exactArgs: 1 },
  
  // Logical functions
  'IF': { exactArgs: 3 },
  'IFS': { minArgs: 2 },  // condition1, value1, condition2, value2, ...
  'SWITCH': { minArgs: 3 },  // expression, value1, result1, ...
  'AND': { minArgs: 1 },
  'OR': { minArgs: 1 },
  'NOT': { exactArgs: 1 },
  'XOR': { minArgs: 1 },
  'TRUE': { exactArgs: 0 },
  'FALSE': { exactArgs: 0 },
  'IFERROR': { exactArgs: 2 },
  'IFNA': { exactArgs: 2 },
  
  // Text functions
  'CONCATENATE': { minArgs: 1 },
  'LEN': { exactArgs: 1 },
  'TRIM': { exactArgs: 1 },
  'UPPER': { exactArgs: 1 },
  'LOWER': { exactArgs: 1 },
  'EXACT': { exactArgs: 2 },
  'TEXT': { exactArgs: 2 },
  
  // Lookup functions
  'VLOOKUP': { minArgs: 3, maxArgs: 4 },
  'HLOOKUP': { minArgs: 3, maxArgs: 4 },
  'INDEX': { minArgs: 2, maxArgs: 3 },
  'MATCH': { minArgs: 2, maxArgs: 3 },
  'XLOOKUP': { minArgs: 3, maxArgs: 6 },
  'CHOOSE': { minArgs: 2 },
  'OFFSET': { minArgs: 3, maxArgs: 5 },
  'COLUMN': { minArgs: 0, maxArgs: 1 },
  'COLUMNS': { exactArgs: 1 },
  'ROW': { minArgs: 0, maxArgs: 1 },
  'ROWS': { exactArgs: 1 },
  'ADDRESS': { minArgs: 2, maxArgs: 5 },
  'FORMULATEXT': { exactArgs: 1 },
  
  // Info functions
  'ISBLANK': { exactArgs: 1 },
  'ISERROR': { exactArgs: 1 },
  'ISERR': { exactArgs: 1 },
  'ISNA': { exactArgs: 1 },
  'ISNUMBER': { exactArgs: 1 },
  'ISTEXT': { exactArgs: 1 },
  'ISLOGICAL': { exactArgs: 1 },
  'ISNONTEXT': { exactArgs: 1 },
  'ISFORMULA': { exactArgs: 1 },
  'ISEVEN': { exactArgs: 1 },
  'ISODD': { exactArgs: 1 },
  'ISBINARY': { exactArgs: 1 },
  'ISREF': { exactArgs: 1 },
  'SHEET': { minArgs: 0, maxArgs: 1 },
  'SHEETS': { minArgs: 0, maxArgs: 1 },
  'NA': { exactArgs: 0 },
  
  // Array functions
  'FILTER': { minArgs: 2 },
  'ARRAY_CONSTRAIN': { exactArgs: 3 },
  
  // FE internal operators as functions
  'FE.ADD': { exactArgs: 2 },
  'FE.MINUS': { exactArgs: 2 },
  'FE.MULTIPLY': { exactArgs: 2 },
  'FE.DIVIDE': { exactArgs: 2 },
  'FE.POW': { exactArgs: 2 },
  'FE.UMINUS': { exactArgs: 1 },
  'FE.UPLUS': { exactArgs: 1 },
  'FE.UNARY_PERCENT': { exactArgs: 1 },
  'FE.EQ': { exactArgs: 2 },
  'FE.NE': { exactArgs: 2 },
  'FE.LT': { exactArgs: 2 },
  'FE.LTE': { exactArgs: 2 },
  'FE.GT': { exactArgs: 2 },
  'FE.GTE': { exactArgs: 2 },
  'FE.CONCAT': { exactArgs: 2 },
};

/**
 * Get function argument constraints
 */
export function getFunctionConstraints(functionName: string): FunctionConstraints | undefined {
  return FUNCTION_CONSTRAINTS[functionName.toUpperCase()];
}

/**
 * Validate function argument count
 */
export function validateFunctionArgCount(functionName: string, argCount: number): boolean {
  const constraints = getFunctionConstraints(functionName);
  if (!constraints) {
    // Unknown function, allow any number of arguments
    return true;
  }
  
  if (constraints.exactArgs !== undefined) {
    return argCount === constraints.exactArgs;
  }
  
  if (constraints.minArgs !== undefined && argCount < constraints.minArgs) {
    return false;
  }
  
  if (constraints.maxArgs !== undefined && argCount > constraints.maxArgs) {
    return false;
  }
  
  return true;
}
