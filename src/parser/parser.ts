/**
 * Recursive descent parser for formula expressions
 */

import type { 
  FormulaAST, 
  ASTNode,
  ValueNode,
  ReferenceNode,
  RangeNode,
  FunctionNode,
  UnaryOpNode,
  BinaryOpNode,
  ArrayNode,
  NamedExpressionNode,
  ErrorNode
} from './ast';
import type { SimpleCellAddress, SimpleCellRange, FormulaError } from '../core/types';
import { 
  createValueNode, 
  createReferenceNode, 
  createRangeNode,
  createFunctionNode,
  createUnaryOpNode,
  createBinaryOpNode,
  createArrayNode,
  createNamedExpressionNode,
  createErrorNode
} from './ast';
import { Lexer, TokenStream, type Token, type TokenType } from './lexer';
import { 
  isBinaryOperator,
  getOperatorPrecedence,
  getOperatorAssociativity,
  parseCellReference,
  parseInfiniteRange,
  validateFunctionArgCount,
  SPECIAL_FUNCTIONS
} from './grammar';
import { letterToColNumber } from '../core/types';

/**
 * Parser error class
 */
export class ParseError extends Error {
  constructor(
    message: string,
    public position?: { start: number; end: number }
  ) {
    super(message);
    this.name = 'ParseError';
  }
}

/**
 * Sheet resolver function type
 */
export type SheetResolver = (sheetName: string) => number;

/**
 * Formula parser class
 */
export class Parser {
  private tokens: TokenStream;
  private contextSheetId: number;
  private sheetResolver?: SheetResolver;
  
  constructor(tokens: Token[], contextSheetId: number = 0, sheetResolver?: SheetResolver) {
    this.tokens = new TokenStream(tokens);
    this.contextSheetId = contextSheetId;
    this.sheetResolver = sheetResolver;
  }
  
  /**
   * Parse a formula string
   */
  static parse(formula: string, contextSheetId: number = 0, sheetResolver?: SheetResolver): FormulaAST {
    // Remove leading '=' if present
    const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;
    
    // Tokenize
    const lexer = new Lexer(cleanFormula);
    const tokens = lexer.tokenize();
    
    // Parse
    const parser = new Parser(tokens, contextSheetId, sheetResolver);
    return parser.parseFormula();
  }
  
  /**
   * Parse the entire formula
   */
  parseFormula(): FormulaAST {
    if (this.tokens.isAtEnd()) {
      // Empty formula
      return createValueNode(undefined);
    }
    
    // Handle edge case of just '=' with nothing after
    if (this.tokens.peek().type === 'EOF') {
      return createValueNode(undefined);
    }
    
    try {
      const expr = this.parseExpression();
      
      // Ensure we've consumed all tokens
      if (!this.tokens.isAtEnd()) {
        throw new ParseError(
          `Unexpected token: ${this.tokens.peek().value}`,
          this.tokens.peek().position
        );
      }
      
      return expr as FormulaAST;
    } catch (error) {
      if (error instanceof ParseError) {
        throw error;
      }
      throw new ParseError(String(error));
    }
  }
  
  /**
   * Parse an expression (entry point for recursive descent)
   */
  private parseExpression(): ASTNode {
    return this.parseBinaryExpression(0);
  }
  
  /**
   * Parse binary expressions using precedence climbing
   */
  private parseBinaryExpression(minPrecedence: number): ASTNode {
    let left = this.parseUnaryExpression();
    
    while (true) {
      const token = this.tokens.peek();
      
      if (token.type !== 'OPERATOR' || !isBinaryOperator(token.value)) {
        break;
      }
      
      const precedence = getOperatorPrecedence(token.value);
      if (precedence < minPrecedence) {
        break;
      }
      
      const operator = token.value as any;
      const associativity = getOperatorAssociativity(operator);
      const start = token.position.start;
      
      this.tokens.consume(); // Consume operator
      
      // For right-associative operators, use same precedence
      // For left-associative, use precedence + 1
      const nextMinPrecedence = associativity === 'right' ? precedence : precedence + 1;
      const right = this.parseBinaryExpression(nextMinPrecedence);
      
      left = createBinaryOpNode(operator, left, right, start, right.position?.end);
    }
    
    return left;
  }
  
  /**
   * Parse unary expressions
   */
  private parseUnaryExpression(): ASTNode {
    const token = this.tokens.peek();
    
    // Check for unary operators
    if (token.type === 'OPERATOR' && (token.value === '+' || token.value === '-')) {
      const start = token.position.start;
      const operator = token.value as '+' | '-';
      
      this.tokens.consume();
      const operand = this.parseUnaryExpression();
      
      return createUnaryOpNode(operator, operand, start, operand.position?.end);
    }
    
    return this.parsePostfixExpression();
  }
  
  /**
   * Parse postfix expressions (currently just %)
   */
  private parsePostfixExpression(): ASTNode {
    let expr = this.parsePrimaryExpression();
    
    // Check for percentage operator
    if (this.tokens.match('OPERATOR') && this.tokens.peek().value === '%') {
      const token = this.tokens.consume();
      expr = createUnaryOpNode(
        '%',
        expr,
        expr.position?.start,
        token.position.end
      );
    }
    
    return expr;
  }
  
  /**
   * Parse primary expressions
   */
  private parsePrimaryExpression(): ASTNode {
    const token = this.tokens.peek();
    const start = token.position.start;
    
    switch (token.type) {
      case 'NUMBER':
        return this.parseNumber();
        
      case 'STRING':
        return this.parseString();
        
      case 'BOOLEAN':
        return this.parseBoolean();
        
      case 'ERROR':
        return this.parseError();
        
      case 'FUNCTION':
        return this.parseFunctionCall();
        
      case 'IDENTIFIER':
        return this.parseIdentifier();
        
      case 'DOLLAR':
        // Handle absolute reference starting with $
        return this.parseAbsoluteReference();
        
      case 'LPAREN':
        return this.parseParenthesizedExpression();
        
      case 'LBRACE':
        return this.parseArrayLiteral();
        
      default:
        throw new ParseError(
          `Unexpected token: ${token.value}`,
          token.position
        );
    }
  }
  
  /**
   * Parse an absolute reference starting with $
   */
  private parseAbsoluteReference(): ASTNode {
    const start = this.tokens.peek().position.start;
    let ref = '';
    
    // First $
    ref += this.tokens.consume().value;
    
    // Get next token
    if (this.tokens.match('IDENTIFIER')) {
      const identifier = this.tokens.consume().value;
      
      // Check if this identifier contains both column and row (e.g., "C3")
      const match = identifier.match(/^([A-Z]+)(\d+)$/i);
      if (match) {
        // It's a complete cell reference like $C3
        ref += identifier;
      } else if (/^[A-Z]+$/i.test(identifier)) {
        // It's just the column part
        ref += identifier;
        
        // Check for $ before row
        if (this.tokens.match('DOLLAR')) {
          ref += this.tokens.consume().value;
        }
        
        // Row number
        if (this.tokens.match('NUMBER')) {
          ref += this.tokens.consume().value;
        } else {
          throw new ParseError('Expected row number', this.tokens.peek().position);
        }
      } else {
        throw new ParseError('Invalid cell reference format', this.tokens.peek().position);
      }
    } else {
      throw new ParseError('Expected column letter after $', this.tokens.peek().position);
    }
    
    // Check for range
    if (this.tokens.match('COLON')) {
      this.tokens.consume();
      const endRef = this.parseRangeEnd();
      return this.parseRange(ref, endRef, start, this.tokens.peek().position.start);
    }
    
    // Parse as single cell reference
    const cellRef = this.parseCellReferenceString(ref);
    if (cellRef) {
      return cellRef;
    }
    
    throw new ParseError(`Invalid cell reference: ${ref}`, { start, end: this.tokens.peek().position.start });
  }
  
  /**
   * Parse a number literal or row range
   */
  private parseNumber(): ASTNode {
    const token = this.tokens.peek();
    const start = token.position.start;
    
    // Check if this could be a row range (e.g., 5:5, 1:10)
    if (this.tokens.peekNext() && this.tokens.peekNext()!.type === 'COLON') {
      // This is a row range
      const startRow = this.tokens.consume().value;
      this.tokens.consume(); // Consume ':'
      
      // Get the end row
      let endRow: string;
      if (this.tokens.match('NUMBER')) {
        endRow = this.tokens.consume().value;
      } else {
        throw new ParseError('Expected row number after :', this.tokens.peek().position);
      }
      
      // Parse as an infinite row range
      return this.parseRange(startRow, endRow, start, this.tokens.peek().position.start);
    }
    
    // Otherwise, parse as a regular number
    this.tokens.consume();
    const value = parseFloat(token.value);
    
    if (isNaN(value)) {
      throw new ParseError(
        `Invalid number: ${token.value}`,
        token.position
      );
    }
    
    return createValueNode(value, token.position.start, token.position.end);
  }
  
  /**
   * Parse a string literal
   */
  private parseString(): ASTNode {
    const token = this.tokens.consume();
    return createValueNode(token.value, token.position.start, token.position.end);
  }
  
  /**
   * Parse a boolean literal
   */
  private parseBoolean(): ASTNode {
    const token = this.tokens.consume();
    const value = token.value.toUpperCase() === 'TRUE';
    return createValueNode(value, token.position.start, token.position.end);
  }
  
  /**
   * Parse an error literal
   */
  private parseError(): ASTNode {
    const token = this.tokens.consume();
    const error = token.value as FormulaError;
    return createErrorNode(
      error,
      `Error literal: ${error}`,
      token.position.start,
      token.position.end
    );
  }
  
  /**
   * Parse a function call
   */
  private parseFunctionCall(): ASTNode {
    const nameToken = this.tokens.consume();
    const functionName = nameToken.value;
    const start = nameToken.position.start;
    
    // Check for special functions that don't require parentheses
    if (SPECIAL_FUNCTIONS.has(functionName.toUpperCase()) && !this.tokens.match('LPAREN')) {
      return createFunctionNode(
        functionName,
        [],
        start,
        nameToken.position.end
      );
    }
    
    // Expect opening parenthesis
    if (!this.tokens.match('LPAREN')) {
      throw new ParseError(
        `Expected '(' after function name ${functionName}`,
        this.tokens.peek().position
      );
    }
    
    this.tokens.consume(); // Consume '('
    
    // Parse arguments
    const args: ASTNode[] = [];
    
    // Handle empty argument list
    if (this.tokens.match('RPAREN')) {
      const rparenToken = this.tokens.consume();
      const endPos = rparenToken.position.end;
      const node = createFunctionNode(
        functionName,
        args,
        start,
        endPos
      );
      
      // Validate argument count
      if (!validateFunctionArgCount(functionName, args.length)) {
        throw new ParseError(
          `Invalid number of arguments for function ${functionName}`,
          { start, end: endPos }
        );
      }
      
      return node;
    }
    
    // Parse arguments
    while (true) {
      args.push(this.parseExpression());
      
      if (this.tokens.match('COMMA')) {
        this.tokens.consume();
        // Continue parsing next argument
      } else if (this.tokens.match('RPAREN')) {
        const rparenToken = this.tokens.consume();
        const end = rparenToken.position.end;
        const node = createFunctionNode(functionName, args, start, end);
        
        // Validate argument count
        if (!validateFunctionArgCount(functionName, args.length)) {
          throw new ParseError(
            `Invalid number of arguments for function ${functionName}`,
            { start, end }
          );
        }
        
        return node;
      } else {
        throw new ParseError(
          `Expected ',' or ')' in function arguments`,
          this.tokens.peek().position
        );
      }
    }
    
    // This should never be reached
    throw new ParseError(
      'Unexpected end of function argument parsing',
      this.tokens.peek().position
    );
  }
  
  /**
   * Parse an identifier (cell reference, range, or named expression)
   */
  private parseIdentifier(): ASTNode {
    const start = this.tokens.peek().position.start;
    
    // Regular identifier
    const token = this.tokens.consume();
    let value = token.value;
    
    // Check if this is a sheet reference (Sheet1! or 'My Sheet'!)
    if (this.tokens.match('EXCLAMATION')) {
      this.tokens.consume(); // Consume '!'
      
      // Extract sheet name (remove quotes if present)
      let sheetName = value;
      if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
        sheetName = sheetName.slice(1, -1).replace(/''/g, "'"); // Remove quotes and unescape
      }
      
      // Now parse the cell reference after the sheet name
      const cellRef = this.parseCellOrRangeWithSheet(sheetName, start);
      return cellRef;
    }
    
    // Check if this could be part of a cell reference (e.g., D in D$4)
    if (this.tokens.match('DOLLAR') && this.isColumnIdentifier(value)) {
      // This is a mixed reference like D$4
      value += this.tokens.consume().value; // Add $
      
      // Get the row number
      if (this.tokens.match('NUMBER')) {
        value += this.tokens.consume().value;
      } else {
        throw new ParseError('Expected row number after $', this.tokens.peek().position);
      }
    } else if (this.tokens.match('NUMBER') && this.isColumnIdentifier(value)) {
      // Regular cell reference like D4
      value += this.tokens.consume().value;
    }
    
    // Check for colon (range operator)
    if (this.tokens.match('COLON')) {
      this.tokens.consume(); // Consume ':'
      
      // Parse end of range with potential $ signs
      const endStart = this.tokens.peek().position.start;
      const endRef = this.parseRangeEnd();
      const endPos = this.tokens.peek().position.start;
      
      // Parse as range
      return this.parseRange(value, endRef, start, endPos);
    }
    
    // Try to parse as cell reference
    const cellRef = this.parseCellReferenceString(value);
    if (cellRef) {
      return cellRef;
    }
    
    // Otherwise, it's a named expression
    return createNamedExpressionNode(
      value,
      undefined,
      start,
      token.position.end
    );
  }
  
  /**
   * Check if a string is a valid column identifier (A-Z, AA-ZZ, etc.)
   */
  private isColumnIdentifier(str: string): boolean {
    return /^[A-Z]+$/i.test(str);
  }
  
  /**
   * Parse the end part of a range (handling $ signs and infinite ranges)
   */
  private parseRangeEnd(): string {
    let result = '';
    
    // Check for $ before column or row
    if (this.tokens.match('DOLLAR')) {
      result += this.tokens.consume().value;
    }
    
    // Get identifier part (for column) or number part (for row)
    if (this.tokens.match('IDENTIFIER')) {
      result += this.tokens.consume().value;
      
      // For normal ranges, check for $ before row
      if (this.tokens.match('DOLLAR')) {
        result += this.tokens.consume().value;
      }
      
      // Get number part if present (normal range)
      if (this.tokens.match('NUMBER')) {
        result += this.tokens.consume().value;
      }
      // If no number, it's an infinite column range (e.g., A:A)
    } else if (this.tokens.match('NUMBER')) {
      // Infinite row range (e.g., 5:5)
      result += this.tokens.consume().value;
    } else {
      throw new ParseError('Expected cell reference after :', this.tokens.peek().position);
    }
    
    return result;
  }
  
  /**
   * Parse a parenthesized expression
   */
  private parseParenthesizedExpression(): ASTNode {
    const start = this.tokens.peek().position.start;
    this.tokens.consume(); // Consume '('
    
    const expr = this.parseExpression();
    
    if (!this.tokens.match('RPAREN')) {
      throw new ParseError(
        `Expected ')'`,
        this.tokens.peek().position
      );
    }
    
    const end = this.tokens.peek().position.end;
    this.tokens.consume(); // Consume ')'
    
    // Update position to include parentheses
    if (expr.position) {
      expr.position.start = start;
      expr.position.end = end;
    }
    
    return expr;
  }
  
  /**
   * Parse an array literal
   */
  private parseArrayLiteral(): ASTNode {
    const start = this.tokens.peek().position.start;
    this.tokens.consume(); // Consume '{'
    
    const rows: ASTNode[][] = [];
    let currentRow: ASTNode[] = [];
    
    // Handle empty array
    if (this.tokens.match('RBRACE')) {
      this.tokens.consume();
      return createArrayNode(
        [[createValueNode(undefined)]],
        start,
        this.tokens.peek().position.start
      );
    }
    
    // Parse array elements
    while (true) {
      currentRow.push(this.parseExpression());
      
      if (this.tokens.match('COMMA')) {
        this.tokens.consume();
        // Continue current row
      } else if (this.tokens.match('SEMICOLON')) {
        this.tokens.consume();
        // Start new row
        rows.push(currentRow);
        currentRow = [];
      } else if (this.tokens.match('RBRACE')) {
        this.tokens.consume();
        // End of array
        rows.push(currentRow);
        break;
      } else {
        throw new ParseError(
          `Expected ',', ';', or '}' in array literal`,
          this.tokens.peek().position
        );
      }
    }
    
    // Validate that all rows have the same length
    if (rows.length > 0 && rows[0]) {
      const rowLength = rows[0].length;
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row && row.length !== rowLength) {
          throw new ParseError(
            `Inconsistent row lengths in array literal`,
            { start, end: this.tokens.peek().position.start }
          );
        }
      }
    }
    
    return createArrayNode(
      rows,
      start,
      this.tokens.peek().position.start
    );
  }
  
  /**
   * Parse a cell reference string
   */
  private parseCellReferenceString(value: string): ASTNode | null {
    const parsed = parseCellReference(value);
    if (!parsed) {
      return null;
    }
    
    // Convert to SimpleCellAddress
    const colNum = letterToColNumber(parsed.col);
    const rowNum = parseInt(parsed.row) - 1; // Convert to 0-based
    
    if (colNum < 0 || rowNum < 0) {
      return null;
    }
    
    const address: SimpleCellAddress = {
      sheet: parsed.sheet ? this.getSheetId(parsed.sheet) : this.contextSheetId,
      col: colNum,
      row: rowNum
    };
    
    return createReferenceNode(
      address,
      {
        col: parsed.colAbsolute,
        row: parsed.rowAbsolute
      }
    );
  }
  
  /**
   * Parse a range reference (including infinite ranges)
   */
  private parseRange(startRef: string, endRef: string, startPos: number, endPos: number): ASTNode {
    // For cross-sheet ranges, handle the case where only startRef includes the sheet
    let fullRange = `${startRef}:${endRef}`;
    let sheetName: string | undefined;
    
    // Check if start has a sheet prefix
    const sheetMatch = startRef.match(/^(?:([A-Za-z_][A-Za-z0-9_]*)|'([^']+)')!/);
    if (sheetMatch) {
      sheetName = sheetMatch[1] || sheetMatch[2];
      
      // First, try to parse as an infinite range without modifying the range
      // This handles cases like Sheet1!A:A or Sheet1!5:5
      const infiniteTest = parseInfiniteRange(fullRange);
      if (infiniteTest) {
        // It's an infinite range, skip to processing it below
      } else {
        // Not an infinite range, so handle normal cross-sheet ranges
        const endSheetMatch = endRef.match(/^(?:([A-Za-z_][A-Za-z0-9_]*)|'([^']+)')!/);
        if (!endSheetMatch && sheetName) {
          // Normal case: prepend sheet name to endRef for consistent parsing
          const quotedSheetName = sheetName.includes(' ') ? `'${sheetName}'` : sheetName;
          endRef = `${quotedSheetName}!${endRef}`;
          fullRange = `${startRef}:${endRef}`;
        }
      }
    }
    
    // Try to parse as an infinite range
    const infiniteParsed = parseInfiniteRange(fullRange);
    
    if (infiniteParsed) {
      // Handle infinite range
      const sheetId = infiniteParsed.sheet ? this.getSheetId(infiniteParsed.sheet) : this.contextSheetId;
      
      if (infiniteParsed.type === 'column') {
        // Infinite column range (e.g., A:A, B:D)
        const startCol = letterToColNumber(infiniteParsed.start);
        const endCol = letterToColNumber(infiniteParsed.end);
        
        if (startCol < 0 || endCol < 0) {
          throw new ParseError(
            `Invalid column range: ${startRef}:${endRef}`,
            { start: startPos, end: endPos }
          );
        }
        
        // Use special row values to indicate infinite range
        // We'll use -1 to indicate infinite
        const range: SimpleCellRange = {
          start: {
            sheet: sheetId,
            col: Math.min(startCol, endCol),
            row: 0  // Start from row 0
          },
          end: {
            sheet: sheetId,
            col: Math.max(startCol, endCol),
            row: Number.MAX_SAFE_INTEGER  // Use max value to indicate infinite
          }
        };
        
        return createRangeNode(
          range,
          {
            start: {
              col: infiniteParsed.startAbsolute,
              row: false
            },
            end: {
              col: infiniteParsed.endAbsolute,
              row: false
            }
          },
          startPos,
          endPos
        );
      } else {
        // Infinite row range (e.g., 5:5, 1:10)
        const startRow = parseInt(infiniteParsed.start) - 1;
        const endRow = parseInt(infiniteParsed.end) - 1;
        
        if (startRow < 0 || endRow < 0) {
          throw new ParseError(
            `Invalid row range: ${startRef}:${endRef}`,
            { start: startPos, end: endPos }
          );
        }
        
        const range: SimpleCellRange = {
          start: {
            sheet: sheetId,
            col: 0,  // Start from column 0
            row: Math.min(startRow, endRow)
          },
          end: {
            sheet: sheetId,
            col: Number.MAX_SAFE_INTEGER,  // Use max value to indicate infinite
            row: Math.max(startRow, endRow)
          }
        };
        
        return createRangeNode(
          range,
          {
            start: {
              col: false,
              row: infiniteParsed.startAbsolute
            },
            end: {
              col: false,
              row: infiniteParsed.endAbsolute
            }
          },
          startPos,
          endPos
        );
      }
    }
    
    // Otherwise, parse as normal range
    const startParsed = parseCellReference(startRef);
    const endParsed = parseCellReference(endRef);
    
    if (!startParsed || !endParsed) {
      throw new ParseError(
        `Invalid range reference: ${startRef}:${endRef}`,
        { start: startPos, end: endPos }
      );
    }
    
    // Ensure both references are on the same sheet
    const startSheet = startParsed.sheet;
    const endSheet = endParsed.sheet;
    
    if (startSheet !== endSheet) {
      throw new ParseError(
        `Range references must be on the same sheet`,
        { start: startPos, end: endPos }
      );
    }
    
    // Convert to SimpleCellRange
    const startCol = letterToColNumber(startParsed.col);
    const startRow = parseInt(startParsed.row) - 1;
    const endCol = letterToColNumber(endParsed.col);
    const endRow = parseInt(endParsed.row) - 1;
    
    if (startCol < 0 || startRow < 0 || endCol < 0 || endRow < 0) {
      throw new ParseError(
        `Invalid range reference: ${startRef}:${endRef}`,
        { start: startPos, end: endPos }
      );
    }
    
    const sheetId = startSheet ? this.getSheetId(startSheet) : this.contextSheetId;
    
    const range: SimpleCellRange = {
      start: {
        sheet: sheetId,
        col: Math.min(startCol, endCol),
        row: Math.min(startRow, endRow)
      },
      end: {
        sheet: sheetId,
        col: Math.max(startCol, endCol),
        row: Math.max(startRow, endRow)
      }
    };
    
    return createRangeNode(
      range,
      {
        start: {
          col: startParsed.colAbsolute,
          row: startParsed.rowAbsolute
        },
        end: {
          col: endParsed.colAbsolute,
          row: endParsed.rowAbsolute
        }
      },
      startPos,
      endPos
    );
  }
  
  /**
   * Parse a cell or range reference with a known sheet name
   */
  private parseCellOrRangeWithSheet(sheetName: string, startPos: number): ASTNode {
    // Build the full reference string for parsing
    let ref = sheetName.includes(' ') ? `'${sheetName}'!` : sheetName + '!';
    
    // Check for $ before column
    if (this.tokens.match('DOLLAR')) {
      ref += this.tokens.consume().value;
    }
    
    // Get the cell reference part
    if (this.tokens.match('IDENTIFIER')) {
      const identifier = this.tokens.consume();
      ref += identifier.value;
      
      // Check if this is an infinite column range (Sheet1!A:A)
      if (this.tokens.match('COLON')) {
        this.tokens.consume();
        const endRef = this.parseRangeEnd();
        // For cross-sheet ranges, do NOT prepend the sheet name to the end reference
        // parseRange will handle sheet name extraction from the start reference
        return this.parseRange(ref, endRef, startPos, this.tokens.peek().position.start);
      }
      
      // Check for $ before row or just row number
      if (this.tokens.match('DOLLAR')) {
        ref += this.tokens.consume().value;
      }
      
      if (this.tokens.match('NUMBER')) {
        ref += this.tokens.consume().value;
      }
    } else if (this.tokens.match('NUMBER')) {
      // Handle infinite row range (Sheet1!5:5)
      const number = this.tokens.consume();
      ref += number.value;
      
      // Check for range
      if (this.tokens.match('COLON')) {
        this.tokens.consume();
        const endRef = this.parseRangeEnd();
        // For cross-sheet ranges, do NOT prepend the sheet name to the end reference
        // parseRange will handle sheet name extraction from the start reference
        return this.parseRange(ref, endRef, startPos, this.tokens.peek().position.start);
      }
    } else {
      throw new ParseError(`Expected cell reference after ${sheetName}!`, this.tokens.peek().position);
    }
    
    // Check for range (normal cell range like Sheet1!A1:B2)
    if (this.tokens.match('COLON')) {
      this.tokens.consume();
      const endRef = this.parseRangeEnd();
      // For cross-sheet ranges, do NOT prepend the sheet name to the end reference
      // parseRange will handle sheet name extraction from the start reference
      return this.parseRange(ref, endRef, startPos, this.tokens.peek().position.start);
    }
    
    // Parse as single cell reference
    const cellRef = this.parseCellReferenceString(ref);
    if (cellRef) {
      return cellRef;
    }
    
    throw new ParseError(`Invalid cell reference: ${ref}`, { start: startPos, end: this.tokens.peek().position.start });
  }
  
  /**
   * Get sheet ID from sheet name
   */
  private getSheetId(sheetName: string): number {
    if (this.sheetResolver) {
      const sheetId = this.sheetResolver(sheetName);
      if (sheetId === -1) {
        throw new ParseError(`Sheet '${sheetName}' not found`);
      }
      return sheetId;
    }
    // Fallback to context sheet ID if no resolver provided
    return this.contextSheetId;
  }
}

/**
 * Parse a formula string into an AST
 */
export function parseFormula(formula: string, contextSheetId: number = 0, sheetResolver?: SheetResolver): FormulaAST {
  return Parser.parse(formula, contextSheetId, sheetResolver) as FormulaAST;
}

/**
 * Normalize a formula string (remove extra whitespace, standardize casing)
 */
export function normalizeFormula(formula: string): string {
  try {
    // Parse and reconstruct
    const ast = parseFormula(formula);
    return astToFormula(ast);
  } catch (error) {
    // If parsing fails, return original formula
    return formula;
  }
}

/**
 * Convert AST back to formula string (for normalization)
 */
export function astToFormula(node: ASTNode): string {
  switch (node.type) {
    case 'value': {
      const valueNode = node as ValueNode;
      if (valueNode.value === undefined) return '';
      if (typeof valueNode.value === 'string') return `"${valueNode.value.replace(/"/g, '""')}"`;
      if (typeof valueNode.value === 'boolean') return valueNode.value ? 'TRUE' : 'FALSE';
      return String(valueNode.value);
    }
      
        case 'reference': {
      const refNode = node as ReferenceNode;
      // TODO: Convert back to A1 notation
      return `R${refNode.address.row + 1}C${refNode.address.col + 1}`;
    }
      
    case 'range': {
      const rangeNode = node as RangeNode;
      // TODO: Convert back to A1 notation
      return `R${rangeNode.range.start.row + 1}C${rangeNode.range.start.col + 1}:R${rangeNode.range.end.row + 1}C${rangeNode.range.end.col + 1}`;
    }
      
    case 'function': {
      const funcNode = node as FunctionNode;
      const args = funcNode.args.map(astToFormula).join(',');
      return `${funcNode.name}(${args})`;
    }
      
    case 'unary-op': {
      const unaryNode = node as UnaryOpNode;
      const operand = astToFormula(unaryNode.operand);
      if (unaryNode.operator === '%') {
        return `${operand}%`;
      }
      return `${unaryNode.operator}${operand}`;
    }
      
    case 'binary-op': {
      const binaryNode = node as BinaryOpNode;
      const left = astToFormula(binaryNode.left);
      const right = astToFormula(binaryNode.right);
      return `${left}${binaryNode.operator}${right}`;
    }
      
    case 'array': {
      const arrayNode = node as ArrayNode;
      const rows = arrayNode.elements.map(row => 
        row.map(astToFormula).join(',')
      ).join(';');
      return `{${rows}}`;
    }
      
    case 'named-expression':
      return (node as NamedExpressionNode).name;
      
    case 'error':
      return (node as ErrorNode).error;
  }
}

/**
 * Validate a formula string
 */
export function validateFormula(formula: string, contextSheetId: number = 0): boolean {
  try {
    parseFormula(formula, contextSheetId);
    return true;
  } catch (error) {
    return false;
  }
}

/**
 * Extract all named expressions from a formula
 */
export function extractNamedExpressions(formula: string, contextSheetId: number = 0): string[] {
  try {
    const ast = parseFormula(formula, contextSheetId);
    return getNamedExpressionsFromAST(ast);
  } catch (error) {
    return [];
  }
}

/**
 * Get all named expressions from an AST
 */
function getNamedExpressionsFromAST(node: ASTNode): string[] {
  const names: string[] = [];
  
  function traverse(n: ASTNode) {
    switch (n.type) {
      case 'named-expression':
        names.push((n as NamedExpressionNode).name);
        break;
      case 'function':
        (n as FunctionNode).args.forEach(traverse);
        break;
      case 'unary-op':
        traverse((n as UnaryOpNode).operand);
        break;
      case 'binary-op':
        traverse((n as BinaryOpNode).left);
        traverse((n as BinaryOpNode).right);
        break;
      case 'array':
        (n as ArrayNode).elements.forEach(row => row.forEach(traverse));
        break;
    }
  }
  
  traverse(node);
  return [...new Set(names)]; // Remove duplicates
}
