import { columnToIndex, indexToColumn } from "../core/utils";
import { FormulaError, type SpreadsheetRange } from "../core/types";
import type { ASTNode, RangeNode, ReferenceNode } from "./ast";
import {
  createArrayNode,
  createBinaryOpNode,
  createEmptyNode,
  createErrorNode,
  createFunctionNode,
  createInfinityNode,
  createNamedExpressionNode,
  createRangeNode,
  createReferenceNode,
  createStructuredReferenceNode,
  createThreeDRangeNode,
  createUnaryOpNode,
  createValueNode,
} from "./ast";
import {
  getOperatorAssociativity,
  getOperatorPrecedence,
  isBinaryOperator,
  parseCellReference,
  parseInfiniteRange,
  parseOpenEndedRange,
  SPECIAL_FUNCTIONS,
} from "./grammar";
import { Lexer, TokenStream, type Token } from "./lexer";

/**
 * Parser error class
 */
export class ParseError extends Error {
  constructor(
    message: string,
    public position?: { start: number; end: number }
  ) {
    super(message);
    this.name = "ParseError";
  }
}

/**
 * Formula parser class
 */
export class Parser {
  private tokens: TokenStream;

  constructor(tokens: Token[]) {
    this.tokens = new TokenStream(tokens);
  }

  /**
   * Parse bare column reference like [Column] or [Column1:Column2]
   */
  private parseBareColumnReference(): ASTNode {
    const startPos = this.tokens.peek().position.start;
    this.tokens.consume(); // [

    let selector: "#All" | "#Data" | "#Headers" | undefined;
    let cols: { startCol: string; endCol: string } | undefined;
    let isCurrentRow = false;

    if (this.tokens.match("HASH")) {
      // Selector like [#Data]
      this.tokens.consume(); // #
      if (!this.tokens.match("IDENTIFIER")) {
        throw new ParseError(
          "Expected selector name after #",
          this.tokens.peek().position
        );
      }
      const selectorName = this.tokens.consume().value;
      selector = `#${selectorName}` as "#All" | "#Data" | "#Headers";
    } else if (this.tokens.match("IDENTIFIER")) {
      // Column reference like [Column] or [Column1:Column2]
      const colStart = this.parseColumnName();

      // Check if it's a column range Column1:Column2
      if (this.tokens.match("COLON")) {
        this.tokens.consume(); // :
        const colEnd = this.parseColumnName();

        cols = {
          startCol: colStart,
          endCol: colEnd,
        };
      } else {
        // Single column
        cols = {
          startCol: colStart,
          endCol: colStart,
        };
      }
    } else {
      throw new ParseError(
        "Expected column name or selector in bare column reference",
        this.tokens.peek().position
      );
    }

    if (!this.tokens.match("RBRACKET")) {
      throw new ParseError(
        "Expected ] to close column reference",
        this.tokens.peek().position
      );
    }
    this.tokens.consume(); // ]

    return createStructuredReferenceNode({
      tableName: undefined, // No table name for bare references
      cols,
      selector,
      isCurrentRow,
      position: {
        start: startPos,
        end: this.tokens.peek().position?.end ?? 0,
      },
    });
  }

  /**
   * Parse a column name that might consist of multiple identifiers separated by spaces or contain dashes
   */
  private parseColumnName(): string {
    if (!this.tokens.match("IDENTIFIER")) {
      throw new ParseError("Expected column name", this.tokens.peek().position);
    }

    let columnName = this.tokens.consume().value;

    // Handle multi-word column names (e.g., "Net Sales") and column names with dashes (e.g., "ORDER-ID")
    while (
      this.tokens.match("IDENTIFIER") || 
      (this.tokens.match("OPERATOR") && this.tokens.peek().value === "-")
    ) {
      if (this.tokens.match("IDENTIFIER")) {
        columnName += " " + this.tokens.consume().value;
      } else if (this.tokens.match("OPERATOR") && this.tokens.peek().value === "-") {
        // Handle dash as part of column name
        columnName += this.tokens.consume().value;
        // After a dash, we expect another identifier
        if (this.tokens.match("IDENTIFIER")) {
          columnName += this.tokens.consume().value;
        } else {
          throw new ParseError("Expected identifier after dash in column name", this.tokens.peek().position);
        }
      }
    }

    return columnName;
  }

  /**
   * Parse a formula string
   */
  static parse(formula: string): ASTNode {
    // Tokenize
    const lexer = new Lexer(formula);
    const tokens = lexer.tokenize();

    // Parse
    const parser = new Parser(tokens);
    return parser.parseFormula();
  }

  /**
   * Parse the entire formula
   */
  parseFormula(): ASTNode {
    if (this.tokens.isAtEnd()) {
      // Empty formula
      return createEmptyNode();
    }

    if (this.tokens.peek().type === "EOF") {
      // Handle edge case of just '=' with nothing after
      return createErrorNode(FormulaError.ERROR, "Empty formula");
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

      return expr;
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

      if (token.type !== "OPERATOR" || !isBinaryOperator(token.value)) {
        break;
      }

      const precedence = getOperatorPrecedence(token.value);
      if (precedence < minPrecedence) {
        break;
      }

      const operator = token.value;
      const associativity = getOperatorAssociativity(operator);
      const start = token.position.start;

      this.tokens.consume(); // Consume operator

      // For right-associative operators, use same precedence
      // For left-associative, use precedence + 1
      const nextMinPrecedence =
        associativity === "right" ? precedence : precedence + 1;
      const right = this.parseBinaryExpression(nextMinPrecedence);

      left = createBinaryOpNode(
        operator,
        left,
        right,
        right.position?.end
          ? {
              start,
              end: right.position.end,
            }
          : undefined
      );
    }

    return left;
  }

  /**
   * Parse unary expressions
   */
  private parseUnaryExpression(): ASTNode {
    const token = this.tokens.peek();

    // Check for unary operators
    if (
      token.type === "OPERATOR" &&
      (token.value === "+" || token.value === "-")
    ) {
      const start = token.position.start;
      const operator = token.value as "+" | "-";

      this.tokens.consume();
      const operand = this.parseUnaryExpression();

      return createUnaryOpNode(
        operator,
        operand,
        operand.position
          ? {
              start,
              end: operand.position.end,
            }
          : undefined
      );
    }

    return this.parsePostfixExpression();
  }

  /**
   * Parse postfix expressions (currently just %)
   */
  private parsePostfixExpression(): ASTNode {
    let expr = this.parsePrimaryExpression();

    // Check for percentage operator
    if (this.tokens.match("OPERATOR") && this.tokens.peek().value === "%") {
      const token = this.tokens.consume();
      expr = createUnaryOpNode(
        "%",
        expr,
        expr.position
          ? {
              start: expr.position.start,
              end: token.position.end,
            }
          : undefined
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
      case "NUMBER":
        return this.parseNumber();

      case "STRING":
        return this.parseString();

      case "BOOLEAN":
        return this.parseBoolean();

      case "ERROR":
        return this.parseError();

      case "FUNCTION":
        return this.parseFunctionCall();

      case "IDENTIFIER":
        return this.parseIdentifier();

      case "INFINITY":
        const infinityToken = this.tokens.consume();
        return createInfinityNode(infinityToken.position);

      case "AT":
        return this.parseCurrentRowReference();

      case "HASH":
        return this.parseTableSelector();

      case "DOLLAR":
        // Handle absolute reference starting with $
        return this.parseAbsoluteReference();

      case "LPAREN":
        return this.parseParenthesizedExpression();

      case "LBRACE":
        return this.parseArrayLiteral();

      case "LBRACKET":
        // Could be [@Column], [Column], or [#Selector] syntax
        if (this.tokens.peekAhead(1)?.type === "AT") {
          return this.parseCurrentRowReference();
        } else if (
          this.tokens.peekAhead(1)?.type === "IDENTIFIER" ||
          this.tokens.peekAhead(1)?.type === "HASH"
        ) {
          // Bare column reference like [result] or [#Data]
          return this.parseBareColumnReference();
        }
        throw new ParseError(
          `Unexpected bracket: ${token.value}`,
          token.position
        );

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
    let ref = "";

    // First $
    ref += this.tokens.consume().value;

    // Get next token
    if (this.tokens.match("IDENTIFIER")) {
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
        if (this.tokens.match("DOLLAR")) {
          ref += this.tokens.consume().value;
        }

        // Row number (optional for infinite column ranges)
        if (this.tokens.match("NUMBER")) {
          ref += this.tokens.consume().value;
        } else if (!this.tokens.match("COLON")) {
          // Only require row number if this isn't a range starting with colon
          throw new ParseError(
            "Expected row number",
            this.tokens.peek().position
          );
        }
      } else {
        throw new ParseError(
          "Invalid cell reference format",
          this.tokens.peek().position
        );
      }
    } else if (this.tokens.match("NUMBER")) {
      // Handle absolute row ranges like $1:$100
      const number = this.tokens.consume().value;
      ref += number;
    } else {
      throw new ParseError(
        "Expected column letter or row number after $",
        this.tokens.peek().position
      );
    }

    // Check for range
    if (this.tokens.match("COLON")) {
      this.tokens.consume();
      const endRef = this.parseRangeEnd();
      return this.parseRange(
        ref,
        endRef,
        start,
        this.tokens.peek().position.start
      );
    }

    // Parse as single cell reference
    const cellRef = this.parseCellReferenceString(ref);
    if (cellRef) {
      return cellRef;
    }

    throw new ParseError(`Invalid cell reference: ${ref}`, {
      start,
      end: this.tokens.peek().position.start,
    });
  }

  /**
   * Parse a number literal or row range
   */
  private parseNumber(): ASTNode {
    const token = this.tokens.peek();
    const start = token.position.start;

    // Check if this could be a row range (e.g., 5:5, 1:10)
    if (this.tokens.peekNext() && this.tokens.peekNext()!.type === "COLON") {
      // This is a row range
      const startRow = this.tokens.consume().value;
      this.tokens.consume(); // Consume ':'

      // Get the end row (handle absolute references like $5)
      let endRow: string = "";
      if (this.tokens.match("DOLLAR")) {
        endRow += this.tokens.consume().value;
      }
      if (this.tokens.match("NUMBER")) {
        endRow += this.tokens.consume().value;
      } else {
        throw new ParseError(
          "Expected row number after :",
          this.tokens.peek().position
        );
      }

      // Parse as an infinite row range
      return this.parseRange(
        startRow,
        endRow,
        start,
        this.tokens.peek().position.start
      );
    }

    // Otherwise, parse as a regular number
    this.tokens.consume();
    const value = parseFloat(token.value);

    if (isNaN(value)) {
      throw new ParseError(`Invalid number: ${token.value}`, token.position);
    }

    return createValueNode(
      {
        type: "number",
        value,
      },
      {
        start: token.position.start,
        end: token.position.end,
      }
    );
  }

  /**
   * Parse a string literal
   */
  private parseString(): ASTNode {
    const token = this.tokens.consume();
    return createValueNode(
      {
        type: "string",
        value: token.value,
      },
      {
        start: token.position.start,
        end: token.position.end,
      }
    );
  }

  /**
   * Parse a boolean literal
   */
  private parseBoolean(): ASTNode {
    const token = this.tokens.consume();
    const value = token.value.toUpperCase() === "TRUE";
    return createValueNode(
      {
        type: "boolean",
        value,
      },
      {
        start: token.position.start,
        end: token.position.end,
      }
    );
  }

  /**
   * Parse an error literal
   */
  private parseError(): ASTNode {
    const token = this.tokens.consume();
    const error = token.value as FormulaError;
    return createErrorNode(error, `Error literal: ${error}`, {
      start: token.position.start,
      end: token.position.end,
    });
  }

  /**
   * Parse a function call
   */
  private parseFunctionCall(): ASTNode {
    const nameToken = this.tokens.consume();
    const functionName = nameToken.value;
    const start = nameToken.position.start;

    // Check for special functions that don't require parentheses
    if (
      SPECIAL_FUNCTIONS.has(functionName.toUpperCase()) &&
      !this.tokens.match("LPAREN")
    ) {
      return createFunctionNode(functionName, [], {
        start: start,
        end: nameToken.position.end,
      });
    }

    // Expect opening parenthesis
    if (!this.tokens.match("LPAREN")) {
      throw new ParseError(
        `Expected '(' after function name ${functionName}`,
        this.tokens.peek().position
      );
    }

    this.tokens.consume(); // Consume '('

    // Parse arguments
    const args: ASTNode[] = [];

    // Handle empty argument list
    if (this.tokens.match("RPAREN")) {
      const rparenToken = this.tokens.consume();
      const endPos = rparenToken.position.end;
      const node = createFunctionNode(functionName, args, {
        start: start,
        end: endPos,
      });

      return node;
    }

    // Parse arguments
    while (true) {
      args.push(this.parseExpression());

      if (this.tokens.match("COMMA")) {
        this.tokens.consume();
        // Continue parsing next argument
      } else if (this.tokens.match("RPAREN")) {
        const rparenToken = this.tokens.consume();
        const end = rparenToken.position.end;
        const node = createFunctionNode(functionName, args, {
          start: start,
          end: end,
        });

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
      "Unexpected end of function argument parsing",
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

    // Check if this could be a 3D range (Sheet1:Sheet3!)
    // To distinguish from normal ranges, we check:
    // 1. If there's a colon
    // 2. If what follows the colon is an identifier
    // 3. If that identifier is followed by !
    if (this.tokens.match("COLON")) {
      const colonPos = this.tokens.getPosition();

      // Look ahead to see if this looks like a 3D range
      const nextToken = this.tokens.peekAhead(1);
      const tokenAfterNext = this.tokens.peekAhead(2);

      if (
        nextToken?.type === "IDENTIFIER" &&
        tokenAfterNext?.type === "EXCLAMATION"
      ) {
        // This looks like a 3D range, consume the tokens
        this.tokens.consume(); // :
        const endSheetToken = this.tokens.consume(); // sheet name
        const endSheet = endSheetToken.value;
        this.tokens.consume(); // !

        // Extract sheet names
        let startSheet = value;
        if (startSheet.startsWith("'") && startSheet.endsWith("'")) {
          startSheet = startSheet.slice(1, -1).replace(/''/g, "'");
        }

        let endSheetName = endSheet;
        if (endSheetName.startsWith("'") && endSheetName.endsWith("'")) {
          endSheetName = endSheetName.slice(1, -1).replace(/''/g, "'");
        }

        // Parse the reference after the sheet range
        const ref = this.parseCellOrRangeAfterSheets();
        return createThreeDRangeNode(
          startSheet,
          endSheetName,
          ref as ReferenceNode | RangeNode,
          {
            start,
            end: this.tokens.peek().position?.end ?? 0,
          }
        );
      }
      // Not a 3D range, continue with normal parsing
    }

    // Check if this is a sheet reference (Sheet1! or 'My Sheet'!)
    if (this.tokens.match("EXCLAMATION")) {
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

    // Check if this is a table reference (Table1[Column1])
    if (this.tokens.match("LBRACKET")) {
      return this.parseTableReference(value, start);
    }

    // Check if this could be part of a cell reference (e.g., D in D$4)
    if (this.tokens.match("DOLLAR") && this.isColumnIdentifier(value)) {
      // This is a mixed reference like D$4
      value += this.tokens.consume().value; // Add $

      // Get the row number
      if (this.tokens.match("NUMBER")) {
        value += this.tokens.consume().value;
      } else {
        throw new ParseError(
          "Expected row number after $",
          this.tokens.peek().position
        );
      }
    } else if (this.tokens.match("NUMBER") && this.isColumnIdentifier(value)) {
      // Regular cell reference like D4
      value += this.tokens.consume().value;
    }

    // Check for colon (range operator)
    if (this.tokens.match("COLON")) {
      this.tokens.consume(); // Consume ':'

      // Parse end of range with potential $ signs
      const endStart = this.tokens.peek().position.start;
      const endRef = this.parseRangeEnd();
      const endPos = this.tokens.peek().position.start;

      // Parse as range
      return this.parseRange(value, endRef, start, endPos);
    }

    // Try to parse as cell reference
    const parsedCellRef = this.parseCellReferenceString(value);
    if (parsedCellRef) {
      return parsedCellRef;
    }

    // Otherwise, it's a named expression
    return createNamedExpressionNode(value, {
      start: start,
      end: token.position.end,
    });
  }

  /**
   * Check if a string is a valid column identifier (A-Z, AA-ZZ, etc.)
   */
  private isColumnIdentifier(str: string): boolean {
    return /^[A-Z]+$/i.test(str);
  }

  /**
   * Parse the end part of a range (handling $ signs, infinite ranges, and open-ended ranges)
   */
  private parseRangeEnd(): string {
    let result = "";

    // Check for INFINITY token (for A5:INFINITY syntax)
    if (this.tokens.match("INFINITY")) {
      result += this.tokens.consume().value;
      return result;
    }

    // Check for $ before column or row
    if (this.tokens.match("DOLLAR")) {
      result += this.tokens.consume().value;
    }

    // Get identifier part (for column) or number part (for row)
    if (this.tokens.match("IDENTIFIER")) {
      result += this.tokens.consume().value;

      // For normal ranges, check for $ before row
      if (this.tokens.match("DOLLAR")) {
        result += this.tokens.consume().value;
      }

      // Get number part if present (normal range)
      if (this.tokens.match("NUMBER")) {
        result += this.tokens.consume().value;
      }
      // If no number, it's an infinite column range (e.g., A:A) or open-ended range (e.g., A5:D)
    } else if (this.tokens.match("NUMBER")) {
      // This handles cases like:
      // - Infinite row range (e.g., 5:5) 
      // - Open-ended range (e.g., A5:15)
      // - Absolute row reference (e.g., A5:$15)
      result += this.tokens.consume().value;
    } else {
      throw new ParseError(
        "Expected cell reference, column, row, or INFINITY after :",
        this.tokens.peek().position
      );
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

    if (!this.tokens.match("RPAREN")) {
      throw new ParseError(`Expected ')'`, this.tokens.peek().position);
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
    if (this.tokens.match("RBRACE")) {
      this.tokens.consume();
      return createArrayNode([[createEmptyNode()]], {
        start: start,
        end: this.tokens.peek().position.start,
      });
    }

    // Parse array elements
    while (true) {
      currentRow.push(this.parseExpression());

      if (this.tokens.match("COMMA")) {
        this.tokens.consume();
        // Continue current row
      } else if (this.tokens.match("SEMICOLON")) {
        this.tokens.consume();
        // Start new row
        rows.push(currentRow);
        currentRow = [];
      } else if (this.tokens.match("RBRACE")) {
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
          throw new ParseError(`Inconsistent row lengths in array literal`, {
            start,
            end: this.tokens.peek().position.start,
          });
        }
      }
    }

    return createArrayNode(rows, {
      start: start,
      end: this.tokens.peek().position.start,
    });
  }

  /**
   * Parse a cell reference string
   */
  private parseCellReferenceString(value: string): ReferenceNode | null {
    const parsed = parseCellReference(value);
    if (!parsed) {
      return null;
    }

    // Convert to SimpleCellAddress
    const colNum = columnToIndex(parsed.col);
    const rowNum = parseInt(parsed.row) - 1; // Convert to 0-based

    if (colNum < 0 || rowNum < 0) {
      return null;
    }

    return createReferenceNode({
      address: {
        colIndex: colNum,
        rowIndex: rowNum,
      },
      isAbsolute: {
        col: parsed.colAbsolute,
        row: parsed.rowAbsolute,
      },
      sheetName: parsed.sheet,
    });
  }

  /**
   * Parse table reference (e.g., Table1[Column1])
   */
  private parseTableReference(tableName: string, startPos: number): ASTNode {
    this.tokens.consume(); // [

    let selector: "#All" | "#Data" | "#Headers" | undefined;
    let cols: { startCol: string; endCol: string } | undefined;
    let isCurrentRow = false;

    // Check for complex syntax like [[#Headers],[Column1]] or [[Column1]:[Column2]]
    if (this.tokens.match("LBRACKET")) {
      this.tokens.consume(); // second [

      if (this.tokens.match("HASH")) {
        this.tokens.consume(); // #
        if (!this.tokens.match("IDENTIFIER")) {
          throw new ParseError(
            "Expected selector name after #",
            this.tokens.peek().position
          );
        }
        const selectorName = this.tokens.consume().value;
        selector = `#${selectorName}` as "#All" | "#Data" | "#Headers"

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after selector",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]

        // Check for column part
        if (this.tokens.match("COMMA")) {
          this.tokens.consume(); // ,

          if (!this.tokens.match("LBRACKET")) {
            throw new ParseError(
              "Expected [ after comma",
              this.tokens.peek().position
            );
          }
          this.tokens.consume(); // [

          // Parse column specification that could be either Column1 or Column1:Column2
          if (!this.tokens.match("IDENTIFIER")) {
            throw new ParseError(
              "Expected column name",
              this.tokens.peek().position
            );
          }
          const colStart = this.parseColumnName();

          // Check if it's a column range Column1:Column2
          if (this.tokens.match("COLON")) {
            this.tokens.consume(); // :

            if (!this.tokens.match("IDENTIFIER")) {
              throw new ParseError(
                "Expected end column name after :",
                this.tokens.peek().position
              );
            }
            const colEnd = this.parseColumnName();

            cols = {
              startCol: colStart,
              endCol: colEnd,
            };
          } else {
            // Single column
            cols = {
              startCol: colStart,
              endCol: colStart,
            };
          }

          if (!this.tokens.match("RBRACKET")) {
            throw new ParseError(
              "Expected ] after column specification",
              this.tokens.peek().position
            );
          }
          this.tokens.consume(); // ]
        }

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] to close table reference",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // outer ]
      } else if (this.tokens.match("IDENTIFIER")) {
        // Handle [[Column1]:[Column2]] syntax
        const colStart = this.parseColumnName();

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after column name",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]

        if (!this.tokens.match("COLON")) {
          throw new ParseError(
            "Expected : after first column in [[Column1]:[Column2]]",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // :

        if (!this.tokens.match("LBRACKET")) {
          throw new ParseError(
            "Expected [ before second column name",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // [

        const colEnd = this.parseColumnName();

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after second column name",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] to close table reference",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // outer ]

        cols = {
          startCol: colStart,
          endCol: colEnd,
        };
      }
    } else if (this.tokens.match("AT")) {
      // Current row reference like Table1[@Column]
      this.tokens.consume(); // @
      isCurrentRow = true;

      if (this.tokens.match("LBRACKET")) {
        // Handle [@[Column]] or [@[Column1]:[Column2]] syntax
        this.tokens.consume(); // [

        const colStart = this.parseColumnName();

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after column name",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]

        // Check if this is a column range [@[Column1]:[Column2]]
        if (this.tokens.match("COLON")) {
          this.tokens.consume(); // :

          if (!this.tokens.match("LBRACKET")) {
            throw new ParseError(
              "Expected [ before second column name",
              this.tokens.peek().position
            );
          }
          this.tokens.consume(); // [

          const colEnd = this.parseColumnName();

          if (!this.tokens.match("RBRACKET")) {
            throw new ParseError(
              "Expected ] after second column name",
              this.tokens.peek().position
            );
          }
          this.tokens.consume(); // ]

          cols = {
            startCol: colStart,
            endCol: colEnd,
          };
        } else {
          // Single column [@[Column]]
          cols = {
            startCol: colStart,
            endCol: colStart,
          };
        }

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] to close table reference",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]
      } else {
        // Handle [@Column] or [@Column1:Column2] syntax (simple current row)
        const colStart = this.parseColumnName();

        // Check if it's a column range Column1:Column2
        if (this.tokens.match("COLON")) {
          this.tokens.consume(); // :

          const colEnd = this.parseColumnName();

          cols = {
            startCol: colStart,
            endCol: colEnd,
          };
        } else {
          // Single column
          cols = {
            startCol: colStart,
            endCol: colStart,
          };
        }

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after column reference",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]
      }
    } else if (this.tokens.match("IDENTIFIER")) {
      // Simple column reference like Table1[Column1] or range Table1[Column1:Column2]
      const colStart = this.parseColumnName();

      // Check if it's a column range Column1:Column2
      if (this.tokens.match("COLON")) {
        this.tokens.consume(); // :

        const colEnd = this.parseColumnName();

        cols = {
          startCol: colStart,
          endCol: colEnd,
        };
      } else {
        // Single column
        cols = {
          startCol: colStart,
          endCol: colStart,
        };
      }

      if (!this.tokens.match("RBRACKET")) {
        throw new ParseError(
          "Expected ] after column name",
          this.tokens.peek().position
        );
      }
      this.tokens.consume(); // ]
    } else if (this.tokens.match("HASH")) {
      // Simple selector reference like Table1[#Data]
      this.tokens.consume(); // #

      if (!this.tokens.match("IDENTIFIER")) {
        throw new ParseError(
          "Expected selector name after #",
          this.tokens.peek().position
        );
      }
      const selectorName = this.tokens.consume().value;
      selector = `#${selectorName}` as "#All" | "#Data" | "#Headers";

      if (!this.tokens.match("RBRACKET")) {
        throw new ParseError(
          "Expected ] after selector",
          this.tokens.peek().position
        );
      }
      this.tokens.consume(); // ]
    } else {
      throw new ParseError(
        "Expected column name or selector in table reference",
        this.tokens.peek().position
      );
    }

    return createStructuredReferenceNode({
      tableName,
      cols,
      selector,
      isCurrentRow,
      position: {
        start: startPos,
        end: this.tokens.peek().position?.end ?? 0,
      },
    });
  }

  /**
   * Parse table reference with sheet name (e.g., Sheet1!Table1[Column1])
   */
  private parseTableReferenceWithSheet(
    tableName: string,
    sheetName: string,
    startPos: number
  ): ASTNode {
    this.tokens.consume(); // [

    let selector: "#All" | "#Data" | "#Headers" | undefined;
    let cols: { startCol: string; endCol: string } | undefined;
    let isCurrentRow = false;

    // Check for complex syntax like [[#Headers],[Column1]] or [[Column1]:[Column2]]
    if (this.tokens.match("LBRACKET")) {
      this.tokens.consume(); // second [

      if (this.tokens.match("HASH")) {
        this.tokens.consume(); // #
        if (!this.tokens.match("IDENTIFIER")) {
          throw new ParseError(
            "Expected selector name after #",
            this.tokens.peek().position
          );
        }
        const selectorName = this.tokens.consume().value;
        selector = `#${selectorName}` as "#All" | "#Data" | "#Headers";

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after selector",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]

        // Check for column part
        if (this.tokens.match("COMMA")) {
          this.tokens.consume(); // ,

          if (!this.tokens.match("LBRACKET")) {
            throw new ParseError(
              "Expected [ after comma",
              this.tokens.peek().position
            );
          }
          this.tokens.consume(); // [

          // Parse column specification that could be either Column1 or Column1:Column2
          if (!this.tokens.match("IDENTIFIER")) {
            throw new ParseError(
              "Expected column name",
              this.tokens.peek().position
            );
          }
          const colStart = this.parseColumnName();

          // Check if it's a column range Column1:Column2
          if (this.tokens.match("COLON")) {
            this.tokens.consume(); // :

            if (!this.tokens.match("IDENTIFIER")) {
              throw new ParseError(
                "Expected end column name after :",
                this.tokens.peek().position
              );
            }
            const colEnd = this.parseColumnName();

            cols = {
              startCol: colStart,
              endCol: colEnd,
            };
          } else {
            // Single column
            cols = {
              startCol: colStart,
              endCol: colStart,
            };
          }

          if (!this.tokens.match("RBRACKET")) {
            throw new ParseError(
              "Expected ] after column specification",
              this.tokens.peek().position
            );
          }
          this.tokens.consume(); // ]
        }

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] to close table reference",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // outer ]
      } else if (this.tokens.match("IDENTIFIER")) {
        // Handle [[Column1]:[Column2]] syntax
        const colStart = this.parseColumnName();

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after column name",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]

        if (!this.tokens.match("COLON")) {
          throw new ParseError(
            "Expected : after first column in [[Column1]:[Column2]]",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // :

        if (!this.tokens.match("LBRACKET")) {
          throw new ParseError(
            "Expected [ before second column name",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // [

        const colEnd = this.parseColumnName();

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after second column name",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] to close table reference",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // outer ]

        cols = {
          startCol: colStart,
          endCol: colEnd,
        };
      }
    } else if (this.tokens.match("AT")) {
      // Current row reference like Table1[@Column]
      this.tokens.consume(); // @
      isCurrentRow = true;

      if (this.tokens.match("LBRACKET")) {
        // Handle [@[Column]] or [@[Column1]:[Column2]] syntax
        this.tokens.consume(); // [

        const colStart = this.parseColumnName();

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after column name",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]

        // Check if this is a column range [@[Column1]:[Column2]]
        if (this.tokens.match("COLON")) {
          this.tokens.consume(); // :

          if (!this.tokens.match("LBRACKET")) {
            throw new ParseError(
              "Expected [ before second column name",
              this.tokens.peek().position
            );
          }
          this.tokens.consume(); // [

          const colEnd = this.parseColumnName();

          if (!this.tokens.match("RBRACKET")) {
            throw new ParseError(
              "Expected ] after second column name",
              this.tokens.peek().position
            );
          }
          this.tokens.consume(); // ]

          cols = {
            startCol: colStart,
            endCol: colEnd,
          };
        } else {
          // Single column [@[Column]]
          cols = {
            startCol: colStart,
            endCol: colStart,
          };
        }

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] to close table reference",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]
      } else {
        // Handle [@Column] or [@Column1:Column2] syntax (simple current row)
        const colStart = this.parseColumnName();

        // Check if it's a column range Column1:Column2
        if (this.tokens.match("COLON")) {
          this.tokens.consume(); // :

          const colEnd = this.parseColumnName();

          cols = {
            startCol: colStart,
            endCol: colEnd,
          };
        } else {
          // Single column
          cols = {
            startCol: colStart,
            endCol: colStart,
          };
        }

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] after column reference",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]
      }
    } else if (this.tokens.match("IDENTIFIER")) {
      // Simple column reference like Table1[Column1] or range Table1[Column1:Column2]
      const colStart = this.parseColumnName();

      // Check if it's a column range Column1:Column2
      if (this.tokens.match("COLON")) {
        this.tokens.consume(); // :

        const colEnd = this.parseColumnName();

        cols = {
          startCol: colStart,
          endCol: colEnd,
        };
      } else {
        // Single column
        cols = {
          startCol: colStart,
          endCol: colStart,
        };
      }

      if (!this.tokens.match("RBRACKET")) {
        throw new ParseError(
          "Expected ] after column name",
          this.tokens.peek().position
        );
      }
      this.tokens.consume(); // ]
    } else if (this.tokens.match("HASH")) {
      // Simple selector reference like Table1[#Data]
      this.tokens.consume(); // #

      if (!this.tokens.match("IDENTIFIER")) {
        throw new ParseError(
          "Expected selector name after #",
          this.tokens.peek().position
        );
      }
      const selectorName = this.tokens.consume().value;
      selector = `#${selectorName}` as "#All" | "#Data" | "#Headers";

      if (!this.tokens.match("RBRACKET")) {
        throw new ParseError(
          "Expected ] after selector",
          this.tokens.peek().position
        );
      }
      this.tokens.consume(); // ]
    } else {
      throw new ParseError(
        "Expected column name or selector in table reference",
        this.tokens.peek().position
      );
    }

    return createStructuredReferenceNode({
      tableName,
      cols,
      selector,
      isCurrentRow,
      position: {
        start: startPos,
        end: this.tokens.peek().position?.end ?? 0,
      },
    });
  }

  /**
   * Parse cell or range after sheet range in 3D reference
   */
  private parseCellOrRangeAfterSheets(): RangeNode | ReferenceNode {
    const start = this.tokens.peek().position?.start ?? 0;

    // Could be a simple cell reference like A1, or a range like A1:B2
    if (this.tokens.match("IDENTIFIER")) {
      const firstIdent = this.tokens.consume().value;

      // Check if it's a complete cell reference (e.g., A1)
      const cellRef = this.parseCellReferenceString(firstIdent);
      if (cellRef) {
        // Check if this is part of a range
        if (this.tokens.match("COLON")) {
          this.tokens.consume(); // :

          // Parse end of range
          const endRef = this.parseRangeEnd();
          const endPos = this.tokens.peek().position?.end ?? 0;

          // Parse as range without sheet (sheet is handled by 3D range)
          const parsed = this.parseRange(firstIdent, endRef, start, endPos);

          // Remove sheet from range node if present
          if (parsed.type === "range") {
            parsed.sheetName = undefined;
          }

          return parsed;
        }

        // Single cell reference
        if (cellRef.type === "reference") {
          cellRef.sheetName = undefined;
        }
        return cellRef;
      }
    }

    // Handle absolute references starting with $
    if (this.tokens.match("DOLLAR")) {
      const dollarStart = this.tokens.getPosition();
      this.tokens.consume(); // $

      let ref = "$";

      // Get column
      if (this.tokens.match("IDENTIFIER")) {
        const col = this.tokens.consume().value;
        if (this.isColumnIdentifier(col)) {
          ref += col;
        }
      }

      // Check for row absolute
      if (this.tokens.match("DOLLAR")) {
        ref += this.tokens.consume().value;
      }

      // Get row number
      if (this.tokens.match("NUMBER")) {
        ref += this.tokens.consume().value;
      }

      // Check if this is a range
      if (this.tokens.match("COLON")) {
        this.tokens.consume(); // :

        // Parse end of range
        const endRef = this.parseRangeEnd();
        const endPos = this.tokens.peek().position?.end ?? 0;

        // Parse as range without sheet (sheet is handled by 3D range)
        const parsed = this.parseRange(ref, endRef, start, endPos);

        // Remove sheet from range node if present
        if (parsed.type === "range") {
          parsed.sheetName = undefined;
        }

        return parsed;
      }

      // Single cell reference
      const cellRef = this.parseCellReferenceString(ref);
      if (!cellRef) {
        throw new ParseError(
          `Invalid cell reference: ${ref}`,
          this.tokens.peek().position
        );
      }

      // Remove sheet from reference node if present
      if (cellRef.type === "reference") {
        cellRef.sheetName = undefined;
      }

      return cellRef;
    }

    throw new ParseError(
      "Expected cell or range reference after sheet range",
      this.tokens.peek().position
    );
  }

  /**
   * Parse current row reference (e.g., [@Column], [@[Column Name]], or just @Column)
   */
  private parseCurrentRowReference(): ASTNode {
    const start = this.tokens.getPosition();

    // Handle [@Column] or [@[Column Name]] format
    if (this.tokens.match("LBRACKET")) {
      this.tokens.consume(); // [

      if (!this.tokens.match("AT")) {
        throw new ParseError("Expected @ after [", this.tokens.peek().position);
      }
      this.tokens.consume(); // @

      let columnName: string;

      // Check if we have double brackets [@[Column Name]]
      if (this.tokens.match("LBRACKET")) {
        this.tokens.consume(); // [
        columnName = this.parseColumnName();

        if (!this.tokens.match("RBRACKET")) {
          throw new ParseError(
            "Expected ] to close column name",
            this.tokens.peek().position
          );
        }
        this.tokens.consume(); // ]
      } else {
        // Single bracket format [@Column] or [@Column1:Column2]
        const colStart = this.parseColumnName();

        // Check if this is a column range [@Column1:Column2]
        if (this.tokens.match("COLON")) {
          this.tokens.consume(); // :
          const colEnd = this.parseColumnName();

          if (!this.tokens.match("RBRACKET")) {
            throw new ParseError(
              "Expected ] after column range",
              this.tokens.peek().position
            );
          }
          this.tokens.consume(); // ]

          return createStructuredReferenceNode({
            tableName: undefined,
            cols: {
              startCol: colStart,
              endCol: colEnd,
            },
            isCurrentRow: true,
            position: {
              start: this.tokens.getTokens()[start]?.position?.start ?? 0,
              end: this.tokens.peek().position?.end ?? 0,
            },
          });
        } else {
          // Single column [@Column]
          columnName = colStart;
        }
      }

      if (!this.tokens.match("RBRACKET")) {
        throw new ParseError(
          "Expected ] after column reference",
          this.tokens.peek().position
        );
      }
      this.tokens.consume(); // ]

      return createStructuredReferenceNode({
        tableName: undefined,
        cols: {
          startCol: columnName,
          endCol: columnName,
        },
        isCurrentRow: true,
        position: {
          start: this.tokens.getTokens()[start]?.position?.start ?? 0,
          end: this.tokens.peek().position?.end ?? 0,
        },
      });
    }

    // Handle @Column format (without brackets)
    this.tokens.consume(); // @
    const columnName = this.parseColumnName();

    return createStructuredReferenceNode({
      tableName: undefined,
      cols: {
        startCol: columnName,
        endCol: columnName,
      },
      isCurrentRow: true,
      position: {
        start: this.tokens.getTokens()[start]?.position?.start ?? 0,
        end: this.tokens.peek().position?.end ?? 0,
      },
    });
  }

  /**
   * Parse table selector (e.g., #Headers, #Data)
   */
  private parseTableSelector(): ASTNode {
    const hashToken = this.tokens.consume(); // #

    if (!this.tokens.match("IDENTIFIER")) {
      throw new ParseError(
        "Expected selector name after #",
        this.tokens.peek().position
      );
    }

    const selectorName = this.tokens.consume().value;
    const selector = `#${selectorName}`;

    // Check if it's a valid selector
    if (!["#All", "#Data", "#Headers"].includes(selector.toUpperCase())) {
      throw new ParseError(
        `Invalid table selector: ${selector}`,
        hashToken.position
      );
    }

    // For now, return an error node as selectors need to be part of a table reference
    throw new ParseError(
      "Table selector must be part of a table reference",
      hashToken.position
    );
  }

  /**
   * Parse a range reference (including infinite ranges)
   */
  private parseRange(
    startRef: string,
    endRef: string,
    startPos: number,
    endPos: number
  ): RangeNode {
    // For cross-sheet ranges, handle the case where only startRef includes the sheet
    let fullRange = `${startRef}:${endRef}`;
    let sheetName: string | undefined;

    // Check if start has a sheet prefix
    const sheetMatch = startRef.match(
      /^(?:([A-Za-z_][A-Za-z0-9_]*)|'([^']+)')!/
    );
    if (sheetMatch) {
      sheetName = sheetMatch[1] || sheetMatch[2];

      // First, try to parse as an infinite range or open-ended range without modifying the range
      // This handles cases like Sheet1!A:A, Sheet1!5:5, Sheet1!A5:INFINITY, etc.
      const infiniteTest = parseInfiniteRange(fullRange);
      const openEndedTest = parseOpenEndedRange(fullRange);
      if (infiniteTest || openEndedTest) {
        // It's an infinite or open-ended range, skip to processing it below
      } else {
        // Not an infinite or open-ended range, so handle normal cross-sheet ranges
        const endSheetMatch = endRef.match(
          /^(?:([A-Za-z_][A-Za-z0-9_]*)|'([^']+)')!/
        );
        if (!endSheetMatch && sheetName) {
          // Normal case: prepend sheet name to endRef for consistent parsing
          const quotedSheetName = sheetName.includes(" ")
            ? `'${sheetName}'`
            : sheetName;
          endRef = `${quotedSheetName}!${endRef}`;
          fullRange = `${startRef}:${endRef}`;
        }
      }
    }

    // Try to parse as an infinite range
    const infiniteParsed = parseInfiniteRange(fullRange);

    if (infiniteParsed) {
      // Handle infinite range
      sheetName = sheetName ?? infiniteParsed.sheet;

      if (infiniteParsed.type === "column") {
        // Infinite column range (e.g., A:A, B:D)
        const startCol = columnToIndex(infiniteParsed.start);
        const endCol = columnToIndex(infiniteParsed.end);

        if (startCol < 0 || endCol < 0) {
          throw new ParseError(`Invalid column range: ${startRef}:${endRef}`, {
            start: startPos,
            end: endPos,
          });
        }

        // Use special row values to indicate infinite range
        // We'll use -1 to indicate infinite
        const range: SpreadsheetRange = {
          start: {
            col: Math.min(startCol, endCol),
            row: 0, // Start from row 0
          },
          end: {
            col: {
              type: "number",
              value: Math.max(startCol, endCol),
            },
            row: {
              type: "infinity",
              sign: "positive",
            },
          },
        };

        return createRangeNode({
          sheetName,
          range,
          isAbsolute: {
            start: {
              col: infiniteParsed.startAbsolute,
              row: false,
            },
            end: {
              col: infiniteParsed.endAbsolute,
              row: false,
            },
          },
          position: {
            start: startPos,
            end: endPos,
          },
        });
      } else {
        // Infinite row range (e.g., 5:5, 1:10)
        const startRow = parseInt(infiniteParsed.start) - 1;
        const endRow = parseInt(infiniteParsed.end) - 1;

        if (startRow < 0 || endRow < 0) {
          throw new ParseError(`Invalid row range: ${startRef}:${endRef}`, {
            start: startPos,
            end: endPos,
          });
        }

        const range: SpreadsheetRange = {
          start: {
            col: 0, // Start from column 0
            row: Math.min(startRow, endRow),
          },
          end: {
            col: {
              type: "infinity",
              sign: "positive",
            },
            row: {
              type: "number",
              value: Math.max(startRow, endRow),
            },
          },
        };

        return createRangeNode({
          sheetName,
          range,
          isAbsolute: {
            start: {
              col: false,
              row: infiniteParsed.startAbsolute,
            },
            end: {
              col: false,
              row: infiniteParsed.endAbsolute,
            },
          },
          position: {
            start: startPos,
            end: endPos,
          },
        });
      }
    }

    // Try to parse as an open-ended range (A5:INFINITY, A5:D, A5:15)
    const openEndedParsed = parseOpenEndedRange(fullRange);
    
    if (openEndedParsed) {
      // Handle open-ended range
      sheetName = sheetName ?? openEndedParsed.sheet;
      
      const startCol = columnToIndex(openEndedParsed.startCol);
      const startRow = parseInt(openEndedParsed.startRow) - 1; // Convert to 0-based
      
      if (startCol < 0 || startRow < 0) {
        throw new ParseError(`Invalid range reference: ${startRef}:${endRef}`, {
          start: startPos,
          end: endPos,
        });
      }
      
      if (openEndedParsed.type === "infinity") {
        // A5:INFINITY - both row and column unbounded
        const range: SpreadsheetRange = {
          start: {
            col: startCol,
            row: startRow,
          },
          end: {
            col: {
              type: "infinity",
              sign: "positive",
            },
            row: {
              type: "infinity",
              sign: "positive",
            },
          },
        };
        
        return createRangeNode({
          sheetName,
          range,
          isAbsolute: {
            start: {
              col: openEndedParsed.startColAbsolute,
              row: openEndedParsed.startRowAbsolute,
            },
            end: {
              col: false, // INFINITY is never absolute
              row: false, // INFINITY is never absolute
            },
          },
          position: {
            start: startPos,
            end: endPos,
          },
        });
      } else if (openEndedParsed.type === "column-bounded") {
        // A5:D - open down only (bounded columns, unbounded rows)
        const endCol = columnToIndex(openEndedParsed.endCol!);
        
        if (endCol < 0) {
          throw new ParseError(`Invalid column range: ${startRef}:${endRef}`, {
            start: startPos,
            end: endPos,
          });
        }
        
        const range: SpreadsheetRange = {
          start: {
            col: Math.min(startCol, endCol),
            row: startRow,
          },
          end: {
            col: {
              type: "number",
              value: Math.max(startCol, endCol),
            },
            row: {
              type: "infinity",
              sign: "positive",
            },
          },
        };
        
        return createRangeNode({
          sheetName,
          range,
          isAbsolute: {
            start: {
              col: openEndedParsed.startColAbsolute,
              row: openEndedParsed.startRowAbsolute,
            },
            end: {
              col: openEndedParsed.endColAbsolute!,
              row: false, // Row is infinite, so not absolute
            },
          },
          position: {
            start: startPos,
            end: endPos,
          },
        });
      } else if (openEndedParsed.type === "row-bounded") {
        // A5:15 - open right only (bounded rows, unbounded columns)
        const endRow = parseInt(openEndedParsed.endRow!) - 1; // Convert to 0-based
        
        if (endRow < 0) {
          throw new ParseError(`Invalid row range: ${startRef}:${endRef}`, {
            start: startPos,
            end: endPos,
          });
        }
        
        const range: SpreadsheetRange = {
          start: {
            col: startCol,
            row: Math.min(startRow, endRow),
          },
          end: {
            col: {
              type: "infinity",
              sign: "positive",
            },
            row: {
              type: "number",
              value: Math.max(startRow, endRow),
            },
          },
        };
        
        return createRangeNode({
          sheetName,
          range,
          isAbsolute: {
            start: {
              col: openEndedParsed.startColAbsolute,
              row: openEndedParsed.startRowAbsolute,
            },
            end: {
              col: false, // Column is infinite, so not absolute
              row: openEndedParsed.endRowAbsolute!,
            },
          },
          position: {
            start: startPos,
            end: endPos,
          },
        });
      }
    }

    // Otherwise, parse as normal range
    const startParsed = parseCellReference(startRef);
    const endParsed = parseCellReference(endRef);

    if (!startParsed || !endParsed) {
      throw new ParseError(`Invalid range reference: ${startRef}:${endRef}`, {
        start: startPos,
        end: endPos,
      });
    }

    // Ensure both references are on the same sheet
    const startSheet = startParsed.sheet;
    const endSheet = endParsed.sheet;

    if (startSheet !== endSheet) {
      throw new ParseError(`Range references must be on the same sheet`, {
        start: startPos,
        end: endPos,
      });
    }

    sheetName = sheetName ?? startSheet;

    // Convert to SimpleCellRange
    const startCol = columnToIndex(startParsed.col);
    const startRow = parseInt(startParsed.row) - 1;
    const endCol = columnToIndex(endParsed.col);
    const endRow = parseInt(endParsed.row) - 1;

    if (startCol < 0 || startRow < 0 || endCol < 0 || endRow < 0) {
      throw new ParseError(`Invalid range reference: ${startRef}:${endRef}`, {
        start: startPos,
        end: endPos,
      });
    }

    const range: SpreadsheetRange = {
      start: {
        col: Math.min(startCol, endCol),
        row: Math.min(startRow, endRow),
      },
      end: {
        col: {
          type: "number",
          value: Math.max(startCol, endCol),
        },
        row: {
          type: "number",
          value: Math.max(startRow, endRow),
        },
      },
    };

    return createRangeNode({
      sheetName,
      range,
      isAbsolute: {
        start: {
          col: startParsed.colAbsolute,
          row: startParsed.rowAbsolute,
        },
        end: {
          col: endParsed.colAbsolute,
          row: endParsed.rowAbsolute,
        },
      },
      position: {
        start: startPos,
        end: endPos,
      },
    });
  }

  /**
   * Parse a cell or range reference with a known sheet name
   */
  private parseCellOrRangeWithSheet(
    sheetName: string,
    startPos: number
  ): ASTNode {
    // Build the full reference string for parsing
    let ref = sheetName.includes(" ") ? `'${sheetName}'!` : sheetName + "!";

    // Check for $ before column
    if (this.tokens.match("DOLLAR")) {
      ref += this.tokens.consume().value;
    }

    // Get the cell reference part
    if (this.tokens.match("IDENTIFIER")) {
      const identifier = this.tokens.consume();
      ref += identifier.value;

      // Check if this is a table reference (Sheet1!Table1[...])
      if (this.tokens.match("LBRACKET")) {
        // This is a table reference with sheet name
        return this.parseTableReferenceWithSheet(
          identifier.value,
          sheetName,
          startPos
        );
      }

      // Check if this is an infinite column range (Sheet1!A:A)
      if (this.tokens.match("COLON")) {
        this.tokens.consume();
        const endRef = this.parseRangeEnd();
        // For cross-sheet ranges, do NOT prepend the sheet name to the end reference
        // parseRange will handle sheet name extraction from the start reference
        return this.parseRange(
          ref,
          endRef,
          startPos,
          this.tokens.peek().position.start
        );
      }

      // Check for $ before row or just row number
      if (this.tokens.match("DOLLAR")) {
        ref += this.tokens.consume().value;
      }

      if (this.tokens.match("NUMBER")) {
        ref += this.tokens.consume().value;
      }
    } else if (this.tokens.match("NUMBER")) {
      // Handle infinite row range (Sheet1!5:5)
      const number = this.tokens.consume();
      ref += number.value;

      // Check for range
      if (this.tokens.match("COLON")) {
        this.tokens.consume();
        const endRef = this.parseRangeEnd();
        // For cross-sheet ranges, do NOT prepend the sheet name to the end reference
        // parseRange will handle sheet name extraction from the start reference
        return this.parseRange(
          ref,
          endRef,
          startPos,
          this.tokens.peek().position.start
        );
      }
    } else {
      throw new ParseError(
        `Expected cell reference after ${sheetName}!`,
        this.tokens.peek().position
      );
    }

    // Check for range (normal cell range like Sheet1!A1:B2)
    if (this.tokens.match("COLON")) {
      this.tokens.consume();
      const endRef = this.parseRangeEnd();
      // For cross-sheet ranges, do NOT prepend the sheet name to the end reference
      // parseRange will handle sheet name extraction from the start reference
      return this.parseRange(
        ref,
        endRef,
        startPos,
        this.tokens.peek().position.start
      );
    }

    // Parse as single cell reference
    const cellRef = this.parseCellReferenceString(ref);
    if (cellRef) {
      return cellRef;
    }

    // If it's not a cell reference, check if the remaining tokens represent a named expression
    // At this point, we've consumed the sheet name and !, but have an identifier that's not a cell/range/table
    if (this.tokens.isAtEnd() || !this.tokens.match("IDENTIFIER")) {
      // If we're at the end or the next token isn't an identifier, treat the current identifier as a named expression
      // Extract the identifier after the sheet name from ref (remove "SheetName!" prefix)
      const sheetPrefix = sheetName.includes(" ")
        ? `'${sheetName}'!`
        : `${sheetName}!`;
      const identifier = ref.substring(sheetPrefix.length);

      // Validate it's a valid identifier for a named expression
      if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(identifier)) {
        return createNamedExpressionNode(
          identifier,
          {
            start: startPos,
            end: this.tokens.peek().position.start,
          },
          sheetName
        );
      }
    }

    throw new ParseError(`Invalid cell reference: ${ref}`, {
      start: startPos,
      end: this.tokens.peek().position.start,
    });
  }
}

/**
 * Parse a formula string into an AST
 */
export function parseFormula(formula: string): ASTNode {
  return Parser.parse(formula);
}
