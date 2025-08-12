/**
 * Lexer for tokenizing formula strings
 */

import type { FormulaError } from '../core/types';

/**
 * Token types for formula lexing
 */
export type TokenType =
  | 'NUMBER'
  | 'STRING'
  | 'BOOLEAN'
  | 'IDENTIFIER'
  | 'FUNCTION'
  | 'OPERATOR'
  | 'LPAREN'
  | 'RPAREN'
  | 'LBRACE'
  | 'RBRACE'
  | 'LBRACKET'
  | 'RBRACKET'
  | 'COMMA'
  | 'SEMICOLON'
  | 'COLON'
  | 'DOLLAR'
  | 'EXCLAMATION'
  | 'AT'
  | 'HASH'
  | 'INFINITY'
  | 'ERROR'
  | 'EOF'
  | 'WHITESPACE';

/**
 * Token interface
 */
export interface Token {
  type: TokenType;
  value: string;
  position: {
    start: number;
    end: number;
  };
}

/**
 * Lexer class for tokenizing formula strings
 */
export class Lexer {
  private input: string;
  private position: number;
  private tokens: Token[];
  
  constructor(input: string) {
    this.input = input;
    this.position = 0;
    this.tokens = [];
  }
  
  /**
   * Tokenize the entire input string
   */
  tokenize(): Token[] {
    this.tokens = [];
    this.position = 0;
    
    while (this.position < this.input.length) {
      const token = this.nextToken();
      if (token && token.type !== 'WHITESPACE') {
        this.tokens.push(token);
      }
    }
    
    // Add EOF token
    this.tokens.push({
      type: 'EOF',
      value: '',
      position: { start: this.position, end: this.position }
    });
    
    return this.tokens;
  }
  
  /**
   * Get the next token from the input
   */
  private nextToken(): Token | null {
    const start = this.position;
    
    // Check for single-character tokens
    const char = this.input[this.position];
    if (char === undefined) {
      return null;
    }
    
    // Skip whitespace
    if (this.isWhitespace()) {
      return this.readWhitespace();
    }
    
    switch (char) {
      case '(':
        this.position++;
        return { type: 'LPAREN', value: char, position: { start, end: this.position } };
      case ')':
        this.position++;
        return { type: 'RPAREN', value: char, position: { start, end: this.position } };
      case '{':
        this.position++;
        return { type: 'LBRACE', value: char, position: { start, end: this.position } };
      case '}':
        this.position++;
        return { type: 'RBRACE', value: char, position: { start, end: this.position } };
      case '[':
        this.position++;
        return { type: 'LBRACKET', value: char, position: { start, end: this.position } };
      case ']':
        this.position++;
        return { type: 'RBRACKET', value: char, position: { start, end: this.position } };
      case ',':
        this.position++;
        return { type: 'COMMA', value: char, position: { start, end: this.position } };
      case ';':
        this.position++;
        return { type: 'SEMICOLON', value: char, position: { start, end: this.position } };
      case ':':
        this.position++;
        return { type: 'COLON', value: char, position: { start, end: this.position } };
      case '$':
        this.position++;
        return { type: 'DOLLAR', value: char, position: { start, end: this.position } };
      case '!':
        this.position++;
        return { type: 'EXCLAMATION', value: char, position: { start, end: this.position } };
      case '@':
        this.position++;
        return { type: 'AT', value: char, position: { start, end: this.position } };
      case '"':
        return this.readString();
      case "'":
        return this.readSheetName();
      case '#':
        // Check if it's a table selector or an error
        if (this.position + 1 < this.input.length) {
          const nextChar = this.input[this.position + 1];
          if (nextChar && this.isAlpha(nextChar)) {
            // Might be a table selector (#Headers, #Data, etc.) or error
            const lookahead = this.peekSelector();
            if (lookahead) {
              this.position++;
              return { type: 'HASH', value: char, position: { start, end: this.position } };
            }
          }
        }
        return this.readError();
      case '+':
      case '-':
        // Always treat as operator - parser will handle unary operators
        this.position++;
        return { type: 'OPERATOR', value: char, position: { start, end: this.position } };
      case '*':
      case '/':
      case '^':
      case '&':
      case '%':
        this.position++;
        return { type: 'OPERATOR', value: char, position: { start, end: this.position } };
      case '=':
        this.position++;
        return { type: 'OPERATOR', value: char, position: { start, end: this.position } };
      case '<':
        if (this.position + 1 < this.input.length && this.input[this.position + 1] === '>') {
          this.position += 2;
          return { type: 'OPERATOR', value: '<>', position: { start, end: this.position } };
        }
        if (this.position + 1 < this.input.length && this.input[this.position + 1] === '=') {
          this.position += 2;
          return { type: 'OPERATOR', value: '<=', position: { start, end: this.position } };
        }
        this.position++;
        return { type: 'OPERATOR', value: char, position: { start, end: this.position } };
      case '>':
        if (this.position + 1 < this.input.length && this.input[this.position + 1] === '=') {
          this.position += 2;
          return { type: 'OPERATOR', value: '>=', position: { start, end: this.position } };
        }
        this.position++;
        return { type: 'OPERATOR', value: char, position: { start, end: this.position } };
    }
    
    // Check for numbers
    if (char && this.isDigit(char)) {
      return this.readNumber();
    }
    
    // Check for decimal numbers starting with .
    if (char === '.') {
      const nextChar = this.input[this.position + 1];
      if (this.position + 1 < this.input.length && nextChar !== undefined && this.isDigit(nextChar)) {
        return this.readNumber();
      }
      // Otherwise it's an error
      this.position++;
      return { type: 'ERROR', value: char, position: { start, end: this.position } };
    }
    
    // Check for identifiers (cell references, function names, boolean values)
    if (char && (this.isAlpha(char) || char === '_')) {
      return this.readIdentifier();
    }
    
    // Unknown character
    this.position++;
    return { type: 'ERROR', value: char, position: { start, end: this.position } };
  }
  
  /**
   * Read a whitespace token
   */
  private readWhitespace(): Token {
    const start = this.position;
    while (this.position < this.input.length && this.isWhitespace()) {
      this.position++;
    }
    return {
      type: 'WHITESPACE',
      value: this.input.substring(start, this.position),
      position: { start, end: this.position }
    };
  }
  
  /**
   * Read a number token
   */
  private readNumber(): Token {
    const start = this.position;
    let hasDecimal = false;
    let hasExponent = false;
    
    // Check for sign
    let currentChar = this.input[this.position];
    if (currentChar === '+' || currentChar === '-') {
      this.position++;
    }
    
    // Read integer part
    while (this.position < this.input.length) {
      currentChar = this.input[this.position];
      if (currentChar && this.isDigit(currentChar)) {
        this.position++;
      } else {
        break;
      }
    }
    
    // Read decimal part
    if (this.position < this.input.length) {
      currentChar = this.input[this.position];
      if (currentChar === '.') {
        hasDecimal = true;
        this.position++;
        while (this.position < this.input.length) {
          currentChar = this.input[this.position];
          if (currentChar && this.isDigit(currentChar)) {
            this.position++;
          } else {
            break;
          }
        }
      }
    }
    
    // Read exponent part
    if (this.position < this.input.length) {
      currentChar = this.input[this.position];
      if (currentChar === 'E' || currentChar === 'e') {
        hasExponent = true;
        this.position++;
        
        // Optional sign
        if (this.position < this.input.length) {
          currentChar = this.input[this.position];
          if (currentChar === '+' || currentChar === '-') {
            this.position++;
          }
        }
        
        // Exponent digits
        while (this.position < this.input.length) {
          currentChar = this.input[this.position];
          if (currentChar && this.isDigit(currentChar)) {
            this.position++;
          } else {
            break;
          }
        }
      }
    }
    
    const value = this.input.substring(start, this.position);
    return { type: 'NUMBER', value, position: { start, end: this.position } };
  }
  
  /**
   * Read a string token (enclosed in double quotes)
   */
  private readString(): Token {
    const start = this.position;
    this.position++; // Skip opening quote
    
    let value = '';
    while (this.position < this.input.length) {
      const char = this.input[this.position];
      
      if (char === '"') {
        // Check for escaped quote
        if (this.position + 1 < this.input.length && this.input[this.position + 1] === '"') {
          value += '"';
          this.position += 2;
        } else {
          // End of string
          this.position++;
          break;
        }
      } else {
        value += char;
        this.position++;
      }
    }
    
    return { type: 'STRING', value, position: { start, end: this.position } };
  }
  
  /**
   * Read a sheet name (enclosed in single quotes)
   */
  private readSheetName(): Token {
    const start = this.position;
    this.position++; // Skip opening quote
    
    let value = "'";
    while (this.position < this.input.length) {
      const char = this.input[this.position];
      
      if (char === "'") {
        // Check for escaped quote
        if (this.position + 1 < this.input.length && this.input[this.position + 1] === "'") {
          value += "''";
          this.position += 2;
        } else {
          // End of sheet name
          value += char;
          this.position++;
          break;
        }
      } else {
        value += char;
        this.position++;
      }
    }
    
    // This is actually part of an identifier (sheet reference)
    return { type: 'IDENTIFIER', value: value, position: { start, end: this.position } };
  }
  
  /**
   * Read an error token (e.g., #DIV/0!)
   */
  private readError(): Token {
    const start = this.position;
    this.position++; // Skip #
    
    // Read until we hit a non-alphanumeric character or special error characters
    while (this.position < this.input.length) {
      const char = this.input[this.position];
      if (char && (this.isAlnum(char) || char === '/' || char === '!' || char === '?')) {
        this.position++;
      } else {
        break;
      }
    }
    
    const value = this.input.substring(start, this.position);
    return { type: 'ERROR', value, position: { start, end: this.position } };
  }
  
  /**
   * Read an identifier (cell reference, function name, boolean)
   */
  private readIdentifier(): Token {
    const start = this.position;
    
    // Read identifier characters
    while (this.position < this.input.length) {
      const char = this.input[this.position];
      if (char && (this.isAlnum(char) || char === '_' || char === '.')) {
        this.position++;
      } else {
        break;
      }
    }
    
    const value = this.input.substring(start, this.position);
    
    // Check if it's followed by an opening parenthesis (function call)
    let lookahead = this.position;
    while (lookahead < this.input.length) {
      const lookChar = this.input[lookahead];
      if (lookChar && this.isWhitespaceChar(lookChar)) {
        lookahead++;
      } else {
        break;
      }
    }
    
    if (lookahead < this.input.length) {
      const lookChar = this.input[lookahead];
      if (lookChar === '(') {
        return { type: 'FUNCTION', value: value.toUpperCase(), position: { start, end: this.position } };
      }
    }
    
    // Check if it's a boolean (only if not a function)
    if (value.toUpperCase() === 'TRUE' || value.toUpperCase() === 'FALSE') {
      return { type: 'BOOLEAN', value: value.toUpperCase(), position: { start, end: this.position } };
    }
    
    // Check if it's INFINITY
    if (value.toUpperCase() === 'INFINITY') {
      return { type: 'INFINITY', value: value.toUpperCase(), position: { start, end: this.position } };
    }
    
    return { type: 'IDENTIFIER', value, position: { start, end: this.position } };
  }
  
  /**
   * Check if current position is whitespace
   */
  private isWhitespace(): boolean {
    const char = this.input[this.position];
    return this.position < this.input.length && char !== undefined && this.isWhitespaceChar(char);
  }
  
  /**
   * Check if a character is whitespace
   */
  private isWhitespaceChar(char: string): boolean {
    return char === ' ' || char === '\t' || char === '\n' || char === '\r';
  }
  
  /**
   * Check if a character is a digit
   */
  private isDigit(char: string): boolean {
    return char >= '0' && char <= '9';
  }
  
  /**
   * Check if a character is alphabetic
   */
  private isAlpha(char: string): boolean {
    return (char >= 'A' && char <= 'Z') || (char >= 'a' && char <= 'z');
  }
  
  /**
   * Check if a character is alphanumeric
   */
  private isAlnum(char: string): boolean {
    return this.isAlpha(char) || this.isDigit(char);
  }
  
  /**
   * Peek ahead to check if we have a table selector
   */
  private peekSelector(): boolean {
    const selectors = ['All', 'Data', 'Headers', 'ThisRow'];
    const currentPos = this.position + 1; // Skip the #
    
    for (const selector of selectors) {
      if (this.input.substring(currentPos, currentPos + selector.length).toUpperCase() === selector.toUpperCase()) {
        // Check that the next character is not alphanumeric (to avoid matching #DataSomething)
        const nextPos = currentPos + selector.length;
        if (nextPos >= this.input.length || !this.isAlnum(this.input[nextPos]!)) {
          return true;
        }
      }
    }
    
    return false;
  }
}

/**
 * Quick tokenization function
 */
export function tokenize(input: string): Token[] {
  const lexer = new Lexer(input);
  return lexer.tokenize();
}

/**
 * Token stream for parser consumption
 */
export class TokenStream {
  private tokens: Token[];
  private position: number;
  
  constructor(tokens: Token[]) {
    this.tokens = tokens;
    this.position = 0;
  }
  
  /**
   * Peek at current token without consuming
   */
  peek(): Token {
    if (this.position >= this.tokens.length || this.tokens.length === 0) {
      // Return a default EOF token if no tokens
      return { type: 'EOF', value: '', position: { start: 0, end: 0 } };
    }
    return this.tokens[this.position]!;
  }
  
  /**
   * Peek ahead by n tokens
   */
  peekAhead(n: number): Token | null {
    const pos = this.position + n;
    if (pos >= this.tokens.length || pos < 0) {
      return null;
    }
    return this.tokens[pos] ?? null;
  }
  
  /**
   * Peek at the next token without consuming
   */
  peekNext(): Token | null {
    const pos = this.position + 1;
    if (pos >= this.tokens.length) {
      return null;
    }
    return this.tokens[pos] ?? null;
  }
  
  /**
   * Consume current token and advance
   */
  consume(): Token {
    const token = this.peek();
    if (this.position < this.tokens.length - 1) {
      this.position++;
    }
    return token;
  }
  
  /**
   * Check if current token matches type
   */
  match(type: TokenType): boolean {
    return this.peek().type === type;
  }
  
  /**
   * Check if current token matches value
   */
  matchValue(value: string): boolean {
    return this.peek().value === value;
  }
  
  /**
   * Consume token if it matches type
   */
  consumeIf(type: TokenType): Token | null {
    if (this.match(type)) {
      return this.consume();
    }
    return null;
  }
  
  /**
   * Check if at end of stream
   */
  isAtEnd(): boolean {
    return this.peek().type === 'EOF';
  }
  
  /**
   * Get current position in stream
   */
  getPosition(): number {
    return this.position;
  }
  
  /**
   * Set position in stream
   */
  setPosition(position: number): void {
    this.position = Math.max(0, Math.min(position, this.tokens.length - 1));
  }
  
  /**
   * Get all tokens
   */
  getTokens(): Token[] {
    return this.tokens;
  }
}
