import { test, expect, describe } from "bun:test";
import {
  Lexer,
  TokenStream,
  tokenize,
  type Token,
  type TokenType,
} from "../../../src/parser/lexer";

describe("Lexer - Basic Tokenization", () => {
  test("should tokenize numbers correctly", () => {
    const tokens = tokenize("42");
    expect(tokens[0]).toMatchObject({ type: "NUMBER", value: "42" });

    const decimal = tokenize("3.14159");
    expect(decimal[0]).toMatchObject({ type: "NUMBER", value: "3.14159" });

    const negative = tokenize("-42");
    expect(negative[0]).toMatchObject({ type: "OPERATOR", value: "-" });
    expect(negative[1]).toMatchObject({ type: "NUMBER", value: "42" });

    const scientific = tokenize("1.23E-4");
    expect(scientific[0]).toMatchObject({ type: "NUMBER", value: "1.23E-4" });
  });

  test("should tokenize strings correctly", () => {
    const tokens = tokenize('"hello world"');
    expect(tokens[0]).toMatchObject({ type: "STRING", value: "hello world" });

    const escaped = tokenize('"Say ""Hello"""');
    expect(escaped[0]).toMatchObject({ type: "STRING", value: 'Say "Hello"' });
  });

  test("should tokenize booleans correctly", () => {
    const trueTokens = tokenize("TRUE");
    expect(trueTokens[0]).toMatchObject({ type: "BOOLEAN", value: "TRUE" });

    const falseTokens = tokenize("false");
    expect(falseTokens[0]).toMatchObject({ type: "BOOLEAN", value: "FALSE" });
  });

  test("should tokenize operators correctly", () => {
    const operators = [
      "+",
      "-",
      "*",
      "/",
      "^",
      "&",
      "=",
      "<>",
      "<",
      ">",
      "<=",
      ">=",
      "%",
    ];
    operators.forEach((op) => {
      const tokens = tokenize(op);
      expect(tokens[0]).toMatchObject({ type: "OPERATOR", value: op });
    });
  });

  test("should tokenize parentheses and brackets correctly", () => {
    const tokens = tokenize("(){}");
    expect(tokens[0]).toMatchObject({ type: "LPAREN", value: "(" });
    expect(tokens[1]).toMatchObject({ type: "RPAREN", value: ")" });
    expect(tokens[2]).toMatchObject({ type: "LBRACE", value: "{" });
    expect(tokens[3]).toMatchObject({ type: "RBRACE", value: "}" });
  });

  test("should tokenize separators correctly", () => {
    const tokens = tokenize("A1:B2,C3;D4");
    expect(tokens[0]).toMatchObject({ type: "IDENTIFIER", value: "A1" });
    expect(tokens[1]).toMatchObject({ type: "COLON", value: ":" });
    expect(tokens[2]).toMatchObject({ type: "IDENTIFIER", value: "B2" });
    expect(tokens[3]).toMatchObject({ type: "COMMA", value: "," });
    expect(tokens[4]).toMatchObject({ type: "IDENTIFIER", value: "C3" });
    expect(tokens[5]).toMatchObject({ type: "SEMICOLON", value: ";" });
    expect(tokens[6]).toMatchObject({ type: "IDENTIFIER", value: "D4" });
  });

  test("should tokenize error values correctly", () => {
    const errors = ["#DIV/0!", "#N/A", "#NAME?", "#NUM!", "#REF!", "#VALUE!"];
    errors.forEach((error) => {
      const tokens = tokenize(error);
      expect(tokens[0]).toMatchObject({ type: "ERROR", value: error });
    });
  });

  test("should handle whitespace correctly", () => {
    const tokens = tokenize("  A1   +   B2  ");
    const nonWhitespace = tokens.filter((t) => t.type !== "EOF");
    expect(nonWhitespace).toHaveLength(3);
    expect(nonWhitespace[0]).toMatchObject({ type: "IDENTIFIER", value: "A1" });
    expect(nonWhitespace[1]).toMatchObject({ type: "OPERATOR", value: "+" });
    expect(nonWhitespace[2]).toMatchObject({ type: "IDENTIFIER", value: "B2" });
  });

  test("should add EOF token at the end", () => {
    const tokens = tokenize("42");
    expect(tokens[tokens.length - 1]).toMatchObject({ type: "EOF", value: "" });
  });
});

describe("Lexer - Cell References and Identifiers", () => {
  test("should tokenize cell references correctly", () => {
    const tokens = tokenize("A1");
    expect(tokens[0]).toMatchObject({ type: "IDENTIFIER", value: "A1" });

    const absolute = tokenize("$A$1");
    expect(absolute[0]).toMatchObject({ type: "DOLLAR", value: "$" });
    expect(absolute[1]).toMatchObject({ type: "IDENTIFIER", value: "A" });
    expect(absolute[2]).toMatchObject({ type: "DOLLAR", value: "$" });
    expect(absolute[3]).toMatchObject({ type: "NUMBER", value: "1" });
  });

  test("should tokenize sheet references correctly", () => {
    const tokens = tokenize("'Sheet 1'!A1");
    expect(tokens[0]).toMatchObject({ type: "IDENTIFIER", value: "'Sheet 1'" });
    expect(tokens[1]).toMatchObject({ type: "EXCLAMATION", value: "!" });
    expect(tokens[2]).toMatchObject({ type: "IDENTIFIER", value: "A1" });
  });

  test("should handle sheet names with quotes correctly", () => {
    const tokens = tokenize("'John''s Sheet'!A1");
    expect(tokens[0]).toMatchObject({
      type: "IDENTIFIER",
      value: "'John''s Sheet'",
    });
  });

  test("should tokenize named expressions correctly", () => {
    const tokens = tokenize("TaxRate");
    expect(tokens[0]).toMatchObject({ type: "IDENTIFIER", value: "TaxRate" });
  });
});

describe("Lexer - Function Calls", () => {
  test("should identify function names correctly", () => {
    const tokens = tokenize("SUM(A1:A10)");
    expect(tokens[0]).toMatchObject({ type: "FUNCTION", value: "SUM" });
    expect(tokens[1]).toMatchObject({ type: "LPAREN", value: "(" });
  });

  test("should handle nested function calls", () => {
    const tokens = tokenize("IF(A1>0,SUM(B:B),0)");
    expect(tokens[0]).toMatchObject({ type: "FUNCTION", value: "IF" });
    expect(tokens[6]).toMatchObject({ type: "FUNCTION", value: "SUM" });
  });

  test("should handle functions with spaces before parentheses", () => {
    const tokens = tokenize("SUM  (A1)");
    expect(tokens[0]).toMatchObject({ type: "FUNCTION", value: "SUM" });
    expect(tokens[1]).toMatchObject({ type: "LPAREN", value: "(" });
  });
});

describe("Lexer - Complex Formulas", () => {
  test("should tokenize arithmetic expressions", () => {
    const tokens = tokenize("(A1+B1)*2-C1/3");
    const expected = [
      { type: "LPAREN", value: "(" },
      { type: "IDENTIFIER", value: "A1" },
      { type: "OPERATOR", value: "+" },
      { type: "IDENTIFIER", value: "B1" },
      { type: "RPAREN", value: ")" },
      { type: "OPERATOR", value: "*" },
      { type: "NUMBER", value: "2" },
      { type: "OPERATOR", value: "-" },
      { type: "IDENTIFIER", value: "C1" },
      { type: "OPERATOR", value: "/" },
      { type: "NUMBER", value: "3" },
      { type: "EOF", value: "" },
    ];

    tokens.forEach((token, i) => {
      const exp = expected[i];
      if (exp) {
        expect(token.type).toBe(exp.type as TokenType);
        expect(token.value).toBe(exp.value);
      }
    });
  });

  test("should tokenize array formulas", () => {
    const tokens = tokenize("{1,2,3;4,5,6}");
    const types = tokens.map((t) => t.type);
    expect(types).toEqual([
      "LBRACE",
      "NUMBER",
      "COMMA",
      "NUMBER",
      "COMMA",
      "NUMBER",
      "SEMICOLON",
      "NUMBER",
      "COMMA",
      "NUMBER",
      "COMMA",
      "NUMBER",
      "RBRACE",
      "EOF",
    ]);
  });

  test("should tokenize comparison expressions", () => {
    const tokens = tokenize("A1>=10");
    expect(tokens[0]).toMatchObject({ type: "IDENTIFIER", value: "A1" });
    expect(tokens[1]).toMatchObject({ type: "OPERATOR", value: ">=" });
    expect(tokens[2]).toMatchObject({ type: "NUMBER", value: "10" });
  });

  test("should tokenize string concatenation", () => {
    const tokens = tokenize('"Hello "&"World"');
    expect(tokens[0]).toMatchObject({ type: "STRING", value: "Hello " });
    expect(tokens[1]).toMatchObject({ type: "OPERATOR", value: "&" });
    expect(tokens[2]).toMatchObject({ type: "STRING", value: "World" });
  });
});

describe("TokenStream", () => {
  test("should peek at current token", () => {
    const tokens: Token[] = [
      { type: "NUMBER", value: "42", position: { start: 0, end: 2 } },
      { type: "OPERATOR", value: "+", position: { start: 3, end: 4 } },
      { type: "NUMBER", value: "1", position: { start: 5, end: 6 } },
      { type: "EOF", value: "", position: { start: 6, end: 6 } },
    ];

    const stream = new TokenStream(tokens);
    expect(stream.peek().type).toBe("NUMBER");
    expect(stream.peek().value).toBe("42");
  });

  test("should consume tokens correctly", () => {
    const tokens = tokenize("42 + 1"); // Add spaces to ensure proper tokenization
    const stream = new TokenStream(tokens);

    expect(stream.consume().value).toBe("42");
    expect(stream.consume().value).toBe("+");
    expect(stream.consume().value).toBe("1");
    expect(stream.consume().type).toBe("EOF");

    // Should not advance past EOF
    expect(stream.consume().type).toBe("EOF");
  });

  test("should match token types", () => {
    const tokens = tokenize("SUM(42)");
    const stream = new TokenStream(tokens);

    expect(stream.match("FUNCTION")).toBe(true);
    expect(stream.match("NUMBER")).toBe(false);

    stream.consume(); // Skip SUM
    expect(stream.match("LPAREN")).toBe(true);
  });

  test("should match token values", () => {
    const tokens = tokenize("A1+B2");
    const stream = new TokenStream(tokens);

    expect(stream.matchValue("A1")).toBe(true);
    expect(stream.matchValue("B1")).toBe(false);
  });

  test("should consume tokens conditionally", () => {
    const tokens = tokenize("42+");
    const stream = new TokenStream(tokens);

    const number = stream.consumeIf("NUMBER");
    expect(number).not.toBeNull();
    expect(number?.value).toBe("42");

    const lparen = stream.consumeIf("LPAREN");
    expect(lparen).toBeNull();

    const plus = stream.consumeIf("OPERATOR");
    expect(plus).not.toBeNull();
    expect(plus?.value).toBe("+");
  });

  test("should peek ahead correctly", () => {
    const tokens = tokenize("A1+B2*C3");
    const stream = new TokenStream(tokens);

    expect(stream.peekAhead(0)?.value).toBe("A1");
    expect(stream.peekAhead(1)?.value).toBe("+");
    expect(stream.peekAhead(2)?.value).toBe("B2");
    expect(stream.peekAhead(10)).toBeNull();
    expect(stream.peekAhead(-1)).toBeNull();
  });

  test("should check if at end correctly", () => {
    const tokens = tokenize("42");
    const stream = new TokenStream(tokens);

    expect(stream.isAtEnd()).toBe(false);
    stream.consume(); // 42
    expect(stream.isAtEnd()).toBe(true);
  });

  test("should handle position manipulation", () => {
    const tokens = tokenize("A+B+C");
    const stream = new TokenStream(tokens);

    stream.consume(); // A
    stream.consume(); // +
    const pos = stream.getPosition();

    stream.consume(); // B
    stream.consume(); // +

    stream.setPosition(pos);
    expect(stream.peek().value).toBe("B");
  });
});

describe("Lexer - Edge Cases", () => {
  test("should handle empty input", () => {
    const tokens = tokenize("");
    expect(tokens).toHaveLength(1);
    expect(tokens[0]?.type).toBe("EOF");
  });

  test("should handle malformed numbers", () => {
    const tokens = tokenize("12.34.56");
    expect(tokens[0]).toMatchObject({ type: "NUMBER", value: "12.34" });
    expect(tokens[1]).toMatchObject({ type: "NUMBER", value: ".56" }); // .56 is parsed as a decimal number
  });

  test("should handle unclosed strings", () => {
    const tokens = tokenize('"unclosed');
    expect(tokens[0]).toMatchObject({ type: "STRING", value: "unclosed" });
  });

  test("should handle special characters", () => {
    const tokens = tokenize("A1@B2");
    expect(tokens[0]).toMatchObject({ type: "IDENTIFIER", value: "A1" });
    expect(tokens[1]).toMatchObject({ type: "AT", value: "@" });
    expect(tokens[2]).toMatchObject({ type: "IDENTIFIER", value: "B2" });
  });

  test("should handle percentage operator", () => {
    const tokens = tokenize("50%");
    expect(tokens[0]).toMatchObject({ type: "NUMBER", value: "50" });
    expect(tokens[1]).toMatchObject({ type: "OPERATOR", value: "%" });
  });

  test("should handle the INFINITY token", () => {
    const tokens = tokenize("INFINITY");
    expect(tokens[0]).toMatchObject({ type: "INFINITY", value: "INFINITY" });
  });

  test("should tokenize brackets correctly", () => {
    const tokens = tokenize("[]");
    expect(tokens[0]).toMatchObject({ type: "LBRACKET", value: "[" });
    expect(tokens[1]).toMatchObject({ type: "RBRACKET", value: "]" });
  });

  test("should tokenize @ symbol correctly", () => {
    const tokens = tokenize("@");
    expect(tokens[0]).toMatchObject({ type: "AT", value: "@" });
  });

  test("should tokenize # symbol correctly", () => {
    const tokens = tokenize("#");
    expect(tokens[0]).toMatchObject({ type: "ERROR", value: "#" });
  });

  test("should handle table selectors correctly", () => {
    const tokens = tokenize("#Headers");
    expect(tokens[0]).toMatchObject({ type: "HASH", value: "#" });
    expect(tokens[1]).toMatchObject({ type: "IDENTIFIER", value: "Headers" });
  });
});

describe("Lexer - New Syntax Features", () => {
  describe("INFINITY literal", () => {
    test("should tokenize INFINITY as special token", () => {
      const tokens = tokenize("INFINITY");
      expect(tokens[0]).toMatchObject({ type: "INFINITY", value: "INFINITY" });
    });

    test("should handle INFINITY in expressions", () => {
      const tokens = tokenize("SEQUENCE(INFINITY)");
      expect(tokens[0]).toMatchObject({ type: "FUNCTION", value: "SEQUENCE" });
      expect(tokens[2]).toMatchObject({ type: "INFINITY", value: "INFINITY" });
    });

    test("should be case-insensitive", () => {
      const tokens = tokenize("infinity");
      expect(tokens[0]).toMatchObject({ type: "INFINITY", value: "INFINITY" });
      
      const tokens2 = tokenize("Infinity");
      expect(tokens2[0]).toMatchObject({ type: "INFINITY", value: "INFINITY" });
    });
  });

  describe("Structured references", () => {
    test("should tokenize table references", () => {
      const tokens = tokenize("Table1[Column1]");
      expect(tokens[0]).toMatchObject({ type: "IDENTIFIER", value: "Table1" });
      expect(tokens[1]).toMatchObject({ type: "LBRACKET", value: "[" });
      expect(tokens[2]).toMatchObject({ type: "IDENTIFIER", value: "Column1" });
      expect(tokens[3]).toMatchObject({ type: "RBRACKET", value: "]" });
    });

    test("should tokenize current row references", () => {
      const tokens = tokenize("[@Sales]");
      expect(tokens[0]).toMatchObject({ type: "LBRACKET", value: "[" });
      expect(tokens[1]).toMatchObject({ type: "AT", value: "@" });
      expect(tokens[2]).toMatchObject({ type: "IDENTIFIER", value: "Sales" });
      expect(tokens[3]).toMatchObject({ type: "RBRACKET", value: "]" });
    });

    test("should tokenize table selectors", () => {
      const tokens = tokenize("Table1[[#Headers],[Sales]]");
      const expected = [
        { type: "IDENTIFIER", value: "Table1" },
        { type: "LBRACKET", value: "[" },
        { type: "LBRACKET", value: "[" },
        { type: "HASH", value: "#" },
        { type: "IDENTIFIER", value: "Headers" },
        { type: "RBRACKET", value: "]" },
        { type: "COMMA", value: "," },
        { type: "LBRACKET", value: "[" },
        { type: "IDENTIFIER", value: "Sales" },
        { type: "RBRACKET", value: "]" },
        { type: "RBRACKET", value: "]" },
        { type: "EOF", value: "" },
      ];

      tokens.forEach((token, i) => {
        const exp = expected[i];
        if (exp) {
          expect(token.type).toBe(exp.type as TokenType);
          expect(token.value).toBe(exp.value);
        }
      });
    });
  });

  describe("3D ranges", () => {
    test("should tokenize 3D range references", () => {
      const tokens = tokenize("Sheet1:Sheet3!A1");
      expect(tokens[0]).toMatchObject({ type: "IDENTIFIER", value: "Sheet1" });
      expect(tokens[1]).toMatchObject({ type: "COLON", value: ":" });
      expect(tokens[2]).toMatchObject({ type: "IDENTIFIER", value: "Sheet3" });
      expect(tokens[3]).toMatchObject({ type: "EXCLAMATION", value: "!" });
      expect(tokens[4]).toMatchObject({ type: "IDENTIFIER", value: "A1" });
    });

    test("should tokenize 3D range with quoted sheet names", () => {
      const tokens = tokenize("'Sheet 1':'Sheet 3'!A1");
      expect(tokens[0]).toMatchObject({ type: "IDENTIFIER", value: "'Sheet 1'" });
      expect(tokens[1]).toMatchObject({ type: "COLON", value: ":" });
      expect(tokens[2]).toMatchObject({ type: "IDENTIFIER", value: "'Sheet 3'" });
      expect(tokens[3]).toMatchObject({ type: "EXCLAMATION", value: "!" });
      expect(tokens[4]).toMatchObject({ type: "IDENTIFIER", value: "A1" });
    });
  });

  describe("Complex formulas with new syntax", () => {
    test("should tokenize formula with INFINITY", () => {
      const tokens = tokenize("=REPT(\"*\", SEQUENCE(INFINITY))");
      const types = tokens.map((t) => t.type).filter((t) => t !== "EOF");
      expect(types).toEqual([
        "OPERATOR", // =
        "FUNCTION", // REPT
        "LPAREN",
        "STRING",
        "COMMA",
        "FUNCTION", // SEQUENCE
        "LPAREN",
        "INFINITY",
        "RPAREN",
        "RPAREN"
      ]);
    });

    test("should tokenize formula with table reference", () => {
      const tokens = tokenize("=SUM(Table1[Sales])");
      const types = tokens.map((t) => t.type).filter((t) => t !== "EOF");
      expect(types).toEqual([
        "OPERATOR", // =
        "FUNCTION", // SUM
        "LPAREN",
        "IDENTIFIER", // Table1
        "LBRACKET",
        "IDENTIFIER", // Sales
        "RBRACKET",
        "RPAREN"
      ]);
    });

    test("should tokenize formula with 3D range", () => {
      const tokens = tokenize("=AVERAGE(Sheet1:Sheet5!A1)");
      const types = tokens.map((t) => t.type).filter((t) => t !== "EOF");
      expect(types).toEqual([
        "OPERATOR", // =
        "FUNCTION", // AVERAGE
        "LPAREN",
        "IDENTIFIER", // Sheet1
        "COLON",
        "IDENTIFIER", // Sheet5
        "EXCLAMATION",
        "IDENTIFIER", // A1
        "RPAREN"
      ]);
    });
  });
});
