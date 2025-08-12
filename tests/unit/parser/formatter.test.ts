import { describe, expect, test } from "bun:test";
import { astToString, formatFormula } from "../../../src/parser/formatter";
import { parseFormula } from "../../../src/parser/parser";

describe("Formula Formatter", () => {
  describe("Basic Values", () => {
    test("should format numbers", () => {
      expect(formatFormula("42")).toBe("42");
      expect(formatFormula("3.14")).toBe("3.14");
      expect(formatFormula("-5")).toBe("-5");
    });

    test("should format strings", () => {
      expect(formatFormula('"hello"')).toBe('"hello"');
      expect(formatFormula('"hello world"')).toBe('"hello world"');
      expect(formatFormula('""')).toBe('""');
    });

    test("should format strings with escaped quotes", () => {
      expect(formatFormula('"say ""hello"""')).toBe('"say ""hello"""');
    });

    test("should format booleans", () => {
      expect(formatFormula("TRUE")).toBe("TRUE");
      expect(formatFormula("FALSE")).toBe("FALSE");
      expect(formatFormula("true")).toBe("TRUE");
      expect(formatFormula("false")).toBe("FALSE");
    });

    test("should format errors", () => {
      expect(formatFormula("#DIV/0!")).toBe("#DIV/0!");
      expect(formatFormula("#N/A")).toBe("#N/A");
      expect(formatFormula("#NAME?")).toBe("#NAME?");
      expect(formatFormula("#REF!")).toBe("#REF!");
    });

    test("should format infinity", () => {
      expect(formatFormula("INFINITY")).toBe("INFINITY");
      expect(formatFormula("infinity")).toBe("INFINITY");
    });
  });

  describe("Cell References", () => {
    test("should format simple cell references", () => {
      expect(formatFormula("A1")).toBe("A1");
      expect(formatFormula("B10")).toBe("B10");
      expect(formatFormula("Z99")).toBe("Z99");
      expect(formatFormula("AA1")).toBe("AA1");
    });

    test("should format absolute references", () => {
      expect(formatFormula("$A$1")).toBe("$A$1");
      expect(formatFormula("$A1")).toBe("$A1");
      expect(formatFormula("A$1")).toBe("A$1");
    });

    test("should format sheet references", () => {
      expect(formatFormula("Sheet1!A1")).toBe("Sheet1!A1");
      expect(formatFormula("'My Sheet'!A1")).toBe("'My Sheet'!A1");
      expect(formatFormula("Sheet1!$A$1")).toBe("Sheet1!$A$1");
    });
  });

  describe("Range References", () => {
    test("should format simple ranges", () => {
      expect(formatFormula("A1:B2")).toBe("A1:B2");
      expect(formatFormula("A1:Z99")).toBe("A1:Z99");
    });

    test("should format absolute ranges", () => {
      expect(formatFormula("$A$1:$B$2")).toBe("$A$1:$B$2");
      expect(formatFormula("$A1:B$2")).toBe("$A1:B$2");
    });

    test("should format sheet ranges", () => {
      expect(formatFormula("Sheet1!A1:B2")).toBe("Sheet1!A1:B2");
      expect(formatFormula("'My Sheet'!A1:B2")).toBe("'My Sheet'!A1:B2");
    });

    test("should format infinite column ranges", () => {
      expect(formatFormula("A:A")).toBe("A:A");
      expect(formatFormula("A:Z")).toBe("A:Z");
      expect(formatFormula("$A:$Z")).toBe("$A:$Z");
    });

    test("should format infinite row ranges", () => {
      expect(formatFormula("1:1")).toBe("1:1");
      expect(formatFormula("1:100")).toBe("1:100");
      expect(formatFormula("$1:$100")).toBe("$1:$100");
    });
  });

  describe("Function Calls", () => {
    test("should format simple functions", () => {
      expect(formatFormula("SUM(A1:A10)")).toBe("SUM(A1:A10)");
      expect(formatFormula("MAX(B1,B2,B3)")).toBe("MAX(B1,B2,B3)");
      expect(formatFormula("NOW()")).toBe("NOW()");
    });

    test("should format nested functions", () => {
      expect(formatFormula("SUM(MAX(A1:A5),MIN(B1:B5))")).toBe(
        "SUM(MAX(A1:A5),MIN(B1:B5))"
      );
    });

    test("should format functions with mixed arguments", () => {
      expect(formatFormula('IF(A1>10,"High","Low")')).toBe(
        'IF(A1>10,"High","Low")'
      );
      expect(formatFormula("VLOOKUP(A1,B:D,2,FALSE)")).toBe(
        "VLOOKUP(A1,B:D,2,FALSE)"
      );
    });
  });

  describe("Operators", () => {
    test("should format arithmetic operators", () => {
      expect(formatFormula("A1+B1")).toBe("A1+B1");
      expect(formatFormula("A1-B1")).toBe("A1-B1");
      expect(formatFormula("A1*B1")).toBe("A1*B1");
      expect(formatFormula("A1/B1")).toBe("A1/B1");
      expect(formatFormula("A1^B1")).toBe("A1^B1");
    });

    test("should format comparison operators", () => {
      expect(formatFormula("A1=B1")).toBe("A1=B1");
      expect(formatFormula("A1<>B1")).toBe("A1<>B1");
      expect(formatFormula("A1<B1")).toBe("A1<B1");
      expect(formatFormula("A1>B1")).toBe("A1>B1");
      expect(formatFormula("A1<=B1")).toBe("A1<=B1");
      expect(formatFormula("A1>=B1")).toBe("A1>=B1");
    });

    test("should format string concatenation", () => {
      expect(formatFormula("A1&B1")).toBe("A1&B1");
      expect(formatFormula('"Hello "&"World"')).toBe('"Hello "&"World"');
    });

    test("should format unary operators", () => {
      expect(formatFormula("-A1")).toBe("-A1");
      expect(formatFormula("+A1")).toBe("+A1");
      expect(formatFormula("A1%")).toBe("A1%");
    });
  });

  describe("Complex Expressions", () => {
    test("should format expressions with operator precedence", () => {
      expect(formatFormula("A1+B1*C1")).toBe("A1+B1*C1");
      expect(formatFormula("(A1+B1)*C1")).toBe("(A1+B1)*C1");
    });

    test("should format expressions with parentheses", () => {
      expect(formatFormula("(A1+B1)/(C1-D1)")).toBe("(A1+B1)/(C1-D1)");
    });

    test("should format complex nested expressions", () => {
      expect(formatFormula("SUM(A1:A10)+AVERAGE(B1:B10)*COUNT(C1:C10)")).toBe(
        "SUM(A1:A10)+AVERAGE(B1:B10)*COUNT(C1:C10)"
      );
    });
  });

  describe("Array Literals", () => {
    test("should format single row arrays", () => {
      expect(formatFormula("{1,2,3}")).toBe("{1,2,3}");
      expect(formatFormula('{"a","b","c"}')).toBe('{"a","b","c"}');
    });

    test("should format multi-row arrays", () => {
      expect(formatFormula("{1,2;3,4}")).toBe("{1,2;3,4}");
      expect(formatFormula('{"a","b";"c","d"}')).toBe('{"a","b";"c","d"}');
    });

    test("should format arrays with expressions", () => {
      expect(formatFormula("{A1,B1;C1,D1}")).toBe("{A1,B1;C1,D1}");
      expect(formatFormula("{SUM(A:A),MAX(B:B)}")).toBe("{SUM(A:A),MAX(B:B)}");
    });
  });

  describe("Named Expressions", () => {
    test("should format global named expressions", () => {
      expect(formatFormula("MyName")).toBe("MyName");
      expect(formatFormula("SALES_TAX")).toBe("SALES_TAX");
    });

    test("should format sheet-scoped named expressions", () => {
      // Note: This test may need adjustment based on how sheet-scoped names are implemented
      // For now, testing basic named expression format
      expect(formatFormula("MyName")).toBe("MyName");
    });
  });

  describe("3D References", () => {
    test("should format 3D ranges with cell references", () => {
      expect(formatFormula("Sheet1:Sheet3!A1")).toBe("Sheet1:Sheet3!A1");
      expect(formatFormula("'Jan 2023':'Dec 2023'!B2")).toBe(
        "'Jan 2023':'Dec 2023'!B2"
      );
    });

    test("should format 3D ranges with range references", () => {
      expect(formatFormula("Sheet1:Sheet3!A1:B10")).toBe(
        "Sheet1:Sheet3!A1:B10"
      );
    });
  });

  describe("Structured References (Tables)", () => {
    test("should format simple table column references", () => {
      expect(formatFormula("Table1[Column1]")).toBe("Table1[Column1]");
    });

    test("should format table selectors", () => {
      expect(formatFormula("Table1[#Data]")).toBe("Table1[#Data]");
      expect(formatFormula("Table1[#Headers]")).toBe("Table1[#Headers]");
      expect(formatFormula("Table1[#All]")).toBe("Table1[#All]");
    });

    test("should format current row references", () => {
      expect(formatFormula("[@Column1]")).toBe("[@Column1]");
      expect(formatFormula("Table1[@Column1]")).toBe("Table1[@Column1]");
    });

    test("should format complex table references", () => {
      expect(formatFormula("Table1[[#Headers],[Column1]]")).toBe(
        "Table1[[#Headers],[Column1]]"
      );
    });

    test("should format complex table references with range", () => {
      expect(formatFormula("Table1[[#Headers],[Column1:Column2]]")).toBe(
        "Table1[[#Headers],[Column1:Column2]]"
      );
    });
    test("should format complex table references with range resolving to single column", () => {
      expect(formatFormula("Table1[[#Headers],[Column1:Column1]]")).toBe(
        "Table1[[#Headers],[Column1]]"
      );
    });

    describe("Column Names with Special Characters", () => {
      test("should format table columns with spaces (single brackets)", () => {
        // This was the bug: should NOT use double brackets for single columns with spaces
        expect(formatFormula("Table1[CAR ID]")).toBe("Table1[CAR ID]");
        expect(formatFormula("Table1[Net Sales]")).toBe("Table1[Net Sales]");
        expect(formatFormula("Table1[Order Date]")).toBe("Table1[Order Date]");
      });

      test("should format table columns with dashes (single brackets)", () => {
        expect(formatFormula("Table1[CUSTOMER-ID]")).toBe("Table1[CUSTOMER-ID]");
        expect(formatFormula("Table1[ORDER-ID]")).toBe("Table1[ORDER-ID]");
        expect(formatFormula("Table1[ITEM-CODE]")).toBe("Table1[ITEM-CODE]");
      });

      test("should format table columns with underscores", () => {
        expect(formatFormula("Table1[CUSTOMER_ID]")).toBe("Table1[CUSTOMER_ID]");
        expect(formatFormula("Table1[ORDER_DATE]")).toBe("Table1[ORDER_DATE]");
      });

      test("should format current row references with special characters", () => {
        // Current row with spaces - should remove inner brackets when not needed for dashes
        expect(formatFormula("[@[CAR ID]]")).toBe("[@CAR ID]");
        expect(formatFormula("[@[Net Sales]]")).toBe("[@Net Sales]");
        
        // Current row with dashes - should keep simple format since dashes don't need escaping
        expect(formatFormula("[@[CUSTOMER-ID]]")).toBe("[@CUSTOMER-ID]");
        expect(formatFormula("[@CUSTOMER-ID]")).toBe("[@CUSTOMER-ID]");
      });

      test("should format column ranges with special characters (double brackets)", () => {
        // Column ranges should use double brackets when any column has special chars
        expect(formatFormula("Table1[CAR ID:ORDER ID]")).toBe("Table1[[CAR ID]:[ORDER ID]]");
        expect(formatFormula("Table1[Net Sales:Gross Profit]")).toBe("Table1[[Net Sales]:[Gross Profit]]");
        
        // Mixed: one with spaces, one without - should still use double brackets
        expect(formatFormula("Table1[Column1:CAR ID]")).toBe("Table1[[Column1]:[CAR ID]]");
        expect(formatFormula("Table1[CAR ID:Column2]")).toBe("Table1[[CAR ID]:[Column2]]");
      });

      test("should format current row column ranges with special characters", () => {
        expect(formatFormula("Table1[@[CAR ID]:[ORDER ID]]")).toBe("Table1[@[CAR ID]:[ORDER ID]]");
        expect(formatFormula("Table1[@CAR ID:ORDER ID]")).toBe("Table1[@[CAR ID]:[ORDER ID]]");
      });
    });

    describe("Complex Real-World Formulas", () => {
      test("should format INDEX+MATCH with structured references", () => {
        const formula = "INDEX(Table1[CAR ID], MATCH([@[CUSTOMER-ID]], Table1[CUSTOMER-ID],0))";
        const expected = "INDEX(Table1[CAR ID],MATCH([@CUSTOMER-ID],Table1[CUSTOMER-ID],0))";
        expect(formatFormula(formula)).toBe(expected);
      });

      test("should format VLOOKUP with table references", () => {
        const formula = "VLOOKUP([@Customer Name], CustomerTable[Customer Name:Phone Number], 3, FALSE)";
        const expected = "VLOOKUP([@Customer Name],CustomerTable[[Customer Name]:[Phone Number]],3,FALSE)";
        expect(formatFormula(formula)).toBe(expected);
      });

      test("should format SUM with mixed table syntaxes", () => {
        expect(formatFormula("SUM(Table1[Net Sales]) + SUM([Gross Profit])")).toBe(
          "SUM(Table1[Net Sales])+SUM([[Gross Profit]])"
        );
      });

      test("should format formulas with table selectors and columns", () => {
        expect(formatFormula("SUM(Table1[[#Data],[Revenue:Profit]])")).toBe(
          "SUM(Table1[[#Data],[Revenue:Profit]])"
        );
      });
    });

    describe("Bare Column References", () => {
      test("should format bare column references", () => {
        expect(formatFormula("[Column1]")).toBe("[Column1]");
        expect(formatFormula("[result]")).toBe("[result]");
      });

      test("should format bare column ranges", () => {
        expect(formatFormula("[Column1:Column2]")).toBe("[Column1:Column2]");
        expect(formatFormula("[num:result]")).toBe("[num:result]");
      });

      test("should format bare selectors", () => {
        expect(formatFormula("[#Data]")).toBe("[#Data]");
        expect(formatFormula("[#Headers]")).toBe("[#Headers]");
        expect(formatFormula("[#All]")).toBe("[#All]");
      });

      // Note: Bare column references with spaces like [Net Sales] cannot be tested
      // with formatFormula() because they cannot be parsed as standalone formulas.
      // They only exist within table contexts. The formatting behavior is covered
      // by the other structured reference tests.
    });
  });

  describe("Round-trip Tests", () => {
    const testCases = [
      "A1",
      "$A$1",
      "Sheet1!A1",
      "A1:B2",
      "A:A",
      "1:1",
      "SUM(A1:A10)",
      "A1+B1*C1",
      "(A1+B1)*C1",
      "{1,2;3,4}",
      '"Hello"',
      "TRUE",
      "#N/A",
      "INFINITY",
      "MAX(A1,B1,C1)",
      'IF(A1>0,"Positive","Negative")',
      "Sheet1:Sheet3!A1",
      "Table1[Column1]",
      "[@Column1]",
      "Sheet1!Table1[Column1]",
      "TaxRate",
      "Sheet1!TaxRate",
      "'My Sheet'!TaxRate",
      "Table1[[num]:[result]]",
      "Table1[@[num]:[result]]",
      "[result]",
      "[num:result]",
      "[#Data]",
      "[@num] * 10",
      "SUM([result])",
      // Structured references with special characters
      "Table1[CAR ID]",
      "Table1[CUSTOMER-ID]",
      "Table1[Net Sales]",
      "[@CAR ID]",
      "[@CUSTOMER-ID]",
      "[@[Net Sales]]",
      "Table1[[CAR ID]:[ORDER ID]]",
      "Table1[@[Net Sales]:[Gross Profit]]",
      "INDEX(Table1[CAR ID],MATCH([@CUSTOMER-ID],Table1[CUSTOMER-ID],0))",
      "VLOOKUP([@Customer Name],CustomerTable[[Customer Name]:[Phone Number]],3,FALSE)",
    ];

    test.each(testCases)("should round-trip formula: %s", (formula) => {
      const ast = parseFormula(formula);
      const formatted = astToString(ast);
      const reparsed = parseFormula(formatted);
      const reformatted = astToString(reparsed);

      // The formatted string should be stable (formatting twice should give same result)
      expect(reformatted).toBe(formatted);
    });
  });

  describe("Edge Cases", () => {
    test("should handle empty expressions", () => {
      expect(astToString(parseFormula(""))).toBe("");
    });

    test("should handle whitespace preservation", () => {
      // Note: The formatter normalizes whitespace, so this tests normalized output
      expect(formatFormula("A1 + B1")).toBe("A1+B1");
      expect(formatFormula("SUM( A1 : A10 )")).toBe("SUM(A1:A10)");
    });

    test("should handle case normalization", () => {
      expect(formatFormula("sum(a1:a10)")).toBe("SUM(A1:A10)");
      expect(formatFormula("true")).toBe("TRUE");
      expect(formatFormula("false")).toBe("FALSE");
    });
  });

  describe("Regression Tests", () => {
    test("should not use double brackets for single table columns with spaces (Issue Fix)", () => {
      // This was the specific bug: formatter was generating Table1[[CAR ID]] 
      // instead of Table1[CAR ID] for single columns with spaces
      
      // Test the exact formula that was failing
      const originalFormula = "INDEX(Table1[CAR ID], MATCH([@[CUSTOMER-ID]], Table1[CUSTOMER-ID],0))";
      const formatted = formatFormula(originalFormula);
      
      // Should format correctly without double brackets for single columns
      expect(formatted).toBe("INDEX(Table1[CAR ID],MATCH([@CUSTOMER-ID],Table1[CUSTOMER-ID],0))");
      
      // Most importantly: the formatted result should be parseable
      expect(() => parseFormula(formatted)).not.toThrow();
      
      // And round-trip should be stable
      const reparsed = parseFormula(formatted);
      const reformatted = astToString(reparsed);
      expect(reformatted).toBe(formatted);
    });

    test("should handle various problematic column name patterns", () => {
      const testCases = [
        // Spaces in column names
        { input: "Table1[Car Name]", expected: "Table1[Car Name]" },
        { input: "Table1[Customer Full Name]", expected: "Table1[Customer Full Name]" },
        
        // Dashes in column names
        { input: "Table1[ORDER-ID]", expected: "Table1[ORDER-ID]" },
        { input: "Table1[CUSTOMER-ID]", expected: "Table1[CUSTOMER-ID]" },
        
        // Underscores in column names
        { input: "Table1[customer_id]", expected: "Table1[customer_id]" },
        
        // Mixed patterns
        { input: "Table1[Order-Date Time]", expected: "Table1[Order-Date Time]" },
        
        // Current row references with special characters
        { input: "[@[Customer Name]]", expected: "[@Customer Name]" },
        { input: "[@CUSTOMER-ID]", expected: "[@CUSTOMER-ID]" },
      ];

      testCases.forEach(({ input, expected }) => {
        const formatted = formatFormula(input);
        expect(formatted).toBe(expected);
        
        // Ensure the formatted result is parseable
        expect(() => parseFormula(formatted)).not.toThrow();
      });
    });

    test("should preserve double brackets only when required for ranges", () => {
      const testCases = [
        // Double brackets should be used for column ranges with special chars
        { input: "Table1[Car Name:Order Date]", expected: "Table1[[Car Name]:[Order Date]]" },
        { input: "Table1[Column A:Column B]", expected: "Table1[[Column A]:[Column B]]" },
        
        // Double brackets should be used for selector + column combinations
        { input: "Table1[[#Data],[Car Name]]", expected: "Table1[[#Data],[Car Name]]" },
        
        // Single brackets for single columns, even with spaces
        { input: "Table1[Car Name]", expected: "Table1[Car Name]" },
        { input: "Table1[Column A]", expected: "Table1[Column A]" },
      ];

      testCases.forEach(({ input, expected }) => {
        const formatted = formatFormula(input);
        expect(formatted).toBe(expected);
        expect(() => parseFormula(formatted)).not.toThrow();
      });
    });
  });
});
