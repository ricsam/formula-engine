import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("TEXT function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const address = (ref: string) => ({
    sheetName,
    workbookName,
    ...parseCellReference(ref),
  });

  const cell = (ref: string) => engine.getCellValue(address(ref));

  const errorMessage = (ref: string) =>
    engine.getCellErrorDetails(address(ref))?.message;

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(address(ref), content);
  };

  const setContent = (content: Array<[string, SerializedCellValue]>) => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>(content)
    );
  };

  /** Puts `=TEXT(value, "format")` in A1, A2, ... and returns the results. */
  const formatAll = (formulas: string[]) => {
    setContent(formulas.map((formula, index) => [`A${index + 1}`, formula]));
    return formulas.map((_, index) => cell(`A${index + 1}`));
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("numbers", () => {
    test("formats the examples of Excel's documentation", () => {
      expect(
        formatAll([
          '=TEXT(1234.567,"$#,##0.00")',
          '=TEXT(0.285,"0.0%")',
          '=TEXT(4.34,"# ?/?")',
          '=TEXT(0.34,"# ?/?")',
          '=TEXT(12200000,"0.00E+00")',
          '=TEXT(1234567898,"[<=9999999]###-####;(###) ###-####")',
          '=TEXT(1234,"0000000")',
          "=TEXT(123456,\"##0° 00' 00''\")",
        ])
      ).toEqual([
        "$1,234.57",
        "28.5%",
        "4 1/3",
        " 1/3",
        "1.22E+07",
        "(123) 456-7898",
        "0001234",
        "12° 34' 56''",
      ]);
    });

    test("rounds the 15-digit decimal, halves away from zero", () => {
      expect(
        formatAll([
          '=TEXT(2.675,"0.00")',
          '=TEXT(1.005,"0.00")',
          '=TEXT(0.5,"0")',
          '=TEXT(-0.5,"0")',
          '=TEXT(2.5,"0")',
          '=TEXT(-2.5,"0")',
          '=TEXT(0.125,"0.00")',
          '=TEXT(0.1+0.2,"0.0000000000000000")',
          '=TEXT(1/3,"0.00000000000000000000")',
          '=TEXT(123456789012345678,"0")',
          '=TEXT(1234567.891,"#,##0")',
        ])
      ).toEqual([
        "2.68",
        "1.01",
        "1",
        "-1",
        "3",
        "-3",
        "0.13",
        "0.3000000000000000",
        "0.33333333333333300000",
        "123456789012346000",
        "1,234,568",
      ]);
    });

    test("shows an exact tie in the 16th digit as the 15-digit decimal nearer zero", () => {
      expect(
        formatAll([
          '=TEXT(1000000000000005,"0")',
          '=TEXT(123456789012345.5,"0.0")',
          '=TEXT(99999999999999.25,"0.0")',
          '=TEXT(-1000000000000005,"0")',
        ])
      ).toEqual([
        "1000000000000000",
        "123456789012345.0",
        "99999999999999.2",
        "-1000000000000000",
      ]);
    });

    test("never shows a minus sign without a non-zero digit", () => {
      setContent([
        ["B1", -0],
        ["A1", '=TEXT(-0.4,"0")'],
        ["A2", '=TEXT(-0.04,"0.0")'],
        ["A3", '=TEXT(-0.05,"0.0")'],
        ["A4", '=TEXT(B1,"0.00")'],
        ["A5", '=TEXT(-1E-20,"#,##0.00")'],
        // An explicit negative section is used for every negative number
        ["A6", '=TEXT(-0.4,"0;-0")'],
      ]);

      expect(cell("A1")).toBe("0");
      expect(cell("A2")).toBe("0.0");
      expect(cell("A3")).toBe("-0.1");
      expect(cell("A4")).toBe("0.00");
      expect(cell("A5")).toBe("0.00");
      expect(cell("A6")).toBe("-0");
    });

    test("formats digit placeholders, decimals and padding", () => {
      expect(
        formatAll([
          '=TEXT(5,"000")',
          '=TEXT(1.5,"???.??")',
          '=TEXT(1.5,"0.0#")',
          '=TEXT(1.25,"0.0#")',
          '=TEXT(0,"#")',
          '=TEXT(0,"#.##")',
          '=TEXT(0.5,".00")',
          '=TEXT(1234,".0")',
        ])
      ).toEqual(["005", "  1.5 ", "1.5", "1.25", "", ".", ".50", "1234.0"]);
    });

    test("groups thousands and scales by commas and percent signs", () => {
      expect(
        formatAll([
          '=TEXT(1234567.891,"#,##0.00")',
          '=TEXT(1234567,"#,##0,")',
          '=TEXT(1234567,"0.0,,")',
          '=TEXT(12345,"#,##0.0,"" K""")',
          '=TEXT(0.256,"0%")',
          '=TEXT(0.12345,"0.00%")',
          '=TEXT(0.5,"0%%")',
        ])
      ).toEqual([
        "1,234,567.89",
        "1,235",
        "1.2",
        "12.3 K",
        "26%",
        "12.35%",
        "5000%%",
      ]);
    });

    test("fills digit groups between literal text", () => {
      expect(
        formatAll([
          '=TEXT(123456789,"000-00-0000")',
          '=TEXT(12345,"000-00-0000")',
          '=TEXT(5551234567,"(000) 000-0000")',
          '=TEXT(5,"""Well ""00")',
          '=TEXT(12,"0"" µL""")',
          '=TEXT(941051630,"00000\\-0000")',
        ])
      ).toEqual([
        "123-45-6789",
        "000-01-2345",
        "(555) 123-4567",
        "Well 05",
        "12 µL",
        "94105-1630",
      ]);
    });

    test("shows literal characters, spaces of _x and nothing for *x", () => {
      expect(
        formatAll([
          '=TEXT(5,"$-+():!^&\'~{}<>= 0")',
          '=TEXT(-5,"$-+():!^&\'~{}<>= 0")',
          '=TEXT(5,"\\$0.00")',
          '=TEXT(5,"0_)")',
          '=TEXT(-5,"0_);(0)")',
          '=TEXT(5,"*-0")',
          '=TEXT(1234.5,"_($* #,##0.00_)")',
        ])
      ).toEqual([
        "$-+():!^&'~{}<>= 5",
        "-$-+():!^&'~{}<>= 5",
        "$5.00",
        "5 ",
        "(5)",
        "5",
        // The fill `* ` of the accounting format shows nothing
        " $1,234.50 ",
      ]);
    });

    test("formats scientific notation", () => {
      expect(
        formatAll([
          '=TEXT(12345,"0.00E+00")',
          '=TEXT(0.000123,"0.00E+00")',
          '=TEXT(12345,"0.00E-00")',
          '=TEXT(-12345,"0.0E+0")',
          '=TEXT(0,"0.00E+00")',
          '=TEXT(9.995,"0.00E+00")',
          '=TEXT(12345,"##0.0E+0")',
          '=TEXT(12345,"0.0e+0")',
        ])
      ).toEqual([
        "1.23E+04",
        "1.23E-04",
        "1.23E04",
        "-1.2E+4",
        "0.00E+00",
        "1.00E+01",
        "12.3E+3",
        "1.2e+4",
      ]);
    });

    test("formats fractions", () => {
      expect(
        formatAll([
          '=TEXT(4.34,"# ?/?")',
          '=TEXT(-4.34,"# ?/?")',
          '=TEXT(0.3,"# ?/?")',
          '=TEXT(5,"# ?/?")',
          '=TEXT(4.96,"# ?/?")',
          '=TEXT(12.3456789,"# ??/??")',
          '=TEXT(12.3456789,"??/??")',
          '=TEXT(0.3,"# ?/8")',
          '=TEXT(3.14159,"# ??/100")',
        ])
      ).toEqual([
        "4 1/3",
        "-4 1/3",
        " 2/7",
        "5    ",
        "5    ",
        "12 28/81",
        "1000/81",
        " 2/8",
        "3 14/100",
      ]);
    });

    test("formats General like Excel, within 11 characters", () => {
      expect(
        formatAll([
          '=TEXT(1/3,"General")',
          '=TEXT(-1/3,"General")',
          '=TEXT(123456789012,"General")',
          '=TEXT(12345678901,"General")',
          '=TEXT(1234567890.12,"General")',
          '=TEXT(0.0001,"General")',
          '=TEXT(0.00001234567,"General")',
          '=TEXT(0.1+0.2,"general")',
          '=TEXT(1E+21,"GENERAL")',
          '=TEXT(5,"""Total: ""General")',
        ])
      ).toEqual([
        "0.333333333",
        "-0.333333333",
        "1.23457E+11",
        "12345678901",
        "1234567890",
        "0.0001",
        "1.23457E-05",
        "0.3",
        "1E+21",
        "Total: 5",
      ]);
    });
  });

  describe("sections and conditions", () => {
    test("uses the positive, negative and zero sections", () => {
      setContent([
        ["A1", '=TEXT(5,"0;(0);""zero""")'],
        ["A2", '=TEXT(-5,"0;(0);""zero""")'],
        ["A3", '=TEXT(0,"0;(0);""zero""")'],
        // Rounding to zero does not select the zero section
        ["A4", '=TEXT(0.001,"0;(0);""zero""")'],
        // Two sections: zero uses the first
        ["A5", '=TEXT(0,"0.0;(0.0)")'],
        // One section: negative numbers get a minus sign
        ["A6", '=TEXT(-5,"0.0")'],
        // An empty section shows nothing
        ["A7", '=TEXT(-5,"0;;0")'],
        ["A8", '=TEXT(5,";;;")'],
      ]);

      expect(cell("A1")).toBe("5");
      expect(cell("A2")).toBe("(5)");
      expect(cell("A3")).toBe("zero");
      expect(cell("A4")).toBe("0");
      expect(cell("A5")).toBe("0.0");
      expect(cell("A6")).toBe("-5.0");
      expect(cell("A7")).toBe("");
      expect(cell("A8")).toBe("");
    });

    test("picks the section whose condition matches", () => {
      setContent([
        ["A1", '=TEXT(5,"[>=100]0;[<100]0.0")'],
        ["A2", '=TEXT(500,"[>=100]0;[<100]0.0")'],
        ["A3", '=TEXT(-5,"[>=100]0;[<100]0.0")'],
        ["A4", '=TEXT(-10,"[<=-5]""low"";0")'],
        // As in Excel, the section after [<=-5] shows no minus sign
        ["A5", '=TEXT(-1,"[<=-5]""low"";0")'],
        ["A6", '=TEXT(2,"[=1]""one"";[=2]""two"";""other""")'],
        ["A7", '=TEXT(3,"[=1]""one"";[=2]""two"";""other""")'],
        ["A8", '=TEXT(-3.96,"[>=100]#,##0;[<=-100]-#,##0;#,##0.00")'],
        ["A9", '=TEXT(0,"[>0]""pos"";[<0]""neg""")'],
      ]);

      expect(cell("A1")).toBe("5.0");
      expect(cell("A2")).toBe("500");
      expect(cell("A3")).toBe("-5.0");
      expect(cell("A4")).toBe("low");
      expect(cell("A5")).toBe("1");
      expect(cell("A6")).toBe("two");
      expect(cell("A7")).toBe("other");
      expect(cell("A8")).toBe("-3.96");
      // No section applies to 0
      expect(cell("A9")).toBe(FormulaError.VALUE);
      expect(errorMessage("A9")).toContain(
        "no section of the format applies to 0"
      );
    });

    test("ignores colors and locale tags but shows currency symbols", () => {
      expect(
        formatAll([
          '=TEXT(-1.5,"[Red]0.00;[Blue]-0.00")',
          '=TEXT(5,"[Color10]0")',
          '=TEXT(5,"[$-409]0")',
          '=TEXT(5,"[$€-407]0.00")',
          '=TEXT(1234.5,"[$$-409]#,##0.00")',
          '=TEXT(5,"[Red]")',
        ])
      ).toEqual(["-1.50", "5", "5", "€5.00", "$1,234.50", "5"]);
    });
  });

  describe("dates and times", () => {
    test("formats serial numbers as dates", () => {
      expect(
        formatAll([
          '=TEXT(45000,"yyyy-mm-dd")',
          '=TEXT(45000,"dddd, mmmm d, yyyy")',
          '=TEXT(45000,"ddd d mmm yy")',
          '=TEXT(45000,"m/d/yyyy")',
          '=TEXT(45000,"mmmmm")',
          '=TEXT(45000,"DD.MM.YYYY")',
          '=TEXT(45000,"[$-F800]dddd, mmmm dd, yyyy")',
        ])
      ).toEqual([
        "2023-03-15",
        "Wednesday, March 15, 2023",
        "Wed 15 Mar 23",
        "3/15/2023",
        "M",
        "15.03.2023",
        "Wednesday, March 15, 2023",
      ]);
    });

    test("follows Excel's 1900 date system, including 1900-02-29", () => {
      expect(
        formatAll([
          '=TEXT(0,"yyyy-mm-dd dddd")',
          '=TEXT(1,"yyyy-mm-dd dddd")',
          '=TEXT(59,"yyyy-mm-dd")',
          '=TEXT(60,"yyyy-mm-dd dddd")',
          '=TEXT(61,"yyyy-mm-dd dddd")',
          '=TEXT(2958465,"yyyy-mm-dd")',
        ])
      ).toEqual([
        "1900-01-00 Saturday",
        "1900-01-01 Sunday",
        "1900-02-28",
        "1900-02-29 Wednesday",
        "1900-03-01 Thursday",
        "9999-12-31",
      ]);
    });

    test("formats times, with m as minutes after hours and before seconds", () => {
      expect(
        formatAll([
          '=TEXT(0.5,"h:mm AM/PM")',
          '=TEXT(0,"h:mm AM/PM")',
          '=TEXT(0.75,"hh:mm:ss")',
          '=TEXT(0.75,"h:mm am/pm")',
          '=TEXT(0.25,"h A/P")',
          '=TEXT(0.04,"mm:ss")',
          '=TEXT(0.04,"mm")',
          '=TEXT(45000.75,"yyyy-mm-dd hh:mm")',
        ])
      ).toEqual([
        "12:00 PM",
        "12:00 AM",
        "18:00:00",
        "6:00 PM",
        "6 A",
        "57:36",
        "01",
        "2023-03-15 18:00",
      ]);
    });

    test("rounds to the seconds shown, never to minutes", () => {
      setContent([
        // 10:30:59
        ["B1", "=(10*3600+30*60+59)/86400"],
        // 10:30:59.6
        ["B2", "=(10*3600+30*60+59.6)/86400"],
        ["A1", '=TEXT(B1,"h:mm")'],
        ["A2", '=TEXT(B2,"h:mm:ss")'],
        ["A3", '=TEXT(B2,"h:mm:ss.0")'],
        ["A4", '=TEXT(0.123456,"h:mm:ss.00")'],
        // Rounding up to midnight moves to the next day
        ["A5", '=TEXT(45000.9999999,"yyyy-mm-dd hh:mm")'],
        ["A6", '=TEXT(45000.9999999,"yyyy-mm-dd")'],
      ]);

      expect(cell("A1")).toBe("10:30");
      expect(cell("A2")).toBe("10:31:00");
      expect(cell("A3")).toBe("10:30:59.6");
      expect(cell("A4")).toBe("2:57:46.60");
      expect(cell("A5")).toBe("2023-03-16 00:00");
      expect(cell("A6")).toBe("2023-03-15");
    });

    test("formats elapsed time", () => {
      expect(
        formatAll([
          '=TEXT(1.5,"[h]:mm")',
          '=TEXT(1.5,"[hh]:mm:ss")',
          '=TEXT(0.04,"[mm]:ss")',
          '=TEXT(0.5,"[s]")',
          '=TEXT(1,"[hhh]")',
        ])
      ).toEqual(["36:00", "36:00:00", "57:36", "43200", "024"]);
    });

    test("rejects serial numbers outside Excel's dates", () => {
      setContent([
        ["A1", '=TEXT(-1,"yyyy-mm-dd")'],
        ["A2", '=TEXT(2958466,"yyyy-mm-dd")'],
        ["A3", '=TEXT(-0.5,"[h]:mm")'],
        ["A4", '=TEXT(-0.5,"h:mm")'],
        // A number section of the same format still formats negative numbers
        ["A5", '=TEXT(-1,"yyyy-mm-dd;0")'],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(errorMessage("A1")).toContain(
        "is not a date or time that can be shown"
      );
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(cell("A3")).toBe(FormulaError.VALUE);
      expect(cell("A4")).toBe(FormulaError.VALUE);
      expect(cell("A5")).toBe("1");
    });
  });

  describe("value types", () => {
    test("returns text unchanged unless the format has a text section", () => {
      setContent([
        ["B1", "abc"],
        ["B2", "1.5"],
        ["A1", '=TEXT(B1,"0.00")'],
        // Numeric-looking text stays text
        ["A2", '=TEXT(B2,"0.00")'],
        ["A3", '=TEXT(B1,"0;0;0;""[""@""]""")'],
        ["A4", '=TEXT(B1,"@"" units""")'],
        ["A5", '=TEXT(B1,"0;""t""@")'],
        ["A6", '=TEXT(B1,";;;")'],
        // A format of only a text section formats numbers as General
        ["A7", '=TEXT(1.5,"@"" units""")'],
      ]);

      expect(cell("A1")).toBe("abc");
      expect(cell("A2")).toBe("1.5");
      expect(cell("A3")).toBe("[abc]");
      expect(cell("A4")).toBe("abc units");
      expect(cell("A5")).toBe("tabc");
      expect(cell("A6")).toBe("");
      expect(cell("A7")).toBe("1.5");
    });

    test("formats logical values as the text TRUE or FALSE", () => {
      setContent([
        ["A1", '=TEXT(TRUE,"0.00")'],
        ["A2", '=TEXT(FALSE,"0")'],
        ["A3", '=TEXT(TRUE,"0;0;0;""<""@"">""")'],
        ["A4", '=TEXT(1=1,"yyyy-mm-dd")'],
      ]);

      expect(cell("A1")).toBe("TRUE");
      expect(cell("A2")).toBe("FALSE");
      expect(cell("A3")).toBe("<TRUE>");
      expect(cell("A4")).toBe("TRUE");
    });

    test("reads an empty cell as empty text", () => {
      setContent([
        // B1 is empty
        ["A1", '=TEXT(B1,"0.00")'],
        ["A2", '=TEXT(B1,"0;0;0;""x""@")'],
        ["A3", "=TEXT(5,B1)"],
      ]);

      expect(cell("A1")).toBe("");
      expect(cell("A2")).toBe("x");
      // An empty format shows nothing for numbers
      expect(cell("A3")).toBe("");

      setCellContent("B1", 2.5);
      expect(cell("A1")).toBe("2.50");
      expect(cell("A2")).toBe("3");
      expect(cell("A3")).toBe(FormulaError.VALUE);
    });

    test("shows INFINITY whatever the format", () => {
      expect(
        formatAll([
          '=TEXT(1/0,"0.00")',
          '=TEXT(INFINITY,"yyyy-mm-dd")',
          '=TEXT(-INFINITY,"#,##0")',
        ])
      ).toEqual(["INFINITY", "INFINITY", "-INFINITY"]);
    });

    test("requires format_text to be text", () => {
      setContent([
        ["A1", "=TEXT(5,0)"],
        ["A2", "=TEXT(5,TRUE)"],
        ["A3", "=TEXT(5,INFINITY)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(errorMessage("A1")).toBe(
        "TEXT format_text must be text, got number"
      );
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(cell("A3")).toBe(FormulaError.VALUE);
    });

    test("rejects invalid format codes", () => {
      setContent([
        ["A1", '=TEXT(5,"""abc")'],
        ["A2", '=TEXT(5,"[foo]0")'],
        ["A3", '=TEXT(5,"0;0;0;0;0")'],
        ["A4", '=TEXT(5,"yyyy 0")'],
        ["A5", '=TEXT(5,"0.0e")'],
        ["A6", '=TEXT(5,"0 /")'],
        ["A7", '=TEXT(5,"Net 0")'],
        // Excel accepts a second decimal point
        ["A8", '=TEXT(1.25,"0.0.0")'],
      ]);

      for (const ref of ["A1", "A2", "A3", "A4", "A5", "A6", "A7"]) {
        expect(cell(ref)).toBe(FormulaError.VALUE);
      }
      expect(errorMessage("A2")).toBe(
        'TEXT format_text "[foo]0" is not a valid format: [foo] is not a valid format code'
      );
      expect(errorMessage("A4")).toContain(
        "date and time codes cannot be mixed with number codes"
      );
      expect(cell("A8")).toBe("1.2.5");
    });

    test("propagates errors, the value's first", () => {
      setContent([
        ["B1", "=CEILING(5, 0)"],
        ["B2", "=NOSUCHFUNCTION()"],
        ["A1", '=TEXT(B1,"0")'],
        ["A2", "=TEXT(5,B1)"],
        ["A3", "=TEXT(B2,B1)"],
        ["A4", "=TEXT(B1,B2)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.DIV0);
      expect(cell("A2")).toBe(FormulaError.DIV0);
      expect(cell("A3")).toBe(FormulaError.NAME);
      expect(cell("A4")).toBe(FormulaError.DIV0);
    });
  });

  describe("arguments", () => {
    test("reads an omitted value as 0 and an omitted format as empty", () => {
      setContent([
        ["A1", '=TEXT(,"0.00")'],
        ["A2", '=TEXT(,"yyyy-mm-dd")'],
        ["A3", "=TEXT(5,)"],
        ["A4", '=TEXT("abc",)'],
      ]);

      expect(cell("A1")).toBe("0.00");
      expect(cell("A2")).toBe("1900-01-00");
      expect(cell("A3")).toBe("");
      expect(cell("A4")).toBe("abc");
    });

    test("requires exactly two arguments", () => {
      setContent([
        ["A1", "=TEXT()"],
        ["A2", "=TEXT(5)"],
        ["A3", '=TEXT(5,"0","x")'],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(errorMessage("A2")).toBe(
        "TEXT function takes exactly 2 arguments"
      );
      expect(cell("A3")).toBe(FormulaError.VALUE);
    });

    test("recalculates when a referenced cell changes", () => {
      setContent([
        ["B1", 1234.5],
        ["B2", "#,##0.00"],
        ["A1", "=TEXT(B1,B2)"],
        ["A2", '="Total: "&TEXT(B1,"$#,##0")'],
      ]);
      expect(cell("A1")).toBe("1,234.50");
      expect(cell("A2")).toBe("Total: $1,235");

      setCellContent("B1", -9876.543);
      expect(cell("A1")).toBe("-9,876.54");
      expect(cell("A2")).toBe("Total: -$9,877");

      setCellContent("B2", "0.0%");
      expect(cell("A1")).toBe("-987654.3%");
    });
  });

  describe("arrays", () => {
    test("formats a range element by element and spills", () => {
      setContent([
        ["A1", 1],
        ["A2", 2.25],
        ["A4", "text"],
        ["A5", "=CEILING(5, 0)"],
        ["C1", '=TEXT(A1:A5,"0.0")'],
        ["E1", '=ROWS(TEXT(A1:A5,"0.0"))'],
      ]);

      expect(cell("C1")).toBe("1.0");
      expect(cell("C2")).toBe("2.3");
      // A3 is empty
      expect(cell("C3")).toBe("");
      expect(cell("C4")).toBe("text");
      // An error stays in its element
      expect(cell("C5")).toBe(FormulaError.DIV0);
      expect(cell("C6")).toBe("");
      expect(cell("E1")).toBe(5);

      setCellContent("A3", 7);
      expect(cell("C3")).toBe("7.0");
    });

    test("broadcasts values against formats", () => {
      setContent([
        ["A1", '=TEXT(1.5,{"0","0.0","0.00"})'],
        ["A2", '=TEXT({1;2},{"0","0.0"})'],
        ["A4", '=TEXT(SEQUENCE(3),"000")'],
        ["E1", '=TEXT({1,2,3},{"0","0.0","0.00"})'],
      ]);

      expect(cell("A1")).toBe("2");
      expect(cell("B1")).toBe("1.5");
      expect(cell("C1")).toBe("1.50");

      // Two cells of {1;2} by two formats
      expect(cell("A2")).toBe("1");
      expect(cell("B2")).toBe("1.0");
      expect(cell("A3")).toBe("2");
      expect(cell("B3")).toBe("2.0");

      expect(cell("A4")).toBe("001");
      expect(cell("A5")).toBe("002");
      expect(cell("A6")).toBe("003");

      expect(cell("E1")).toBe("1");
      expect(cell("F1")).toBe("2.0");
      expect(cell("G1")).toBe("3.00");
    });

    test("formats the whole spill of a referenced spilling cell", () => {
      setContent([
        ["A1", 3],
        ["B1", "=SEQUENCE(A1,1,45000)"],
        ["D1", '=TEXT(B1,"d mmm")'],
      ]);

      expect(cell("D1")).toBe("15 Mar");
      expect(cell("D2")).toBe("16 Mar");
      expect(cell("D3")).toBe("17 Mar");
      expect(cell("D4")).toBe("");

      setCellContent("A1", 4);
      expect(cell("D4")).toBe("18 Mar");
    });

    test("provides its cells to functions that read the whole array", () => {
      setContent([
        ["A1", 1],
        ["A2", 2.5],
        ["A3", 3],
        ["C1", '=TEXTJOIN(",",TRUE,TEXT(A1:A3,"0.0"))'],
        ["C2", '=TEXTJOIN(",",TRUE,TEXT(B:B,"0%"))'],
        ["C3", '=COUNTA(TEXT(A1:A3,"0"))'],
        ["C4", '=ROWS(UNIQUE(TEXT(A1:A3,"0")))'],
        ["B2", 0.25],
        ["B4", 0.5],
      ]);

      expect(cell("C1")).toBe("1.0,2.5,3.0");
      // An open-ended range formats its populated cells
      expect(cell("C2")).toBe("25%,50%");
      expect(cell("C3")).toBe(3);
      // 1, 3 and 3: 2.5 rounds to 3
      expect(cell("C4")).toBe(2);

      setCellContent("A2", 4);
      expect(cell("C1")).toBe("1.0,4.0,3.0");
      expect(cell("C4")).toBe(3);
    });
  });
});
