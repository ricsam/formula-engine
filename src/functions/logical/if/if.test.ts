import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("IF function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  /** Reads a block of cells row by row, e.g. block("C1", 3, 4). */
  const block = (topLeft: string, rows: number, cols: number) => {
    const start = parseCellReference(topLeft);
    const values: SerializedCellValue[][] = [];
    for (let y = 0; y < rows; y++) {
      const row: SerializedCellValue[] = [];
      for (let x = 0; x < cols; x++) {
        row.push(
          engine.getCellValue({
            sheetName,
            workbookName,
            colIndex: start.colIndex + x,
            rowIndex: start.rowIndex + y,
          })
        );
      }
      values.push(row);
    }
    return values;
  };

  const setContent = (entries: [string, SerializedCellValue][]) => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>(entries)
    );
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("should return true value when condition is true", () => {
      setCellContent("A1", '=IF(TRUE, "Yes", "No")');
      expect(cell("A1")).toBe("Yes");
    });

    test("should return false value when condition is false", () => {
      setCellContent("A1", '=IF(FALSE, "Yes", "No")');
      expect(cell("A1")).toBe("No");
    });

    test("should default to FALSE when no false value provided", () => {
      setCellContent("A1", '=IF(FALSE, "Yes")');
      expect(cell("A1")).toBe(false);
    });

    test("should handle numeric conditions", () => {
      setCellContent("A1", '=IF(1, "Truthy", "Falsy")');
      expect(cell("A1")).toBe("Truthy");

      setCellContent("A2", '=IF(0, "Truthy", "Falsy")');
      expect(cell("A2")).toBe("Falsy");
    });

    test("should handle string conditions", () => {
      setCellContent("A1", '=IF("Hello", "Non-empty", "Empty")');
      expect(cell("A1")).toBe("Non-empty");

      setCellContent("A2", '=IF("", "Non-empty", "Empty")');
      expect(cell("A2")).toBe("Empty");
    });

    test("should work with cell references and comparisons", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 5],
          ["C1", '=IF(A1>B1, "A1 is greater", "B1 is greater")'],
        ])
      );

      expect(cell("C1")).toBe("A1 is greater");
    });
  });

  describe("logical evaluation", () => {
    test("should treat zero as false", () => {
      setCellContent("A1", '=IF(0, "True", "False")');
      expect(cell("A1")).toBe("False");
    });

    test("should treat non-zero numbers as true", () => {
      setCellContent("A1", '=IF(42, "True", "False")');
      expect(cell("A1")).toBe("True");

      setCellContent("A2", '=IF(-1, "True", "False")');
      expect(cell("A2")).toBe("True");

      setCellContent("A3", '=IF(3.14, "True", "False")');
      expect(cell("A3")).toBe("True");
    });

    test("should treat empty string as false", () => {
      setCellContent("A1", '=IF("", "True", "False")');
      expect(cell("A1")).toBe("False");
    });

    test("should treat non-empty strings as true", () => {
      setCellContent("A1", '=IF("Hello", "True", "False")');
      expect(cell("A1")).toBe("True");

      setCellContent("A2", '=IF(" ", "True", "False")');
      expect(cell("A2")).toBe("True"); // Space is not empty
    });

    test("should handle boolean values", () => {
      setCellContent("A1", '=IF(TRUE, "True", "False")');
      expect(cell("A1")).toBe("True");

      setCellContent("A2", '=IF(FALSE, "True", "False")');
      expect(cell("A2")).toBe("False");
    });

    test("should treat INFINITY as true", () => {
      setCellContent("A1", '=IF("INFINITY", "True", "False")');
      expect(cell("A1")).toBe("True");
    });
  });

  describe("return value types", () => {
    test("should return numbers", () => {
      setCellContent("A1", "=IF(TRUE, 42, 0)");
      expect(cell("A1")).toBe(42);

      setCellContent("A2", "=IF(FALSE, 42, 0)");
      expect(cell("A2")).toBe(0);
    });

    test("should return strings", () => {
      setCellContent("A1", '=IF(TRUE, "Hello", "World")');
      expect(cell("A1")).toBe("Hello");

      setCellContent("A2", '=IF(FALSE, "Hello", "World")');
      expect(cell("A2")).toBe("World");
    });

    test("should return booleans", () => {
      setCellContent("A1", "=IF(TRUE, TRUE, FALSE)");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=IF(FALSE, TRUE, FALSE)");
      expect(cell("A2")).toBe(false);
    });

    test("should return mixed types", () => {
      setCellContent("A1", '=IF(TRUE, 42, "Text")');
      expect(cell("A1")).toBe(42);

      setCellContent("A2", '=IF(FALSE, 42, "Text")');
      expect(cell("A2")).toBe("Text");
    });
  });

  describe("nested IF statements", () => {
    test("should handle nested IF in true branch", () => {
      setCellContent(
        "A1",
        '=IF(TRUE, IF(TRUE, "Inner True", "Inner False"), "Outer False")'
      );
      expect(cell("A1")).toBe("Inner True");
    });

    test("should handle nested IF in false branch", () => {
      setCellContent(
        "A1",
        '=IF(FALSE, "Outer True", IF(TRUE, "Inner True", "Inner False"))'
      );
      expect(cell("A1")).toBe("Inner True");
    });

    test("should handle multiple levels of nesting", () => {
      setCellContent(
        "A1",
        '=IF(TRUE, IF(FALSE, "Level 2 True", IF(TRUE, "Level 3 True", "Level 3 False")), "Level 1 False")'
      );
      expect(cell("A1")).toBe("Level 3 True");
    });
  });

  describe("dynamic arrays (spilled values)", () => {
    test("should handle spilled logical test", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 0],
          ["A3", 5],
          ["B1", '=IF(A1:A3, "True", "False")'],
        ])
      );

      expect(cell("B1")).toBe("True"); // 1 is truthy
      expect(cell("B2")).toBe("False"); // 0 is falsy
      expect(cell("B3")).toBe("True"); // 5 is truthy
    });

    test("should handle spilled true values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Option A"],
          ["A2", "Option B"],
          ["A3", "Option C"],
          ["B1", '=IF(TRUE, A1:A3, "Default")'],
        ])
      );

      expect(cell("B1")).toBe("Option A");
      expect(cell("B2")).toBe("Option B");
      expect(cell("B3")).toBe("Option C");
    });

    test("should handle spilled false values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Default A"],
          ["A2", "Default B"],
          ["A3", "Default C"],
          ["B1", '=IF(FALSE, "True Value", A1:A3)'],
        ])
      );

      expect(cell("B1")).toBe("Default A");
      expect(cell("B2")).toBe("Default B");
      expect(cell("B3")).toBe("Default C");
    });

    test("should handle multiple spilled arrays", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 0],
          ["A3", 1],
          ["B1", "True A"],
          ["B2", "True B"],
          ["B3", "True C"],
          ["C1", "False A"],
          ["C2", "False B"],
          ["C3", "False C"],
          ["D1", "=IF(A1:A3, B1:B3, C1:C3)"],
        ])
      );

      expect(cell("D1")).toBe("True A"); // 1 -> True A
      expect(cell("D2")).toBe("False B"); // 0 -> False B
      expect(cell("D3")).toBe("True C"); // 1 -> True C
    });
  });

  describe("array broadcasting", () => {
    test("a row condition repeats a column value across the columns", () => {
      // The shape of IF(SEQUENCE(,ROWS(b)), a): every row repeats one value of a.
      setContent([
        ["A1", "a1"],
        ["A2", "a2"],
        ["A3", "a3"],
        ["C1", "=IF(SEQUENCE(,2), A1:A3)"],
      ]);

      expect(block("C1", 3, 2)).toEqual([
        ["a1", "a1"],
        ["a2", "a2"],
        ["a3", "a3"],
      ]);
      expect(cell("E1")).toBe("");
      expect(cell("C4")).toBe("");
    });

    test("a column condition repeats a row value down the rows", () => {
      // The shape of IF(SEQUENCE(ROWS(a)), TOROW(b)): every row is all of b.
      setContent([
        ["A1", "b1"],
        ["B1", "b2"],
        ["D1", "=IF(SEQUENCE(3), A1:B1)"],
      ]);

      expect(block("D1", 3, 2)).toEqual([
        ["b1", "b2"],
        ["b1", "b2"],
        ["b1", "b2"],
      ]);
      expect(cell("F1")).toBe("");
      expect(cell("D4")).toBe("");
    });

    test("the two shapes pair every value of a with every value of b", () => {
      setContent([
        ["A1", "x"],
        ["A2", "y"],
        ["B1", 1],
        ["C1", 2],
        ["D1", 3],
        ["F1", '=IF(SEQUENCE(,3), A1:A2)&"/"&IF(SEQUENCE(2), B1:D1)'],
      ]);

      expect(block("F1", 2, 3)).toEqual([
        ["x/1", "x/2", "x/3"],
        ["y/1", "y/2", "y/3"],
      ]);
    });

    test("condition, value and alternative are broadcast together", () => {
      setContent([
        ["A1", 1],
        ["A2", 5],
        ["C1", "=IF(A1:A2>2, SEQUENCE(1,3)*10, -SEQUENCE(1,3))"],
      ]);

      expect(block("C1", 2, 3)).toEqual([
        [-1, -2, -3],
        [10, 20, 30],
      ]);
    });

    test("a row condition chooses between two columns", () => {
      setContent([
        ["A1", "t1"],
        ["A2", "t2"],
        ["B1", "f1"],
        ["B2", "f2"],
        ["D1", "=IF(SEQUENCE(1,3)<>2, A1:A2, B1:B2)"],
        ["D4", '=IF({1,0,1}, {"a";"b"}, "-")'],
      ]);

      expect(block("D1", 2, 3)).toEqual([
        ["t1", "f1", "t1"],
        ["t2", "f2", "t2"],
      ]);
      expect(block("D4", 2, 3)).toEqual([
        ["a", "-", "a"],
        ["b", "-", "b"],
      ]);
    });

    test("single values fill the condition's shape", () => {
      setContent([
        ["A1", "=IF(SEQUENCE(2,2)>2, \"big\", \"small\")"],
        ["D1", "=IF(SEQUENCE(1,3)>1, \"yes\")"],
      ]);

      expect(block("A1", 2, 2)).toEqual([
        ["small", "small"],
        ["big", "big"],
      ]);
      // An omitted alternative is FALSE in every cell that does not pass.
      expect(block("D1", 1, 3)).toEqual([[false, "yes", "yes"]]);
    });

    test("a single-value condition returns the chosen array unchanged", () => {
      setContent([
        ["A1", "=IF(TRUE, SEQUENCE(2), SEQUENCE(1,3))"],
        ["C1", "=IF(FALSE, SEQUENCE(2), SEQUENCE(1,3))"],
      ]);

      expect(block("A1", 3, 1)).toEqual([[1], [2], [""]]);
      expect(block("C1", 1, 4)).toEqual([[1, 2, 3, ""]]);
    });

    test("cells beyond a shorter argument are #N/A", () => {
      setContent([
        ["A1", "=IF(SEQUENCE(3)>0, SEQUENCE(2)*10)"],
        ["C1", "=IF(SEQUENCE(1,2), SEQUENCE(1,3))"],
        ["C3", "=IF(SEQUENCE(2,2)>1, SEQUENCE(3), 0)"],
      ]);

      expect(block("A1", 3, 1)).toEqual([[10], [20], [FormulaError.NA]]);
      expect(block("C1", 1, 3)).toEqual([[1, 2, FormulaError.NA]]);
      // The 2x2 condition is shorter than the 3x1 value: its third row is #N/A.
      expect(block("C3", 3, 2)).toEqual([
        [0, 1],
        [2, 2],
        [FormulaError.NA, FormulaError.NA],
      ]);
    });

    test("omitted values are 0", () => {
      setContent([
        ["A1", "=IF(TRUE,,1)"],
        ["A2", "=IF(FALSE,1,)"],
        ["A3", "=IF(FALSE,1)"],
        ["A4", "=IF(,1,2)"],
        ["C1", '=IF(SEQUENCE(1,2)>1,,"x")'],
      ]);

      expect(cell("A1")).toBe(0);
      expect(cell("A2")).toBe(0);
      expect(cell("A3")).toBe(false);
      expect(cell("A4")).toBe(2);
      expect(block("C1", 1, 2)).toEqual([["x", 0]]);
    });
  });

  describe("array condition laziness", () => {
    test("each cell reads only the branch its condition chooses", () => {
      // B2 takes the alternative, so it never reads C2. Reading C2 from B2 would
      // close the cycle B2 -> C2 -> B1 -> C1 -> B2.
      setContent([
        ["A1", 1],
        ["A2", 0],
        ["B1", "=IF(A1:A2=1, C1:C2, 200)"],
        ["C1", "=B2"],
        ["C2", "=B1"],
      ]);

      expect(cell("B1")).toBe(200);
      expect(cell("B2")).toBe(200);
      expect(cell("C1")).toBe(200);
      expect(cell("C2")).toBe(200);
    });

    test("an error in a single-value branch only shows where it is chosen", () => {
      setContent([
        ["A1", 1],
        ["A2", 5],
        ["A3", 2],
        ["B1", '=IF(A1:A3>2, "big", CEILING(5))'],
      ]);

      expect(block("B1", 3, 1)).toEqual([
        [FormulaError.VALUE],
        ["big"],
        [FormulaError.VALUE],
      ]);
    });

    test("errors in the condition stay per cell", () => {
      setContent([
        ["A1", 1],
        ["A2", "=CEILING(5)"],
        ["A3", 0],
        ["B1", '=IF(A1:A3, "yes", "no")'],
      ]);

      expect(block("B1", 3, 1)).toEqual([
        ["yes"],
        [FormulaError.VALUE],
        ["no"],
      ]);
    });
  });

  describe("aggregates over an array condition", () => {
    test("SUM and COUNT read every broadcast cell", () => {
      setContent([
        ["A1", "=SUM(IF(SEQUENCE(1,3), SEQUENCE(2)))"],
        ["A2", "=COUNT(IF(SEQUENCE(1,3), SEQUENCE(2)))"],
        ["A3", "=SUM(IF(SEQUENCE(4)>2, SEQUENCE(4), 100))"],
        ["A4", "=SUM(IF(SEQUENCE(2), SEQUENCE(1,3)*10))"],
      ]);

      expect(cell("A1")).toBe(9);
      expect(cell("A2")).toBe(6);
      expect(cell("A3")).toBe(207);
      expect(cell("A4")).toBe(120);
    });

    test("conditional sums over ranges", () => {
      setContent([
        ["A1", 1],
        ["A2", 8],
        ["A3", 3],
        ["A4", 9],
        ["B1", 10],
        ["B2", 20],
        ["B3", 30],
        ["B4", 40],
        ["D1", "=SUM(IF(A1:A4>2, B1:B4))"],
        ["D2", "=SUM(IF(A1:A4>2, B1:B4, -B1:B4))"],
        ["D3", "=COUNT(IF(A1:A4>5, A1:A4))"],
      ]);

      expect(cell("D1")).toBe(90);
      expect(cell("D2")).toBe(80);
      expect(cell("D3")).toBe(2);
    });

    test("open-ended conditions read only the populated cells", () => {
      setContent([
        ["A1", 3],
        ["A2", 7],
        ["A3", 9],
        ["B1", 1],
        ["B2", 2],
        ["B3", 4],
        ["D1", "=SUM(IF(A:A>5, A:A))"],
        ["D2", "=SUM(IF(A:A>5, B:B, 100))"],
        ["D3", "=COUNT(IF(A:A>5, A:A))"],
        ["D4", "=SUM(IF(SEQUENCE(1,3), A:A))"],
      ]);

      expect(cell("D1")).toBe(16);
      expect(cell("D2")).toBe(106);
      expect(cell("D3")).toBe(2);
      expect(cell("D4")).toBe(57);
    });

    test("UNIQUE reads an IF result", () => {
      setContent([
        ["A1", 1],
        ["A2", 5],
        ["A3", 7],
        ["C1", '=UNIQUE(IF(A1:A3>2, "big", "small"))'],
        ["E1", "=UNIQUE(IF(SEQUENCE(,2), A1:A3))"],
      ]);

      expect(block("C1", 3, 1)).toEqual([["small"], ["big"], [""]]);
      expect(block("E1", 3, 2)).toEqual([
        [1, 1],
        [5, 5],
        [7, 7],
      ]);
    });

    test("an open range reads the part of an IF spill it overlaps", () => {
      setContent([
        ["A10", "=IF(SEQUENCE(1,3), SEQUENCE(3))"],
        ["A1", "=SUM(B10:B)"],
        ["A2", "=SUM(B11:INFINITY)"],
      ]);

      // A10:C12 holds [1,1,1; 2,2,2; 3,3,3]
      expect(cell("A1")).toBe(6);
      expect(cell("A2")).toBe(10);
    });
  });

  describe("recalculation with an array condition", () => {
    test("spilled cells and aggregates follow edits", () => {
      setContent([
        ["A1", 1],
        ["A2", 5],
        ["B1", 10],
        ["B2", 20],
        ["D1", '=IF(A1:A2>2, B1:B2, "low")'],
        ["E1", "=SUM(IF(A1:A2>2, B1:B2))"],
      ]);

      expect(block("D1", 2, 1)).toEqual([["low"], [20]]);
      expect(cell("E1")).toBe(20);

      setCellContent("A1", 3);
      expect(block("D1", 2, 1)).toEqual([[10], [20]]);
      expect(cell("E1")).toBe(30);

      setCellContent("B2", 200);
      expect(block("D1", 2, 1)).toEqual([[10], [200]]);
      expect(cell("E1")).toBe(210);

      setCellContent("A2", 0);
      expect(block("D1", 2, 1)).toEqual([[10], ["low"]]);
      expect(cell("E1")).toBe(10);
    });

    test("the broadcast shape follows its arguments", () => {
      setContent([
        ["A1", 2],
        ["B1", 3],
        ["D1", "=IF(SEQUENCE(1,B1), SEQUENCE(A1))"],
      ]);

      expect(block("D1", 2, 3)).toEqual([
        [1, 1, 1],
        [2, 2, 2],
      ]);
      expect(cell("G1")).toBe("");

      setCellContent("B1", 4);
      expect(cell("G2")).toBe(2);

      setCellContent("A1", 1);
      expect(cell("G1")).toBe(1);
      expect(cell("D2")).toBe("");
    });
  });

  describe("error handling", () => {
    test("should return error for wrong number of arguments", () => {
      setCellContent("A1", "=IF()");
      expect(cell("A1")).toBe("#VALUE!");

      setCellContent("A2", "=IF(TRUE)");
      expect(cell("A2")).toBe("#VALUE!");

      setCellContent("A3", '=IF(TRUE, "A", "B", "C")');
      expect(cell("A3")).toBe("#VALUE!");
    });

    test("should propagate errors from logical test", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // Results in INFINITY
          ["B1", '=IF(A1, "True", "False")'],
        ])
      );

      expect(cell("B1")).toBe("True"); // INFINITY is truthy
    });

    test("should propagate errors from true value", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // Results in INFINITY
          ["B1", '=IF(TRUE, A1, "False")'],
        ])
      );

      expect(cell("B1")).toBe("INFINITY");
    });

    test("should propagate errors from false value", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // Results in INFINITY
          ["B1", '=IF(FALSE, "True", A1)'],
        ])
      );

      expect(cell("B1")).toBe("INFINITY");
    });

    test("should handle errors in spilled arrays", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", "=1/0"], // Error in logical test array
          ["A3", 0],
          ["B1", '=IF(A1:A3, "True", "False")'],
        ])
      );

      expect(cell("B1")).toBe("True"); // 1 is truthy
      expect(cell("B2")).toBe("True"); // INFINITY is truthy
      expect(cell("B3")).toBe("False"); // 0 is falsy
    });
  });

  describe("edge cases", () => {
    test("should handle empty cells in logical test", () => {
      setCellContent("B1", '=IF(A1, "True", "False")');
      // A1 is empty, should be falsy
      expect(cell("B1")).toBe("False");
    });

    test("should handle complex expressions with comparisons", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 5],
          ["C1", 15],
          ["D1", '=IF(A1+B1>C1, "Sum is greater", "Sum is not greater")'],
        ])
      );

      expect(cell("D1")).toBe("Sum is not greater"); // 10+5 = 15, not > 15
    });

    test("should handle string equality comparisons", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Apple"],
          ["C1", "Banana"],
          ["D1", '=IF(A1=B1, "Match", "No match")'],
          ["D2", '=IF(A1=C1, "Match", "No match")'],
        ])
      );

      expect(cell("D1")).toBe("Match");
      expect(cell("D2")).toBe("No match");
    });

    test("should work with other functions and comparisons", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "World"],
          [
            "B1",
            '=IF(LEN(A1)>LEN(A2), CONCATENATE(A1, " is longer"), CONCATENATE(A2, " is longer"))',
          ],
        ])
      );

      expect(cell("B1")).toBe("World is longer"); // "Hello" (5) vs "World" (5), equal so not >
    });
  });

  describe("lazy evaluation", () => {
    test("should not evaluate false branch when condition is true", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          // If both branches were evaluated, this would cause an error
          // because CEILING with 1 argument is invalid
          ["B1", '=IF(TRUE, "Success", CEILING(5))'],
        ])
      );

      expect(cell("B1")).toBe("Success");
    });

    test("should not evaluate true branch when condition is false", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          // If both branches were evaluated, this would cause an error
          ["B1", '=IF(FALSE, CEILING(5), "Success")'],
        ])
      );

      expect(cell("B1")).toBe("Success");
    });

    test("should prevent cyclic dependencies with same logical test", () => {
      // This is a key use case: two IF cells that would create a cycle
      // if both branches were evaluated, but the same logical test
      // ensures only non-cyclic branches are evaluated
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1], // Control flag: 1 = use B1, 0 = use static value
          ["B1", "=IF(A1=1, 100, C1)"], // When A1=1, use 100; else use C1
          ["C1", "=IF(A1=1, B1, 200)"], // When A1=1, use B1; else use 200
        ])
      );

      // When A1=1: B1 uses 100, C1 uses B1 (which is 100) -> C1=100
      expect(cell("B1")).toBe(100);
      expect(cell("C1")).toBe(100);

      // Change the condition
      setCellContent("A1", 0);

      // When A1=0: B1 uses C1 (which is 200), C1 uses 200 -> B1=200
      expect(cell("B1")).toBe(200);
      expect(cell("C1")).toBe(200);
    });

    test("should work with nested lazy evaluation", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", true],
          // Nested IF with errors in unevaluated branches
          [
            "B1",
            '=IF(A1, IF(TRUE, "Nested Success", CEILING(1)), CEILING(2))',
          ],
        ])
      );

      expect(cell("B1")).toBe("Nested Success");
    });

    test("should evaluate branches with actual errors correctly", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 0],
          // When condition is true, we should get the error from the true branch
          ["C1", "=IF(TRUE, CEILING(5), 100)"], // Missing arg -> error
        ])
      );

      expect(cell("C1")).toBe("#VALUE!"); // Error is properly returned
    });
  });
});
