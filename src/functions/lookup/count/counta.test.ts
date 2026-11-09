import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("COUNTA function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) });

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const setContent = (content: Array<[string, SerializedCellValue]>) => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>(content)
    );
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("ranges", () => {
    test("counts numbers, text, logical values and errors but not empty cells", () => {
      setContent([
        ["A1", 1],
        ["A2", "text"],
        ["A3", true],
        ["A4", false],
        ["A5", "=CEILING()"],
        // A6 is empty
        ["A7", 0],
        ["A8", "=1/0"],
        ["C1", "=COUNTA(A1:A10)"],
      ]);

      expect(cell("A5")).toBe(FormulaError.VALUE);
      expect(cell("C1")).toBe(7);
    });

    test("counts errors inside a range instead of propagating them", () => {
      setContent([
        ["A1", "=CEILING(5, 0)"],
        ["A2", "=CEILING()"],
        ["C1", "=COUNTA(A1:A3)"],
        ["C2", "=COUNT(A1:A3)"],
      ]);

      expect(cell("C1")).toBe(2);
      // COUNT keeps propagating errors.
      expect(cell("C2")).toBe(FormulaError.DIV0);
    });

    test("counts a cell whose formula returns empty text", () => {
      setContent([
        ["A1", '=""'],
        ["A2", '=LEFT("abc", 0)'],
        // A3 is empty
        ["A4", "=A3"],
        ["C1", "=COUNTA(A1:A4)"],
      ]);

      expect(cell("A1")).toBe("");
      expect(cell("C1")).toBe(3);
    });

    test("counts the populated cells of open-ended ranges", () => {
      setContent([
        ["A1", "header"],
        ["A5", 5],
        ["A1000", "=CEILING()"],
        ["B2", "x"],
        ["F2", "y"],
        ["D3", true],
        // The formulas sit in row 1, outside the ranges they count.
        ["E1", "=COUNTA(A:A)"],
        ["F1", "=COUNTA(A2:A)"],
        ["G1", "=COUNTA(B2:2)"],
        ["H1", "=COUNTA(B2:INFINITY)"],
      ]);

      expect(cell("E1")).toBe(3);
      expect(cell("F1")).toBe(2);
      expect(cell("G1")).toBe(2);
      expect(cell("H1")).toBe(3);

      setCellContent("Z500", 1);
      setCellContent("A5", "");
      expect(cell("E1")).toBe(2);
      expect(cell("F1")).toBe(1);
      expect(cell("H1")).toBe(4);
    });

    test("counts cells that other formulas spill onto", () => {
      setContent([
        ["C1", "=SEQUENCE(3)"],
        ["E1", "=COUNTA(C1:C10)"],
        ["E2", "=COUNTA(C2:C10)"],
        ["E3", "=COUNTA(C:C)"],
      ]);

      expect(cell("E1")).toBe(3);
      expect(cell("E2")).toBe(2);
      expect(cell("E3")).toBe(3);

      setCellContent("C1", "=SEQUENCE(5)");
      expect(cell("E1")).toBe(5);
      expect(cell("E3")).toBe(5);
    });

    test("counts ranges on other sheets", () => {
      engine.addSheet({ workbookName, sheetName: "Data" });
      engine.setSheetContent(
        { workbookName, sheetName: "Data" },
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A3", "b"],
        ])
      );
      setCellContent("A1", "=COUNTA(Data!A:A)");

      expect(cell("A1")).toBe(2);
    });
  });

  describe("directly supplied values", () => {
    test("counts every supplied value, including empty text and errors", () => {
      setContent([
        ["A1", '=COUNTA(1, "text", TRUE, FALSE)'],
        ["A2", '=COUNTA("")'],
        ["A3", "=COUNTA(CEILING())"],
        ["A4", "=COUNTA(0, INFINITY)"],
        ["A5", '=COUNTA("", "", "")'],
      ]);

      expect(cell("A1")).toBe(4);
      expect(cell("A2")).toBe(1);
      expect(cell("A3")).toBe(1);
      expect(cell("A4")).toBe(2);
      expect(cell("A5")).toBe(3);
    });

    test("counts a referenced cell unless it is empty", () => {
      setContent([
        ["B1", 5],
        ["B2", "x"],
        ["B3", "=CEILING()"],
        // B4 is empty
        ["A1", "=COUNTA(B1)"],
        ["A2", "=COUNTA(B2)"],
        ["A3", "=COUNTA(B3)"],
        ["A4", "=COUNTA(B4)"],
        ["A5", "=COUNTA(B1, B2, B3, B4)"],
      ]);

      expect(cell("A1")).toBe(1);
      expect(cell("A2")).toBe(1);
      expect(cell("A3")).toBe(1);
      expect(cell("A4")).toBe(0);
      expect(cell("A5")).toBe(3);

      setCellContent("B4", false);
      expect(cell("A4")).toBe(1);
      expect(cell("A5")).toBe(4);

      setCellContent("B1", "");
      expect(cell("A1")).toBe(0);
      expect(cell("A5")).toBe(3);
    });

    test("counts a direct reference to a formula returning empty text as empty", () => {
      // The engine reads an empty cell as "", so a single-cell reference cannot
      // tell the two apart (Excel counts this cell). Inside a range it counts.
      setContent([
        ["B1", '=""'],
        ["A1", "=COUNTA(B1)"],
        ["A2", "=COUNTA(B1:B2)"],
      ]);

      expect(cell("A1")).toBe(0);
      expect(cell("A2")).toBe(1);
    });

    test("requires at least one argument", () => {
      setCellContent("A1", "=COUNTA()");

      expect(cell("A1")).toBe(FormulaError.VALUE);
    });
  });

  describe("spilled arguments", () => {
    test("counts the cells of spilled results", () => {
      setContent([
        ["B1", "a"],
        ["B2", "b"],
        ["B3", "a"],
        ["B4", "c"],
        ["A1", "=COUNTA(SEQUENCE(3, 2))"],
        ["A2", "=COUNTA(UNIQUE(B1:B4))"],
        ["A3", "=COUNTA(SEQUENCE(4)/2)"],
        ["A4", "=COUNTA(SEQUENCE(2), B1:B4, \"x\")"],
      ]);

      expect(cell("A1")).toBe(6);
      expect(cell("A2")).toBe(3);
      expect(cell("A3")).toBe(4);
      expect(cell("A4")).toBe(7);

      setCellContent("B3", "d");
      expect(cell("A2")).toBe(4);
    });

    test("counts errors inside spilled results", () => {
      setContent([
        ["B1", 1.5],
        ["B2", "text"],
        ["B3", 2.5],
        ["A1", "=COUNTA(INT(B1:B3))"],
      ]);

      // INT("text") is #VALUE!, which counts as a value.
      expect(cell("A1")).toBe(3);
    });

    test("counts the whole spill of a referenced spilling cell", () => {
      setContent([
        ["A1", 4],
        ["C1", "=SEQUENCE(A1)"],
        ["E1", "=COUNTA(C1)"],
      ]);

      expect(cell("E1")).toBe(4);

      setCellContent("A1", 2);
      expect(cell("E1")).toBe(2);
    });

    test("returns the error when an argument's cells cannot be read", () => {
      setCellContent("A1", "=COUNTA(Missing!A1:A3)");

      expect(cell("A1")).toBe(FormulaError.REF);

      engine.addSheet({ workbookName, sheetName: "Missing" });
      engine.setCellContent(
        { workbookName, sheetName: "Missing", ...parseCellReference("A2") },
        "found"
      );
      expect(cell("A1")).toBe(1);
    });
  });

  describe("recalculation", () => {
    test("follows edits to the counted range", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", "=COUNTA(A1:A5)"],
      ]);
      expect(cell("C1")).toBe(2);

      setCellContent("A4", "new");
      expect(cell("C1")).toBe(3);

      setCellContent("A1", "");
      expect(cell("C1")).toBe(2);

      setCellContent("A2", '=""');
      expect(cell("C1")).toBe(2);

      setCellContent("A6", "outside");
      expect(cell("C1")).toBe(2);
    });

    test("waits for formulas inside the range to evaluate", () => {
      setContent([
        ["A1", "=B1"],
        ["A2", "=A1 + 1"],
        ["A3", "=SEQUENCE(B2)"],
        ["B1", 1],
        ["B2", 2],
        ["C1", "=COUNTA(A:A)"],
      ]);

      // A1, A2 and the SEQUENCE spill over A3:A4
      expect(cell("C1")).toBe(4);

      setCellContent("B2", 3);
      expect(cell("C1")).toBe(5);
    });
  });

  describe("results of array functions", () => {
    // Wells in A1:A6 with a treatment in B1:B6; the IDs of untreated wells are
    // missing, as in a layout map.
    const layout: Array<[string, SerializedCellValue]> = [
      ["A1", "W1"],
      ["B1", "MAP"],
      ["A2", "W2"],
      ["B2", "control"],
      ["A3", "W3"],
      ["B3", "MAP"],
      // A4 is empty
      ["B4", "MAP"],
      ["A5", '=""'],
      ["B5", "MAP"],
      ["A6", "=NOSUCHFN()"],
      ["B6", "control"],
    ];

    test("counts the populated cells a FILTER selects", () => {
      setContent([
        ...layout,
        ["D1", '=COUNTA(FILTER(A1:A6, B1:B6="MAP"))'],
        ["D2", '=COUNTA(FILTER(A1:A6, B1:B6="control"))'],
        ["D3", '=COUNTA(FILTER(A1:B6, B1:B6="MAP"))'],
        ["D4", '=ROWS(FILTER(A1:A6, B1:B6="MAP"))'],
      ]);

      // W1, W3 and the formula returning "" in A5; A4 is empty
      expect(cell("D1")).toBe(3);
      // W2 and the error in A6
      expect(cell("D2")).toBe(2);
      // the four MAP rows of B and three cells of A
      expect(cell("D3")).toBe(7);
      // every selected row, whatever it holds
      expect(cell("D4")).toBe(4);

      setCellContent("A4", "W4");
      expect(cell("D1")).toBe(4);
    });

    test("counts the cells of TRANSPOSE, IFS and VLOOKUP results", () => {
      setContent([
        ...layout,
        ["D1", "=COUNTA(TRANSPOSE(A1:B6))"],
        ["D2", '=COUNTA(IFS(B1:B6="MAP", A1:A6, TRUE, "none"))'],
        ["D3", '=COUNTA(VLOOKUP({"W1";"W2";"W9"}, A1:B6, 2, FALSE))'],
      ]);

      // the populated cells of A1:B6
      expect(cell("D1")).toBe(11);
      // one cell per row of the broadcast result
      expect(cell("D2")).toBe(6);
      // MAP, control and #N/A
      expect(cell("D3")).toBe(3);
    });

    test("counts every cell of an IS function over a bounded range", () => {
      setContent([
        ...layout,
        ["D1", "=COUNTA(ISBLANK(A1:A6))"],
        ["D2", "=SUM(IF(ISBLANK(A1:A6), 1, 0))"],
      ]);

      expect(cell("D1")).toBe(6);
      // A4, and A5 whose formula returns ""
      expect(cell("D2")).toBe(2);
    });
  });
});
