import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("TOCOL function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string) =>
    engine.getCellValue({
      sheetName,
      workbookName,
      ...parseCellReference(ref),
    });

  /** Values of `column` from row `from` to row `to`, inclusive. */
  const column = (columnName: string, from: number, to: number) => {
    const values: SerializedCellValue[] = [];
    for (let row = from; row <= to; row++) {
      values.push(cell(`${columnName}${row}`));
    }
    return values;
  };

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

  /**
   * A1:B3 holds    1 | #NAME?
   *            blank | 4
   *              "c" | #VALUE!
   */
  const gridWithBlanksAndErrors: Array<[string, SerializedCellValue]> = [
    ["A1", 1],
    ["B1", "=NOSUCHFN()"],
    ["B2", 4],
    ["A3", "c"],
    ["B3", '="a"+1'],
  ];

  describe("scan order", () => {
    test("reads a 2D range row by row by default", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["D1", "=TOCOL(A1:B2)"],
      ]);

      expect(column("D", 1, 5)).toEqual([1, 2, 3, 4, ""]);
    });

    test("reads column by column when scan_by_column is TRUE", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["D1", "=TOCOL(A1:B2,0,TRUE)"],
        ["E1", "=TOCOL(A1:B2,0,FALSE)"],
        ["F1", "=TOCOL(A1:B2,0,1)"],
      ]);

      expect(column("D", 1, 5)).toEqual([1, 3, 2, 4, ""]);
      expect(column("E", 1, 5)).toEqual([1, 2, 3, 4, ""]);
      expect(column("F", 1, 5)).toEqual([1, 3, 2, 4, ""]);
    });

    test("keeps a column as it is", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "c"],
        ["C1", "=TOCOL(A1:A3)"],
      ]);

      expect(column("C", 1, 4)).toEqual(["a", "b", "c", ""]);
    });

    test("turns a row into a column", () => {
      setContent([
        ["A1", "a"],
        ["B1", "b"],
        ["C1", "c"],
        ["A3", "=TOCOL(A1:C1)"],
      ]);

      expect(column("A", 3, 6)).toEqual(["a", "b", "c", ""]);
      expect(cell("B3")).toBe("");
    });

    test("spills away from the anchor cell", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["E5", "=TOCOL(A1:B1)"],
      ]);

      expect(column("E", 5, 7)).toEqual([1, 2, ""]);
    });
  });

  describe("ignore", () => {
    test("0 keeps blanks and errors", () => {
      setContent([...gridWithBlanksAndErrors, ["D1", "=TOCOL(A1:B3,0)"]]);

      expect(column("D", 1, 7)).toEqual([
        1,
        FormulaError.NAME,
        "",
        4,
        "c",
        FormulaError.VALUE,
        "",
      ]);
    });

    test("1 skips blanks and keeps errors", () => {
      setContent([...gridWithBlanksAndErrors, ["D1", "=TOCOL(A1:B3,1)"]]);

      expect(column("D", 1, 6)).toEqual([
        1,
        FormulaError.NAME,
        4,
        "c",
        FormulaError.VALUE,
        "",
      ]);
    });

    test("2 skips errors and keeps blanks", () => {
      setContent([...gridWithBlanksAndErrors, ["D1", "=TOCOL(A1:B3,2)"]]);

      expect(column("D", 1, 5)).toEqual([1, "", 4, "c", ""]);
    });

    test("3 skips blanks and errors", () => {
      setContent([...gridWithBlanksAndErrors, ["D1", "=TOCOL(A1:B3,3)"]]);

      expect(column("D", 1, 4)).toEqual([1, 4, "c", ""]);
    });

    test("every mode also applies when scanning by column", () => {
      setContent([
        ...gridWithBlanksAndErrors,
        ["D1", "=TOCOL(A1:B3,0,TRUE)"],
        ["E1", "=TOCOL(A1:B3,1,TRUE)"],
        ["F1", "=TOCOL(A1:B3,2,TRUE)"],
        ["G1", "=TOCOL(A1:B3,3,TRUE)"],
      ]);

      expect(column("D", 1, 7)).toEqual([
        1,
        "",
        "c",
        FormulaError.NAME,
        4,
        FormulaError.VALUE,
        "",
      ]);
      expect(column("E", 1, 6)).toEqual([
        1,
        "c",
        FormulaError.NAME,
        4,
        FormulaError.VALUE,
        "",
      ]);
      expect(column("F", 1, 5)).toEqual([1, "", "c", 4, ""]);
      expect(column("G", 1, 4)).toEqual([1, "c", 4, ""]);
    });

    test("treats a cell holding empty text as blank", () => {
      setContent([
        ["A1", "a"],
        ["A2", '=""'],
        ["A3", "b"],
        ["C1", "=TOCOL(A1:A3,1)"],
      ]);

      expect(column("C", 1, 3)).toEqual(["a", "b", ""]);
    });

    test("truncates numbers and accepts logical values", () => {
      setContent([
        ["A1", "a"],
        ["A3", "b"],
        ["C1", "=TOCOL(A1:A3,1.9)"],
        ["D1", "=TOCOL(A1:A3,TRUE)"],
        ["E1", "=TOCOL(A1:A3,FALSE)"],
      ]);

      expect(column("C", 1, 3)).toEqual(["a", "b", ""]);
      expect(column("D", 1, 3)).toEqual(["a", "b", ""]);
      expect(column("E", 1, 3)).toEqual(["a", "", "b"]);
    });

    test("reads ignore from a cell", () => {
      setContent([
        ["A1", "a"],
        ["A3", "b"],
        ["B1", 1],
        ["C1", "=TOCOL(A1:A3,B1)"],
      ]);

      expect(column("C", 1, 3)).toEqual(["a", "b", ""]);

      setCellContent("B1", 0);
      expect(column("C", 1, 3)).toEqual(["a", "", "b"]);
    });

    test("rejects other values with #VALUE!", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", "=TOCOL(A1:A2,4)"],
        ["C2", "=TOCOL(A1:A2,-1)"],
        ["C3", '=TOCOL(A1:A2,"skip")'],
        ["C4", "=TOCOL(A1:A2,A1:A2)"],
        ["C5", '=TOCOL(A1:A2,0,"no")'],
      ]);

      expect(cell("C1")).toBe(FormulaError.VALUE);
      expect(cell("C2")).toBe(FormulaError.VALUE);
      expect(cell("C3")).toBe(FormulaError.VALUE);
      expect(cell("C4")).toBe(FormulaError.VALUE);
      expect(cell("C5")).toBe(FormulaError.VALUE);
    });

    test("propagates an error passed as ignore", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", "=TOCOL(A1:A2,NOSUCHFN())"],
      ]);

      expect(cell("C1")).toBe(FormulaError.NAME);
    });
  });

  describe("single values", () => {
    test("treats a value as a 1x1 array", () => {
      setContent([
        ["A1", 7],
        ["C1", "=TOCOL(5)"],
        ["C2", "=TOCOL(A1)"],
        ["C3", '=TOCOL("text",1)'],
      ]);

      expect(cell("C1")).toBe(5);
      expect(cell("C2")).toBe(7);
      expect(cell("C3")).toBe("text");
    });

    test("treats an error as a 1x1 array holding it", () => {
      setContent([
        ["C1", "=TOCOL(NOSUCHFN())"],
        ["C2", "=TOCOL(NOSUCHFN(),1)"],
        ["C3", "=TOCOL(NOSUCHFN(),2)"],
      ]);

      expect(cell("C1")).toBe(FormulaError.NAME);
      expect(cell("C2")).toBe(FormulaError.NAME);
      expect(cell("C3")).toBe(FormulaError.CALC);
    });

    test("collapses a one-cell result to a single value", () => {
      setContent([
        ["A2", "only"],
        ["C1", "=TOCOL(A1:A5,1)"],
      ]);

      expect(cell("C1")).toBe("only");
      expect(cell("C2")).toBe("");
    });
  });

  describe("#CALC!", () => {
    test("is the result when every cell is skipped", () => {
      setContent([
        ["A1", "=NOSUCHFN()"],
        ["C1", "=TOCOL(A1:B2,1,FALSE)"],
        ["C2", "=TOCOL(A2:B3,1)"],
        ["C3", "=TOCOL(A1:B2,3)"],
        ["C4", "=TOCOL(A1,2)"],
        ["C5", "=TOCOL(A:A,3)"],
      ]);

      // A1:B2 still holds an error, which ignore=1 keeps.
      expect(cell("C1")).toBe(FormulaError.NAME);
      expect(cell("C2")).toBe(FormulaError.CALC);
      expect(cell("C3")).toBe(FormulaError.CALC);
      expect(cell("C4")).toBe(FormulaError.CALC);
      expect(cell("C5")).toBe(FormulaError.CALC);
    });
  });

  describe("open-ended ranges", () => {
    test("skips the unbounded blank tail of a whole column", () => {
      setContent([
        ["A2", "a"],
        ["A4", "b"],
        ["A7", "c"],
        ["C1", "=TOCOL(A:A,1)"],
        ["D1", "=TOCOL(A2:A,1)"],
      ]);

      expect(column("C", 1, 4)).toEqual(["a", "b", "c", ""]);
      expect(column("D", 1, 4)).toEqual(["a", "b", "c", ""]);
    });

    test("skips blanks and errors of unbounded columns, in both scan orders", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A3", 3],
        ["B4", "=NOSUCHFN()"],
        ["B5", 5],
        ["D1", "=TOCOL(A:B,3)"],
        ["E1", "=TOCOL(A:B,3,TRUE)"],
        ["F1", "=TOCOL(A:B,1,TRUE)"],
      ]);

      expect(column("D", 1, 5)).toEqual([1, 2, 3, 5, ""]);
      expect(column("E", 1, 5)).toEqual([1, 3, 2, 5, ""]);
      expect(column("F", 1, 6)).toEqual([1, 3, 2, FormulaError.NAME, 5, ""]);
    });

    test("skips the unbounded blank tail of a whole row", () => {
      setContent([
        ["B1", "x"],
        ["E1", "y"],
        ["A3", "=TOCOL(1:1,1)"],
      ]);

      expect(column("A", 3, 5)).toEqual(["x", "y", ""]);
    });

    test("maps a whole column lazily when blanks are kept", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A4", 4],
        ["C1", "=TOCOL(A:A)"],
      ]);

      expect(column("C", 1, 5)).toEqual([1, 2, "", 4, ""]);
      expect(cell("C1000")).toBe("");

      setCellContent("A1000", "far");
      expect(cell("C1000")).toBe("far");
    });

    test("reads the rows of unbounded columns one after the other when blanks are kept", () => {
      setContent([
        ["A1", 1],
        ["B2", 2],
        ["A3", 3],
        ["D1", "=TOCOL(A:B)"],
        ["E1", "=TOCOL(A:B,2)"],
      ]);

      expect(column("D", 1, 7)).toEqual([1, "", "", 2, 3, "", ""]);
      expect(column("E", 1, 7)).toEqual([1, "", "", 2, 3, "", ""]);
    });

    test("skips errors of a whole column and keeps its blanks", () => {
      setContent([
        ["A1", 1],
        ["A2", "=NOSUCHFN()"],
        ["A4", 4],
        ["C1", "=TOCOL(A:A,2)"],
      ]);

      expect(column("C", 1, 5)).toEqual([1, "", 4, "", ""]);
    });

    test("rejects keeping blanks when the first scanned column never ends", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["D1", "=TOCOL(A:B,0,TRUE)"],
        ["D2", "=TOCOL(A:B,2,TRUE)"],
        ["E1", "=TOCOL(1:2)"],
      ]);

      expect(cell("D1")).toBe(FormulaError.VALUE);
      expect(cell("D2")).toBe(FormulaError.VALUE);
      expect(cell("E1")).toBe(FormulaError.VALUE);
    });
  });

  describe("spilled inputs", () => {
    test("flattens a SEQUENCE", () => {
      setContent([
        ["A1", "=TOCOL(SEQUENCE(2,3))"],
        ["B1", "=TOCOL(SEQUENCE(2,3),0,TRUE)"],
      ]);

      expect(column("A", 1, 7)).toEqual([1, 2, 3, 4, 5, 6, ""]);
      expect(column("B", 1, 7)).toEqual([1, 4, 2, 5, 3, 6, ""]);
    });

    test("flattens a UNIQUE result", () => {
      setContent([
        ["A1", "x"],
        ["B1", 1],
        ["A2", "y"],
        ["B2", 2],
        ["A3", "x"],
        ["B3", 1],
        ["D1", "=TOCOL(UNIQUE(A1:B3))"],
      ]);

      expect(column("D", 1, 5)).toEqual(["x", 1, "y", 2, ""]);
    });

    test("flattens a reference to a spilling cell", () => {
      setContent([
        ["A1", "=SEQUENCE(2,2)"],
        ["D1", "=TOCOL(A1)"],
        ["E1", "=TOCOL(A1,1,TRUE)"],
      ]);

      expect(column("D", 1, 5)).toEqual([1, 2, 3, 4, ""]);
      expect(column("E", 1, 5)).toEqual([1, 3, 2, 4, ""]);
    });

    test("flattens an array literal", () => {
      setContent([
        ["A1", "=TOCOL({1,2;3,4})"],
        ["B1", "=SUM(TOCOL({1,2;3,4}))"],
      ]);

      expect(column("A", 1, 5)).toEqual([1, 2, 3, 4, ""]);
      expect(cell("B1")).toBe(10);
    });

    test("reads computed results that refer to a range by their values", () => {
      setContent([
        ["C1", "=TOCOL(ROW(A2:A4),1)"],
        ["D1", "=TOCOL(COLUMN(A1:C1),1)"],
      ]);

      expect(column("C", 1, 4)).toEqual([2, 3, 4, ""]);
      expect(column("D", 1, 4)).toEqual([1, 2, 3, ""]);
    });

    test("reads cells spilled onto the range by formulas inside and above it", () => {
      setContent([
        ["A1", "x"],
        ["A3", "=SEQUENCE(3)"],
        ["C1", "=TOCOL(A1:A6,1)"],
        ["D1", "=TOCOL(A1:A6)"],
        ["E1", "=SEQUENCE(5)"],
        ["F1", "=TOCOL(E3:E6,1)"],
        ["G1", "=TOCOL(A:A,1)"],
        ["H1", "=UNIQUE(TOCOL(A:A,2))"],
        ["I1", "=TOCOL(A2:A,1)"],
      ]);

      expect(column("C", 1, 5)).toEqual(["x", 1, 2, 3, ""]);
      expect(column("D", 1, 7)).toEqual(["x", "", 1, 2, 3, "", ""]);
      expect(column("F", 1, 4)).toEqual([3, 4, 5, ""]);
      // Open-ended ranges, with the spill anchored below the range start.
      expect(column("G", 1, 5)).toEqual(["x", 1, 2, 3, ""]);
      expect(column("H", 1, 6)).toEqual(["x", "", 1, 2, 3, ""]);
      expect(column("I", 1, 4)).toEqual([1, 2, 3, ""]);
    });

    test("nests inside itself", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["D1", "=TOCOL(TOCOL(A1:B2,0,TRUE))"],
        ["E1", "=TOCOL(TOROW(A1:B2),1)"],
      ]);

      expect(column("D", 1, 5)).toEqual([1, 3, 2, 4, ""]);
      expect(column("E", 1, 5)).toEqual([1, 2, 3, 4, ""]);
    });

    test("waits for computed inputs whose cells are not evaluated yet", () => {
      setContent([
        ["B1", 1],
        ["B2", 2],
        ["B3", 3],
        ["A1", "=B1*2"],
        ["A2", "=B2*2"],
        ["A3", "=B3*2"],
        ["D1", "=TOCOL(-A1:A3,1)"],
        ["E1", "=SUM(TOCOL(-A1:A3))"],
      ]);

      expect(column("D", 1, 4)).toEqual([-2, -4, -6, ""]);
      expect(cell("E1")).toBe(-12);

      setCellContent("B2", 20);

      expect(column("D", 1, 4)).toEqual([-2, -40, -6, ""]);
      expect(cell("E1")).toBe(-48);
    });

    test("maps an unbounded SEQUENCE lazily but cannot skip its cells", () => {
      setContent([
        ["A1", "=TOCOL(SEQUENCE(INFINITY))"],
        ["B1", "=TOCOL(SEQUENCE(INFINITY),1)"],
      ]);

      expect(column("A", 1, 3)).toEqual([1, 2, 3]);
      expect(cell("A5000")).toBe(5000);
      // Finding the blanks would mean reading every one of its cells.
      expect(cell("B1")).toBe(FormulaError.REF);
    });
  });

  describe("consumers", () => {
    test("feeds UNIQUE", () => {
      setContent([
        ["A1", "header"],
        ["A2", "b"],
        ["A4", "a"],
        ["A5", "b"],
        ["A9", "c"],
        ["A10", "a"],
        ["C1", "=UNIQUE(TOCOL(A2:A1000,1))"],
      ]);

      expect(column("C", 1, 4)).toEqual(["b", "a", "c", ""]);
    });

    test("feeds SUM and COUNT", () => {
      setContent([
        ...gridWithBlanksAndErrors,
        ["D1", "=SUM(TOCOL(A1:B3,3))"],
        ["D2", "=COUNT(TOCOL(A1:B3,2))"],
        ["D3", "=COUNT(TOCOL(A1:B3))"],
        ["D4", "=SUM(TOCOL(SEQUENCE(2,2)))"],
        ["D5", "=COUNT(TOCOL(SEQUENCE(3,3),1,TRUE))"],
        ["D6", "=SUM(TOCOL(A:A,1))"],
        ["D7", "=COUNT(A1:B3)"],
      ]);

      expect(cell("D1")).toBe(5);
      expect(cell("D2")).toBe(2);
      // Kept errors reach COUNT, which propagates them as it does for the range.
      expect(cell("D3")).toBe(FormulaError.NAME);
      expect(cell("D7")).toBe(FormulaError.NAME);
      expect(cell("D4")).toBe(10);
      expect(cell("D5")).toBe(9);
      expect(cell("D6")).toBe(1);
    });

    test("does not apply an outer operation twice", () => {
      setContent([
        ["A1", "=SUM(TOCOL(SEQUENCE(2,2))*10)"],
        ["A2", "=SUM(-TOCOL(SEQUENCE(2,2),1))"],
      ]);

      expect(cell("A1")).toBe(100);
      expect(cell("A2")).toBe(-10);
    });

    test("is read through ranges that cover its spill", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A4", 4],
        ["C1", "=TOCOL(A1:A4,1)"],
        ["D1", "=SUM(C:C)"],
        ["D2", "=SUM(C2:C3)"],
        ["E1", "=TOCOL(A:A)"],
        ["F1", "=SUM(E:E)"],
      ]);

      expect(cell("D1")).toBe(7);
      expect(cell("D2")).toBe(6);
      expect(cell("F1")).toBe(7);
    });

    test("supports INDEX and MATCH", () => {
      setContent([
        ["A1", "a"],
        ["B1", "b"],
        ["A2", "c"],
        ["B2", "d"],
        ["D1", "=INDEX(TOCOL(A1:B2),3)"],
        ["D2", '=MATCH("d",TOCOL(A1:B2,0,TRUE),0)'],
      ]);

      expect(cell("D1")).toBe("c");
      expect(cell("D2")).toBe(4);
    });
  });

  describe("recalculation", () => {
    test("reflects an edited source cell", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["D1", "=TOCOL(A1:B2)"],
        ["E1", "=TOCOL(A1:B2,1)"],
      ]);

      expect(column("D", 1, 4)).toEqual([1, 2, 3, 4]);
      expect(column("E", 1, 4)).toEqual([1, 2, 3, 4]);

      setCellContent("B1", 20);

      expect(column("D", 1, 4)).toEqual([1, 20, 3, 4]);
      expect(column("E", 1, 4)).toEqual([1, 20, 3, 4]);
    });

    test("grows when a blank cell inside the range is filled", () => {
      setContent([
        ["A1", "a"],
        ["A3", "c"],
        ["C1", "=TOCOL(A1:A5,1)"],
        ["D1", "=TOCOL(A1:A5)"],
      ]);

      expect(column("C", 1, 3)).toEqual(["a", "c", ""]);
      expect(column("D", 1, 3)).toEqual(["a", "", "c"]);

      setCellContent("A2", "b");

      expect(column("C", 1, 4)).toEqual(["a", "b", "c", ""]);
      expect(column("D", 1, 3)).toEqual(["a", "b", "c"]);
    });

    test("picks up a value added below the last populated row of an open-ended range", () => {
      setContent([
        ["A2", "a"],
        ["A3", "b"],
        ["C1", "=TOCOL(A2:A,1)"],
        ["D1", "=UNIQUE(TOCOL(A:A,1))"],
      ]);

      expect(column("C", 1, 3)).toEqual(["a", "b", ""]);
      expect(column("D", 1, 3)).toEqual(["a", "b", ""]);

      setCellContent("A20", "c");

      expect(column("C", 1, 4)).toEqual(["a", "b", "c", ""]);
      expect(column("D", 1, 4)).toEqual(["a", "b", "c", ""]);
    });

    test("collapses and expands as the number of surviving cells changes", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["C1", "=TOCOL(A1:A3,1)"],
      ]);

      expect(column("C", 1, 3)).toEqual(["a", "b", ""]);

      setCellContent("A2", "");
      expect(column("C", 1, 2)).toEqual(["a", ""]);

      setCellContent("A1", "");
      expect(cell("C1")).toBe(FormulaError.CALC);

      setCellContent("A3", "z");
      expect(column("C", 1, 2)).toEqual(["z", ""]);
    });

    test("follows the shape of a spilled input", () => {
      setContent([
        ["A1", 2],
        ["B1", "=SEQUENCE(A1,2)"],
        ["E1", "=TOCOL(B1)"],
      ]);

      expect(column("E", 1, 5)).toEqual([1, 2, 3, 4, ""]);

      setCellContent("A1", 3);

      expect(column("E", 1, 7)).toEqual([1, 2, 3, 4, 5, 6, ""]);
    });

    test("survives a serialized engine round trip", () => {
      setContent([
        ["A1", 1],
        ["A2", "=NOSUCHFN()"],
        ["A4", 4],
        ["C1", "=TOCOL(A1:A4)"],
        ["D1", "=TOCOL(A:A,3)"],
        ["E1", "=SUM(TOCOL(A1:A4,2))"],
      ]);
      const refs = ["C1", "C2", "C3", "C4", "D1", "D2", "D3", "E1"];
      const before = refs.map(cell);
      expect(before).toEqual([1, FormulaError.NAME, "", 4, 1, 4, "", 5]);

      const restored = FormulaEngine.buildEmpty();
      restored.resetToSerializedEngine(engine.serializeEngine());
      engine = restored;
      expect(refs.map(cell)).toEqual(before);

      setCellContent("A6", 6);
      expect(column("D", 1, 4)).toEqual([1, 4, 6, ""]);
      expect(cell("E1")).toBe(5);
    });
  });

  describe("spilling", () => {
    test("returns #SPILL! when the result area is blocked", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["D1", "=TOCOL(A1:B1)"],
        ["D2", "blocker"],
      ]);

      expect(cell("D1")).toBe(FormulaError.SPILL);
    });
  });

  describe("arguments", () => {
    test("rejects the wrong number of arguments", () => {
      setContent([
        ["A1", 1],
        ["C1", "=TOCOL()"],
        ["C2", "=TOCOL(A1,0,FALSE,1)"],
      ]);

      expect(cell("C1")).toBe(FormulaError.VALUE);
      expect(cell("C2")).toBe(FormulaError.VALUE);
    });

    test("keeps the defaults of omitted options", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["A4", 5],
        ["D1", "=TOCOL(A1:B4,,TRUE)"],
        ["E1", "=TOCOL(A1:B2,,)"],
        ["F1", "=TOCOL(A1:B4,1,)"],
      ]);

      // Every cell is kept, column by column.
      expect(column("D", 1, 9)).toEqual([1, 3, "", 5, 2, 4, "", "", ""]);
      expect(column("E", 1, 5)).toEqual([1, 2, 3, 4, ""]);
      expect(column("F", 1, 6)).toEqual([1, 2, 3, 4, 5, ""]);
    });

    test("requires the array", () => {
      setContent([["A1", "=TOCOL(,1)"]]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
    });
  });

  describe("performance", () => {
    test("reads the last cell of a large flattened SEQUENCE without reading the rest", () => {
      const start = performance.now();
      setContent([
        ["A1", "=B90000"],
        ["B1", "=TOCOL(SEQUENCE(300,300))"],
        ["C1", "=D1000000"],
        ["D1", "=TOCOL(SEQUENCE(1000,1000),0,TRUE)"],
      ]);

      expect(cell("A1")).toBe(90000);
      expect(cell("C1")).toBe(1000000);
      expect(performance.now() - start).toBeLessThan(200);
    });

    test("reads a large input once when skipping", () => {
      const start = performance.now();
      setContent([
        ["A1", "=B90000"],
        ["B1", "=TOCOL(SEQUENCE(300,300),1,TRUE)"],
        ["C1", "=SUM(TOCOL(SEQUENCE(300,300),3))"],
      ]);

      // Column by column, the last cell is the bottom-right one.
      expect(cell("A1")).toBe(90000);
      expect(cell("C1")).toBe((90000 * 90001) / 2);
      expect(performance.now() - start).toBeLessThan(2000);
    });
  });
});
