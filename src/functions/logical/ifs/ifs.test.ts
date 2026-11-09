import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("IFS function", () => {
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

  describe("single-value conditions", () => {
    test("returns the value of the first true condition", () => {
      const grade =
        '=IFS(A1>89,"A",A1>79,"B",A1>69,"C",A1>59,"D",TRUE,"F")';
      setContent([
        ["A1", 93],
        ["B1", grade],
      ]);
      expect(cell("B1")).toBe("A");

      setCellContent("A1", 80);
      expect(cell("B1")).toBe("B");

      setCellContent("A1", 65);
      expect(cell("B1")).toBe("D");

      setCellContent("A1", 12);
      expect(cell("B1")).toBe("F");
    });

    test("returns #N/A when no condition is true", () => {
      setContent([
        ["A1", 5],
        ["B1", '=IFS(A1>10,"big",A1<0,"negative")'],
        ["B2", '=IFS(FALSE,1)'],
      ]);

      expect(cell("B1")).toBe(FormulaError.NA);
      expect(cell("B2")).toBe(FormulaError.NA);
    });

    test("tests conditions like IF", () => {
      setContent([
        ["A1", '=IFS(0,"zero",1,"one")'],
        ["A2", '=IFS(-2,"negative",TRUE,"other")'],
        ["A3", '=IFS("","empty text",TRUE,"other")'],
        ["A4", '=IFS("yes","text",TRUE,"other")'],
        // B1 is empty
        ["A5", '=IFS(B1,"blank",TRUE,"other")'],
        ["A6", '=IFS(INFINITY,"infinity",TRUE,"other")'],
      ]);

      expect(cell("A1")).toBe("one");
      expect(cell("A2")).toBe("negative");
      expect(cell("A3")).toBe("other");
      expect(cell("A4")).toBe("text");
      expect(cell("A5")).toBe("other");
      expect(cell("A6")).toBe("infinity");
    });

    test("stops at the first true condition", () => {
      // Neither the conditions after the first true one nor the values of other
      // pairs are evaluated.
      setContent([
        ["A1", "=IFS(TRUE, 1, NOSUCHFN(), 2)"],
        ["A2", "=IFS(FALSE, NOSUCHFN(), TRUE, 2)"],
        ["A3", "=IFS(TRUE, 1, TRUE, 1/\"x\")"],
      ]);

      expect(cell("A1")).toBe(1);
      expect(cell("A2")).toBe(2);
      expect(cell("A3")).toBe(1);
    });

    test("propagates errors from conditions it reaches and the chosen value", () => {
      setContent([
        ["A1", "=IFS(FALSE, 1, NOSUCHFN(), 2)"],
        ["A2", '=IFS(1+"x", 1, TRUE, 2)'],
        ["A3", "=IFS(FALSE, 1, TRUE, NOSUCHFN())"],
      ]);

      expect(cell("A1")).toBe(FormulaError.NAME);
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(cell("A3")).toBe(FormulaError.NAME);
    });

    test("recalculates when a referenced condition or value changes", () => {
      setContent([
        ["A1", false],
        ["A2", "first"],
        ["A3", "second"],
        ["B1", "=IFS(A1, A2, TRUE, A3)"],
      ]);
      expect(cell("B1")).toBe("second");

      setCellContent("A1", true);
      expect(cell("B1")).toBe("first");

      setCellContent("A2", "changed");
      expect(cell("B1")).toBe("changed");
    });

    test("returns an array value as is, so it spills", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["C1", "=IFS(FALSE, 0, TRUE, A1:A3)"],
        ["D1", "=IFS(TRUE, SEQUENCE(1,3))"],
      ]);

      expect(cell("C1")).toBe(1);
      expect(cell("C2")).toBe(2);
      expect(cell("C3")).toBe(3);
      expect(cell("D1")).toBe(1);
      expect(cell("E1")).toBe(2);
      expect(cell("F1")).toBe(3);
    });
  });

  describe("arguments", () => {
    test("requires pairs of arguments", () => {
      setContent([
        ["A1", "=IFS()"],
        ["A2", "=IFS(TRUE)"],
        ["A3", '=IFS(TRUE, 1, FALSE)'],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(cell("A3")).toBe(FormulaError.VALUE);
    });

    test("reads an omitted condition as FALSE and an omitted value as 0", () => {
      setContent([
        ["A1", '=IFS(, "skipped", TRUE, "chosen")'],
        ["A2", "=IFS(TRUE,)"],
        ["A3", "=IFS(FALSE, 1, TRUE,)"],
      ]);

      expect(cell("A1")).toBe("chosen");
      expect(cell("A2")).toBe(0);
      expect(cell("A3")).toBe(0);
    });
  });

  describe("array conditions", () => {
    test("chooses cell by cell and spills", () => {
      setContent([
        ["A1", 5],
        ["A2", 15],
        ["A3", 25],
        ["C1", '=IFS(A1:A3>20, "big", A1:A3>10, "medium", TRUE, "small")'],
      ]);

      expect(cell("C1")).toBe("small");
      expect(cell("C2")).toBe("medium");
      expect(cell("C3")).toBe("big");

      setCellContent("A1", 30);
      expect(cell("C1")).toBe("big");
    });

    test("gives #N/A in cells without a true condition", () => {
      setContent([
        ["A1", 5],
        ["A2", 15],
        ["C1", '=IFS(A1:A2>10, "big")'],
      ]);

      expect(cell("C1")).toBe(FormulaError.NA);
      expect(cell("C2")).toBe("big");
    });

    test("broadcasts conditions and values of different shapes", () => {
      // A row condition against a column value: a 2 x 3 grid.
      setContent([
        ["A1", "x"],
        ["A2", "y"],
        ["C1", '=IFS({TRUE,FALSE,TRUE}, A1:A2, TRUE, "-")'],
      ]);

      expect(cell("C1")).toBe("x");
      expect(cell("D1")).toBe("-");
      expect(cell("E1")).toBe("x");
      expect(cell("C2")).toBe("y");
      expect(cell("D2")).toBe("-");
      expect(cell("E2")).toBe("y");
    });

    test("reads array values per cell", () => {
      setContent([
        ["A1", 1],
        ["A2", -2],
        ["A3", 3],
        ["C1", "=IFS(A1:A3>0, A1:A3*10, TRUE, A1:A3*-1)"],
      ]);

      expect(cell("C1")).toBe(10);
      expect(cell("C2")).toBe(2);
      expect(cell("C3")).toBe(30);
    });

    test("shows an error only in the cells that reach it", () => {
      setContent([
        ["C1", '=IFS({TRUE;FALSE}, "ok", NOSUCHFN(), "never")'],
        ["D1", '=IFS({TRUE;FALSE}, "ok", TRUE, 1+"x")'],
      ]);

      expect(cell("C1")).toBe("ok");
      expect(cell("C2")).toBe(FormulaError.NAME);
      expect(cell("D1")).toBe("ok");
      expect(cell("D2")).toBe(FormulaError.VALUE);
    });

    test("switches to cell by cell at the first array condition", () => {
      // The single-value conditions before the array condition are false, so
      // they are skipped as a whole.
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", '=IFS(FALSE, "never", A1:A2=2, "two", TRUE, "other")'],
        ["D1", '=IFS(TRUE, "always", A1:A2=2, "two")'],
      ]);

      expect(cell("C1")).toBe("other");
      expect(cell("C2")).toBe("two");
      expect(cell("D1")).toBe("always");
    });

    test("is read whole by aggregates", () => {
      setContent([
        ["A1", 5],
        ["A2", 15],
        ["A3", 25],
        ["B1", "=SUM(IFS(A1:A3>10, A1:A3, TRUE, 0))"],
        ["B2", '=COUNTA(IFS(A1:A3>10, "x", TRUE, "y"))'],
      ]);

      expect(cell("B1")).toBe(40);
      expect(cell("B2")).toBe(3);
    });
  });
});
