import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("VLOOKUP function", () => {
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

  /** Product IDs in A2:A5, names in B2:B5 and quantities in C2:C5. */
  const products: Array<[string, SerializedCellValue]> = [
    ["A1", "ID"],
    ["B1", "Name"],
    ["C1", "Qty"],
    ["A2", "P-101"],
    ["B2", "Pipette"],
    ["C2", 5],
    ["A3", "P-102"],
    ["B3", "Plate"],
    ["C3", 12],
    ["A4", "P-103"],
    ["B4", "Tube"],
    // C4 is empty
    ["A5", "P-104"],
    ["B5", "Rack"],
    ["C5", '=1+"x"'],
  ];

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("exact match", () => {
    test("returns the cell in the same row of the requested column", () => {
      setContent([
        ...products,
        ["E1", '=VLOOKUP("P-102", A2:C5, 2, FALSE)'],
        ["E2", '=VLOOKUP("P-102", A2:C5, 3, FALSE)'],
        ["E3", '=VLOOKUP("P-104", A2:C5, 1, FALSE)'],
        ["E4", '=VLOOKUP("P-101", A2:C5, 3, 0)'],
      ]);

      expect(cell("E1")).toBe("Plate");
      expect(cell("E2")).toBe(12);
      expect(cell("E3")).toBe("P-104");
      expect(cell("E4")).toBe(5);
    });

    test("ignores case, like MATCH", () => {
      setContent([...products, ["E1", '=VLOOKUP("p-103", A2:C5, 2, FALSE)']]);

      expect(cell("E1")).toBe("Tube");
    });

    test("returns the first of several matches", () => {
      setContent([
        ["A1", "x"],
        ["B1", 1],
        ["A2", "x"],
        ["B2", 2],
        ["D1", '=VLOOKUP("x", A1:B2, 2, FALSE)'],
      ]);

      expect(cell("D1")).toBe(1);
    });

    test("matches numbers only with numbers", () => {
      setContent([
        ["A1", "1"],
        ["B1", "text one"],
        ["A2", 1],
        ["B2", "number one"],
        ["D1", "=VLOOKUP(1, A1:B2, 2, FALSE)"],
        ["D2", '=VLOOKUP("1", A1:B2, 2, FALSE)'],
      ]);

      expect(cell("D1")).toBe("number one");
      expect(cell("D2")).toBe("text one");
    });

    test("returns #N/A when nothing matches", () => {
      setContent([...products, ["E1", '=VLOOKUP("P-999", A2:C5, 2, FALSE)']]);

      expect(cell("E1")).toBe(FormulaError.NA);
    });

    test("returns a blank cell as empty text and a chosen error", () => {
      setContent([
        ...products,
        ["E1", '=VLOOKUP("P-103", A2:C5, 3, FALSE)'],
        ["E2", '=VLOOKUP("P-104", A2:C5, 3, FALSE)'],
      ]);

      // Excel shows 0 for the blank cell
      expect(cell("E1")).toBe("");
      expect(cell("E2")).toBe(FormulaError.VALUE);
    });

    test("follows edits of the lookup column and the returned cell", () => {
      setContent([
        ...products,
        ["G1", "P-102"],
        ["E1", "=VLOOKUP(G1, A2:C5, 3, FALSE)"],
      ]);
      expect(cell("E1")).toBe(12);

      setCellContent("C3", 99);
      expect(cell("E1")).toBe(99);

      setCellContent("A3", "P-200");
      expect(cell("E1")).toBe(FormulaError.NA);

      setCellContent("G1", "P-200");
      expect(cell("E1")).toBe(99);

      setCellContent("A5", "P-300");
      setCellContent("G1", "P-300");
      expect(cell("E1")).toBe(FormulaError.VALUE);
    });

    test("searches open-ended tables", () => {
      setContent([...products, ["E1", '=VLOOKUP("P-103", A:C, 2, FALSE)']]);
      expect(cell("E1")).toBe("Tube");

      setCellContent("A9", "P-109");
      setCellContent("B9", "Flask");
      setCellContent("E2", '=VLOOKUP("P-109", A:C, 2, FALSE)');
      expect(cell("E2")).toBe("Flask");
    });

    test("searches computed tables and array literals", () => {
      setContent([
        ...products,
        ["E1", '=VLOOKUP("Tube", HSTACK(B2:B5, A2:A5), 2, FALSE)'],
        ["E2", '=VLOOKUP(2, {1,"one";2,"two"}, 2, FALSE)'],
        ["E3", '=VLOOKUP("x", "x", 1, FALSE)'],
      ]);

      expect(cell("E1")).toBe("P-103");
      expect(cell("E2")).toBe("two");
      expect(cell("E3")).toBe("x");
    });
  });

  describe("approximate match", () => {
    const bands: Array<[string, SerializedCellValue]> = [
      ["A1", 0],
      ["B1", "F"],
      ["A2", 60],
      ["B2", "D"],
      ["A3", 70],
      ["B3", "C"],
      ["A4", 80],
      ["B4", "B"],
      ["A5", 90],
      ["B5", "A"],
    ];

    test("finds the largest value less than or equal to lookup_value", () => {
      setContent([
        ...bands,
        ["D1", "=VLOOKUP(72, A1:B5, 2, TRUE)"],
        ["D2", "=VLOOKUP(90, A1:B5, 2, TRUE)"],
        ["D3", "=VLOOKUP(100, A1:B5, 2, TRUE)"],
        ["D4", "=VLOOKUP(59.9, A1:B5, 2, 1)"],
      ]);

      expect(cell("D1")).toBe("C");
      expect(cell("D2")).toBe("A");
      expect(cell("D3")).toBe("A");
      expect(cell("D4")).toBe("F");
    });

    test("is the default when range_lookup is left out", () => {
      setContent([...bands, ["D1", "=VLOOKUP(85, A1:B5, 2)"]]);

      expect(cell("D1")).toBe("B");
    });

    test("returns #N/A below the first value", () => {
      setContent([...bands, ["D1", "=VLOOKUP(-1, A1:B5, 2)"]]);

      expect(cell("D1")).toBe(FormulaError.NA);
    });

    test("reads an omitted or empty range_lookup as FALSE", () => {
      setContent([
        ...bands,
        ["D1", "=VLOOKUP(72, A1:B5, 2,)"],
        ["D2", "=VLOOKUP(70, A1:B5, 2,)"],
        // F1 is empty
        ["D3", "=VLOOKUP(72, A1:B5, 2, F1)"],
      ]);

      expect(cell("D1")).toBe(FormulaError.NA);
      expect(cell("D2")).toBe("C");
      expect(cell("D3")).toBe(FormulaError.NA);
    });
  });

  describe("arguments", () => {
    test("checks col_index_num whether or not the value is found", () => {
      setContent([
        ...products,
        ["E1", '=VLOOKUP("P-102", A2:C5, 4, FALSE)'],
        ["E2", '=VLOOKUP("P-102", A2:C5, 0, FALSE)'],
        ["E3", '=VLOOKUP("P-999", A2:C5, 4, FALSE)'],
        ["E4", '=VLOOKUP("P-102", A2:C5, -1, FALSE)'],
        ["E5", '=VLOOKUP("P-102", A2:C5,, FALSE)'],
      ]);

      expect(cell("E1")).toBe(FormulaError.REF);
      expect(cell("E2")).toBe(FormulaError.VALUE);
      expect(cell("E3")).toBe(FormulaError.REF);
      expect(cell("E4")).toBe(FormulaError.VALUE);
      expect(cell("E5")).toBe(FormulaError.VALUE);
    });

    test("truncates col_index_num and requires a number", () => {
      setContent([
        ...products,
        ["E1", '=VLOOKUP("P-102", A2:C5, 2.9, FALSE)'],
        ["E2", '=VLOOKUP("P-102", A2:C5, "2", FALSE)'],
        ["E3", '=VLOOKUP("P-102", A2:C5, TRUE, FALSE)'],
      ]);

      expect(cell("E1")).toBe("Plate");
      expect(cell("E2")).toBe(FormulaError.VALUE);
      expect(cell("E3")).toBe(FormulaError.VALUE);
    });

    test("requires a number or text lookup value, like MATCH", () => {
      setContent([
        ...products,
        ["E1", '=VLOOKUP(TRUE, A2:C5, 2, FALSE)'],
        ["E2", '=VLOOKUP("P-102", A2:C5, 2, "yes")'],
        ["E3", "=VLOOKUP(, {0,\"zero\"}, 2, FALSE)"],
      ]);

      expect(cell("E1")).toBe(FormulaError.VALUE);
      expect(cell("E2")).toBe(FormulaError.VALUE);
      // an omitted lookup_value is 0
      expect(cell("E3")).toBe("zero");
    });

    test("propagates errors in the arguments", () => {
      setContent([
        ...products,
        ["E1", "=VLOOKUP(NOSUCHFN(), A2:C5, 2, FALSE)"],
        ["E2", '=VLOOKUP("P-102", NOSUCHFN(), 2, FALSE)'],
        ["E3", '=VLOOKUP("P-102", A2:C5, 1+"x", FALSE)'],
      ]);

      expect(cell("E1")).toBe(FormulaError.NAME);
      expect(cell("E2")).toBe(FormulaError.NAME);
      expect(cell("E3")).toBe(FormulaError.VALUE);
    });

    test("requires 3 or 4 arguments", () => {
      setContent([
        ["E1", '=VLOOKUP("x", A1:B2)'],
        ["E2", '=VLOOKUP("x", A1:B2, 2, FALSE, 1)'],
        ["E3", '=VLOOKUP("x",, 2)'],
      ]);

      expect(cell("E1")).toBe(FormulaError.VALUE);
      expect(cell("E2")).toBe(FormulaError.VALUE);
      expect(cell("E3")).toBe(FormulaError.VALUE);
    });
  });

  describe("arrays", () => {
    test("looks up every element of an array lookup_value", () => {
      setContent([
        ...products,
        ["G1", "P-104"],
        ["G2", "P-999"],
        ["G3", "P-101"],
        ["E1", "=VLOOKUP(G1:G3, A2:C5, 2, FALSE)"],
      ]);

      expect(cell("E1")).toBe("Rack");
      expect(cell("E2")).toBe(FormulaError.NA);
      expect(cell("E3")).toBe("Pipette");

      setCellContent("B2", "Tip");
      expect(cell("E3")).toBe("Tip");
    });

    test("returns several columns for an array col_index_num", () => {
      setContent([
        ...products,
        ["E1", '=VLOOKUP("P-102", A2:C5, {1,2,3}, FALSE)'],
      ]);

      expect(cell("E1")).toBe("P-102");
      expect(cell("F1")).toBe("Plate");
      expect(cell("G1")).toBe(12);
    });

    test("is read whole by aggregates", () => {
      setContent([
        ...products,
        ["E1", '=SUM(VLOOKUP({"P-101";"P-102"}, A2:C5, 3, FALSE))'],
      ]);

      expect(cell("E1")).toBe(17);
    });
  });
});
