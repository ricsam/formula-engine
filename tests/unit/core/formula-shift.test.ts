import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { shiftFormulaReferences } from "../../../src/core/formula-shift";

describe("shiftFormulaReferences", () => {
  test("moves relative coordinates and keeps absolute ones", () => {
    expect(shiftFormulaReferences("=A5*$B$5+C$5+$D5", -3, 1)).toBe(
      "=B2*$B$5+D$5+$D2"
    );
    expect(shiftFormulaReferences("=SUM(A1:B3)", 2, 0)).toBe("=SUM(A3:B5)");
    expect(shiftFormulaReferences("=Sheet2!A5", -1, 0)).toBe("=Sheet2!A4");
  });

  test("leaves structured references, names and zero offsets alone", () => {
    expect(shiftFormulaReferences("=[@Price]*Rate", 4, 0)).toBe(
      "=[@Price]*Rate"
    );
    expect(shiftFormulaReferences("=A1 + 1", 0, 0)).toBe("=A1 + 1");
    expect(shiftFormulaReferences("not a formula", 1, 1)).toBe("not a formula");
  });

  test("turns references moved off the sheet into #REF!", () => {
    expect(shiftFormulaReferences("=A1+1", -1, 0)).toBe("=#REF!+1");
    expect(shiftFormulaReferences("=SUM(A1:A3)", 0, -1)).toBe("=SUM(#REF!)");
    expect(shiftFormulaReferences("=Sheet2!A2", -2, 0)).toBe("=#REF!");
  });

  test("keeps whole columns and rows anchored when asked", () => {
    expect(
      shiftFormulaReferences("=COUNTIF(A:A, A5)", -3, 0, {
        preserveWholeRowsAndColumns: true,
      })
    ).toBe("=COUNTIF(A1:A,A2)");
    // Copy semantics move the first row of an open-ended range.
    expect(shiftFormulaReferences("=COUNTIF(A:A, A5)", 2, 0)).toBe(
      "=COUNTIF(A3:A,A7)"
    );
  });

  test("pasting a formula above its references writes #REF!", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: "Book" });
    engine.addSheet({ workbookName: "Book", sheetName: "Sheet1" });
    const cell = (rowIndex: number, colIndex: number) => ({
      workbookName: "Book",
      sheetName: "Sheet1",
      rowIndex,
      colIndex,
    });
    engine.setCellContent(cell(4, 1), "=A1+1");
    engine.pasteCells([cell(4, 1)], cell(1, 1), {
      type: "formula",
      include: ["content"],
    });
    expect(
      engine.getSheet({ workbookName: "Book", sheetName: "Sheet1" })?.content.get("B2")
    ).toBe("=#REF!+1");
    expect(engine.getCellValue(cell(1, 1))).toBe("#REF!");
  });
});
