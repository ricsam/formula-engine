import { describe, expect, test } from "bun:test";

import { FormulaEngine } from "../../../src/core/engine";
import { FormulaError, type CellAddress } from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

describe("FormulaEngine.getCellErrorDetails", () => {
  const workbookName = "TestWorkbook";
  const sheetName = "TestSheet";

  const address = (reference: string): CellAddress => ({
    workbookName,
    sheetName,
    ...parseCellReference(reference),
  });

  const createEngine = () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
    return engine;
  };

  test("returns the originating cell for a propagated evaluation error", () => {
    const engine = createEngine();
    engine.setCellContent(address("B1"), "=1+]");
    engine.setCellContent(address("A1"), "=B1");

    const details = engine.getCellErrorDetails(address("A1"));

    expect(details).toEqual({
      code: FormulaError.ERROR,
      message: expect.stringContaining("ParseError: Unexpected token: ]"),
      failingCellAddress: address("B1"),
    });
  });

  test("returns the evaluated cell when its own formula cannot be parsed", () => {
    const engine = createEngine();
    engine.setCellContent(address("C3"), "=SUM(1,]");

    expect(engine.getCellErrorDetails(address("C3"))).toEqual({
      code: FormulaError.ERROR,
      message: expect.stringContaining("ParseError: Unexpected token: ]"),
      failingCellAddress: address("C3"),
    });
  });

  test("omits the failing cell for an error attributed to an expression node", () => {
    const engine = createEngine();
    engine.setCellContent(address("A1"), "=UNKNOWN_FUNCTION()");

    expect(engine.getCellErrorDetails(address("A1"))).toEqual({
      code: FormulaError.NAME,
      message: "Function UNKNOWN_FUNCTION not found",
    });
  });

  test("retains the deepest referenced cell when its expression node owns the error", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook("Book");
    engine.addSheet({ workbookName: "Book", sheetName: "Summary" });
    engine.setSheetContent(
      { workbookName: "Book", sheetName: "Summary" },
      new Map([
        ["A1", "=B1"],
        ["B1", "=C1"],
        ["C1", "=UNKNOWN_FUNCTION()"],
      ])
    );

    expect(
      engine.getCellErrorDetails({
        workbookName: "Book",
        sheetName: "Summary",
        colIndex: 0,
        rowIndex: 0,
      })
    ).toEqual({
      code: FormulaError.NAME,
      message: "Function UNKNOWN_FUNCTION not found",
      failingCellAddress: {
        workbookName: "Book",
        sheetName: "Summary",
        colIndex: 2,
        rowIndex: 0,
      },
    });
  });

  test("returns undefined for values and empty cells", () => {
    const engine = createEngine();
    engine.setCellContent(address("A1"), 42);

    expect(engine.getCellErrorDetails(address("A1"))).toBeUndefined();
    expect(engine.getCellErrorDetails(address("B1"))).toBeUndefined();
  });
});
