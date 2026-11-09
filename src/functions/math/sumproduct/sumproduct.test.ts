import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("SUMPRODUCT function", () => {
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

  /** Units in A1:A4, prices in B1:B4. */
  const orders: Array<[string, SerializedCellValue]> = [
    ["A1", 3],
    ["A2", 2],
    ["A3", 4],
    ["A4", 5],
    ["B1", 1.5],
    ["B2", 10],
    ["B3", 0.25],
    ["B4", 2],
  ];

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  test("adds the products of corresponding cells", () => {
    setContent([...orders, ["D1", "=SUMPRODUCT(A1:A4, B1:B4)"]]);

    // 4.5 + 20 + 1 + 10
    expect(cell("D1")).toBe(35.5);
  });

  test("multiplies any number of arrays", () => {
    setContent([
      ...orders,
      ["C1", 1],
      ["C2", 0],
      ["C3", 2],
      ["C4", 1],
      ["D1", "=SUMPRODUCT(A1:A4, B1:B4, C1:C4)"],
      ["D2", "=SUMPRODUCT(A1:B2, {1,2;3,4})"],
    ]);

    // 4.5 + 0 + 2 + 10
    expect(cell("D1")).toBe(16.5);
    // 3*1 + 1.5*2 + 2*3 + 10*4
    expect(cell("D2")).toBe(52);
  });

  test("sums a single array", () => {
    setContent([
      ...orders,
      ["D1", "=SUMPRODUCT(A1:A4)"],
      ["D2", "=SUMPRODUCT(SEQUENCE(4))"],
      ["D3", "=SUMPRODUCT(5)"],
      ["D4", "=SUMPRODUCT(2, 3)"],
    ]);

    expect(cell("D1")).toBe(14);
    expect(cell("D2")).toBe(10);
    expect(cell("D3")).toBe(5);
    expect(cell("D4")).toBe(6);
  });

  test("recalculates when a cell changes", () => {
    setContent([...orders, ["D1", "=SUMPRODUCT(A1:A4, B1:B4)"]]);
    expect(cell("D1")).toBe(35.5);

    setCellContent("A2", 3);
    expect(cell("D1")).toBe(45.5);
  });

  test("multiplies computed arrays", () => {
    setContent([
      ...orders,
      ["D1", "=SUMPRODUCT(A1:A4*2, B1:B4)"],
      ["D2", "=SUMPRODUCT(A1:A4*B1:B4)"],
      ["D3", "=SUMPRODUCT(SEQUENCE(3), SEQUENCE(3))"],
    ]);

    expect(cell("D1")).toBe(71);
    expect(cell("D2")).toBe(35.5);
    expect(cell("D3")).toBe(14);
  });

  test("treats cells that are not numbers as 0", () => {
    setContent([
      ["A1", 2],
      ["A2", "text"],
      ["A3", true],
      // A4 is empty
      ["A5", 3],
      ["B1", 10],
      ["B2", 10],
      ["B3", 10],
      ["B4", 10],
      ["B5", 10],
      ["D1", "=SUMPRODUCT(A1:A5, B1:B5)"],
      ["D2", '=SUMPRODUCT({1,"2",TRUE})'],
    ]);

    expect(cell("D1")).toBe(50);
    expect(cell("D2")).toBe(1);
  });

  test("counts conditions through IF, since booleans are not numbers", () => {
    setContent([
      ...orders,
      ["C1", "x"],
      ["C2", "y"],
      ["C3", "x"],
      ["C4", "x"],
      // As in Excel, a boolean array contributes nothing by itself
      ["D1", '=SUMPRODUCT(C1:C4="x", A1:A4)'],
      ["D2", '=SUMPRODUCT(IF(C1:C4="x", A1:A4, 0), B1:B4)'],
      ["D3", '=SUMPRODUCT(IF(C1:C4="x", 1, 0))'],
      // Excel's --(...) idiom gives #VALUE! here
      ["D4", '=SUMPRODUCT(--(C1:C4="x"), A1:A4)'],
    ]);

    expect(cell("D1")).toBe(0);
    // 4.5 + 1 + 10
    expect(cell("D2")).toBe(15.5);
    expect(cell("D3")).toBe(3);
    expect(cell("D4")).toBe(FormulaError.VALUE);
  });

  test("returns #VALUE! for arrays of different sizes", () => {
    setContent([
      ...orders,
      ["D1", "=SUMPRODUCT(A1:A4, B1:B3)"],
      ["D2", "=SUMPRODUCT(A1:A4, TRANSPOSE(B1:B4))"],
      ["D3", "=SUMPRODUCT(A1:A4, 2)"],
      ["D4", "=SUMPRODUCT(A:A, B1:B4)"],
    ]);

    expect(cell("D1")).toBe(FormulaError.VALUE);
    expect(cell("D2")).toBe(FormulaError.VALUE);
    expect(cell("D3")).toBe(FormulaError.VALUE);
    expect(cell("D4")).toBe(FormulaError.VALUE);
  });

  test("returns an error anywhere in the arrays", () => {
    setContent([
      ...orders,
      ["C1", 1],
      ["C2", '=1+"x"'],
      ["C3", 1],
      ["C4", "=NOSUCHFN()"],
      ["D1", "=SUMPRODUCT(A1:A4, C1:C4)"],
      ["D2", "=SUMPRODUCT(C1:C4, NOSUCHFN())"],
      ["D3", "=SUMPRODUCT({1,2}, {0,1/\"x\"})"],
    ]);

    expect(cell("D1")).toBe(FormulaError.VALUE);
    expect(cell("D2")).toBe(FormulaError.NAME);
    expect(cell("D3")).toBe(FormulaError.VALUE);
  });

  test("multiplies and adds infinities", () => {
    setContent([
      ["D1", "=SUMPRODUCT({1,2}, {INFINITY,3})"],
      ["D2", "=SUMPRODUCT({0,2}, {INFINITY,3})"],
      ["D3", "=SUMPRODUCT({1,1}, {INFINITY,-INFINITY})"],
    ]);

    expect(cell("D1")).toBe("INFINITY");
    expect(cell("D2")).toBe(FormulaError.NUM);
    expect(cell("D3")).toBe(FormulaError.NUM);
  });

  test("works over open-ended ranges", () => {
    setContent([...orders, ["D1", "=SUMPRODUCT(A:A, B:B)"]]);
    expect(cell("D1")).toBe(35.5);

    setCellContent("A9", 2);
    setCellContent("B9", 100);
    expect(cell("D1")).toBe(235.5);
  });

  test("rejects omitted and missing arguments", () => {
    setContent([
      ["D1", "=SUMPRODUCT()"],
      ["D2", "=SUMPRODUCT(A1:A2,,B1:B2)"],
    ]);

    expect(cell("D1")).toBe(FormulaError.VALUE);
    expect(cell("D2")).toBe(FormulaError.VALUE);
  });
});
