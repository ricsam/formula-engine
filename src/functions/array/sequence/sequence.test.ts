import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("SEQUENCE function", () => {
  const sheetName = "TestSheet";
  let engine: FormulaEngine;

  const cell = (ref: string) =>
    engine.getCellValue({ sheetName, ...parseCellReference(ref) });

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent({ sheetName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addSheet(sheetName);
  });

  test("basic 1D sequence", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(3)"]])
    );

    // Should generate: 1, 2, 3 in column A
    expect(cell("A1")).toBe(1);
    expect(cell("A2")).toBe(2);
    expect(cell("A3")).toBe(3);
  });

  test("2D sequence", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(2,3)"]])
    );

    // Should generate:
    // A1: 1, B1: 2, C1: 3
    // A2: 4, B2: 5, C2: 6
    expect(cell("A1")).toBe(1);
    expect(cell("B1")).toBe(2);
    expect(cell("C1")).toBe(3);
    expect(cell("A2")).toBe(4);
    expect(cell("B2")).toBe(5);
    expect(cell("C2")).toBe(6);
  });

  test("custom start and step", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(3, 2, 10, 5)"]])
    );

    // Should generate with start=10, step=5:
    // A1: 10, B1: 15
    // A2: 20, B2: 25
    // A3: 30, B3: 35
    expect(cell("A1")).toBe(10);
    expect(cell("B1")).toBe(15);
    expect(cell("A2")).toBe(20);
    expect(cell("B2")).toBe(25);
    expect(cell("A3")).toBe(30);
    expect(cell("B3")).toBe(35);
  });

  test("negative step", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(4, 1, 10, -2)"]])
    );

    // Should generate with start=10, step=-2:
    // A1: 10, A2: 8, A3: 6, A4: 4
    expect(cell("A1")).toBe(10);
    expect(cell("A2")).toBe(8);
    expect(cell("A3")).toBe(6);
    expect(cell("A4")).toBe(4);
  });

  test("single cell", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(1, 1, 42)"]])
    );

    // Should generate just the value 42
    expect(cell("A1")).toBe(42);
  });

  test("error cases", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(0)"], // Invalid: rows must be > 0
        ["A2", "=SEQUENCE(2, 0)"], // Invalid: columns must be > 0
        ["A3", '=SEQUENCE("text")'], // Invalid: rows must be a number
      ])
    );

    expect(cell("A1")).toBe(FormulaError.VALUE);
    expect(cell("A2")).toBe(FormulaError.VALUE);
    expect(cell("A3")).toBe(FormulaError.VALUE);
  });

  test("with computed arguments", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 2],
        ["B1", 3],
        ["C1", 100],
        ["D1", 10],
        ["E1", "=SEQUENCE(A1, B1, C1, D1)"],
      ])
    );

    // Should generate 2x3 sequence starting at 100 with step 10:
    // E1: 100, F1: 110, G1: 120
    // E2: 130, F2: 140, G2: 150
    expect(cell("E1")).toBe(100);
    expect(cell("F1")).toBe(110);
    expect(cell("G1")).toBe(120);
    expect(cell("E2")).toBe(130);
    expect(cell("F2")).toBe(140);
    expect(cell("G2")).toBe(150);
  });

  test("array input", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(SEQUENCE(3))"]])
    );

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    // SEQUENCE(3) produces {1; 2; 3}
    // SEQUENCE(SEQUENCE(3)) uses origin value 1, so it's SEQUENCE(1) which is {1}
    // But since the input is an array with 3 rows, it broadcasts {1} over 3 rows
    expect(cell("A1")).toBe(1);
    expect(cell("A2")).toBe(1);
    expect(cell("A3")).toBe(1);
  });

  test("array input - SEQUENCE(SEQUENCE(5), 1, 10)", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(SEQUENCE(5), 1, 10)"],
      ])
    );

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    // SEQUENCE(5) produces {1; 2; 3; 4; 5}
    // SEQUENCE(SEQUENCE(5), 1, 10) uses origin value 1, so it's SEQUENCE(1, 1, 10) which is {10}
    // But since the input is an array with 5 rows, it broadcasts {10} over 5 rows
    expect(cell("A1")).toBe(10);
    expect(cell("A2")).toBe(10);
    expect(cell("A3")).toBe(10);
    expect(cell("A4")).toBe(10);
    expect(cell("A5")).toBe(10);
  });
});
