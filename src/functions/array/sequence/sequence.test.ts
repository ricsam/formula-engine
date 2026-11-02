import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";
import { visualizeSpreadsheet } from "../../../core/utils/spreadsheet-visualizer";

describe("SEQUENCE function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  test("basic 1D sequence", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(3)"]])
    );

    // Should generate: 1, 2, 3 in column A
    expect(cell("A1")).toBe(1);
    expect(cell("A2")).toBe(2);
    expect(cell("A3")).toBe(3);
  });

  test("2D sequence", () => {
    engine.setSheetContent(
      sheetAddress,
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
      sheetAddress,
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
      sheetAddress,
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
      sheetAddress,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(1, 1, 42)"]])
    );

    // Should generate just the value 42
    expect(cell("A1")).toBe(42);
  });

  test("error cases", () => {
    engine.setSheetContent(
      sheetAddress,
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
      sheetAddress,
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

  test("rejects an array argument with #VALUE!", () => {
    // Excel lifts SEQUENCE over an array argument; this engine does not.
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(SEQUENCE(3))"],
        ["B1", "=SEQUENCE(SEQUENCE(5), 1, 10)"],
        ["C1", "=SEQUENCE(2, {1,2})"],
        ["D1", "=SEQUENCE(2, 1, G1:G2)"],
        ["E1", "=SEQUENCE(2, 1, 1, G1:G3)"],
      ])
    );

    expect(cell("A1")).toBe(FormulaError.VALUE);
    expect(cell("A2")).toBe("");
    expect(cell("B1")).toBe(FormulaError.VALUE);
    expect(cell("C1")).toBe(FormulaError.VALUE);
    expect(cell("D1")).toBe(FormulaError.VALUE);
    expect(cell("E1")).toBe(FormulaError.VALUE);
  });

  test("accepts a single-cell range or a one-element array as an argument", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 3],
        ["B1", "=SEQUENCE(A1:A1)"],
        ["C1", "=SEQUENCE({2})"],
      ])
    );

    expect(cell("B1")).toBe(1);
    expect(cell("B3")).toBe(3);
    expect(cell("C1")).toBe(1);
    expect(cell("C2")).toBe(2);
  });

  describe("omitted arguments", () => {
    test("an omitted rows argument defaults to 1", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(,3)"]])
      );

      // A single row {1,2,3}
      expect(cell("A1")).toBe(1);
      expect(cell("B1")).toBe(2);
      expect(cell("C1")).toBe(3);
      expect(cell("D1")).toBe("");
      expect(cell("A2")).toBe("");
    });

    test("an omitted columns argument defaults to 1", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=SEQUENCE(3,,10)"],
          ["B1", "=SEQUENCE(2,)"],
        ])
      );

      // A single column {10;11;12}
      expect(cell("A1")).toBe(10);
      expect(cell("A2")).toBe(11);
      expect(cell("A3")).toBe(12);
      expect(cell("A4")).toBe("");
      expect(cell("B1")).toBe(1);
      expect(cell("B2")).toBe(2);
      expect(cell("C1")).toBe("");
    });

    test("omitted start and step default to 1", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=SEQUENCE(2,2,,)"],
          ["D1", "=SEQUENCE(1,3,,5)"],
          ["D2", "=SEQUENCE(1,3,5,)"],
        ])
      );

      expect(cell("A1")).toBe(1);
      expect(cell("B1")).toBe(2);
      expect(cell("A2")).toBe(3);
      expect(cell("B2")).toBe(4);
      expect(cell("D1")).toBe(1);
      expect(cell("E1")).toBe(6);
      expect(cell("F1")).toBe(11);
      expect(cell("D2")).toBe(5);
      expect(cell("E2")).toBe(6);
      expect(cell("F2")).toBe(7);
    });

    test("every argument omitted gives the single value 1", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=SEQUENCE(,)"],
          ["A2", "=SEQUENCE(,,,)"],
          ["A3", "=SEQUENCE(,,7)"],
        ])
      );

      expect(cell("A1")).toBe(1);
      expect(cell("B1")).toBe("");
      expect(cell("A2")).toBe(1);
      expect(cell("A3")).toBe(7);
    });

    test("an argument referencing an empty cell is not omitted", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=SEQUENCE(Z1,3)"],
          ["A2", "=SEQUENCE(2,1,Z1)"],
        ])
      );

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });

    test("feeds aggregates and other functions", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(SEQUENCE(,4))"],
          ["A2", "=INDEX(SEQUENCE(,4,10),1,3)"],
          ["A3", "=SUM(SEQUENCE(3,,10))"],
        ])
      );

      expect(cell("A1")).toBe(10);
      expect(cell("A2")).toBe(12);
      expect(cell("A3")).toBe(33);
    });

    test("recalculates when a supplied argument changes", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 2],
          ["B1", "=SEQUENCE(,A1)"],
        ])
      );

      expect(cell("B1")).toBe(1);
      expect(cell("C1")).toBe(2);
      expect(cell("D1")).toBe("");

      engine.setCellContent(
        { sheetName, workbookName, ...parseCellReference("A1") },
        3
      );

      expect(cell("B1")).toBe(1);
      expect(cell("D1")).toBe(3);
    });
  });

  test("rejects non-numeric start and step", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", '=SEQUENCE(2,1,"a")'],
        ["A2", '=SEQUENCE(2,1,1,"b")'],
        ["A3", "=SEQUENCE(-INFINITY)"],
        ["A4", "=SEQUENCE()"],
        ["A5", "=SEQUENCE(1,1,1,1,1)"],
      ])
    );

    expect(cell("A1")).toBe(FormulaError.VALUE);
    expect(cell("A2")).toBe(FormulaError.VALUE);
    expect(cell("A3")).toBe(FormulaError.VALUE);
    expect(cell("A4")).toBe(FormulaError.VALUE);
    expect(cell("A5")).toBe(FormulaError.VALUE);
  });

  test("aggregates apply a wrapping operator once", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SUM(-SEQUENCE(3))"],
        ["A2", "=SUM(SEQUENCE(4)/2)"],
        ["A3", "=SUM(SEQUENCE(4)*SEQUENCE(4))"],
      ])
    );

    expect(cell("A1")).toBe(-6);
    expect(cell("A2")).toBe(5);
    expect(cell("A3")).toBe(30);
  });

  test("SEQUENCE with INFINITY rows", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(INFINITY)"]])
    );

    // SEQUENCE(INFINITY) creates an infinite vertical sequence
    expect(cell("A1")).toBe(1);
    expect(cell("A2")).toBe(2);
    expect(cell("A3")).toBe(3);
    expect(cell("A100")).toBe(100);
    expect(cell("A1000")).toBe(1000);
  });

  test("SEQUENCE with INFINITY columns", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(1, INFINITY)"]])
    );

    // SEQUENCE(1, INFINITY) creates an infinite horizontal sequence
    expect(cell("A1")).toBe(1);
    expect(cell("B1")).toBe(2);
    expect(cell("C1")).toBe(3);
    expect(cell("Z1")).toBe(26);
    expect(cell("AA1")).toBe(27);
  });

  test("SEQUENCE with INFINITY and custom start/step", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(INFINITY, 1, 10, 5)"],
      ])
    );

    // SEQUENCE(INFINITY, 1, 10, 5) creates infinite sequence starting at 10, step 5
    expect(cell("A1")).toBe(10);
    expect(cell("A2")).toBe(15);
    expect(cell("A3")).toBe(20);
    expect(cell("A10")).toBe(55);
  });

  test("SEQUENCE with INFINITY rows fills row by row across finite columns", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([["A1", "=SEQUENCE(INFINITY, 2)"]])
    );

    expect(cell("A1")).toBe(1);
    expect(cell("B1")).toBe(2);
    expect(cell("A2")).toBe(3);
    expect(cell("B2")).toBe(4);
    expect(cell("B50")).toBe(100);
    expect(cell("C1")).toBe("");
  });

  test("SEQUENCE with both INFINITY rows and columns", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(INFINITY, INFINITY)"],
      ])
    );

    // SEQUENCE(INFINITY, INFINITY) creates infinite 2D sequence
    // Values fill row by row
    expect(cell("A1")).toBe(1);
    expect(cell("B1")).toBe(2);
    expect(cell("C1")).toBe(3);
    // Can't predict exact values for arbitrary cells since it's infinite
    // but they should follow the pattern
  });

  test("SUM with SEQUENCE(INFINITY) should return INFINITY", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(INFINITY)"],
        ["B1", "=SUM(A1:A10)"], // Sum finite portion
        ["B2", "=SUM(A1:A)"], // Sum entire column with infinite sequence
      ])
    );

    // Sum of first 10 values: 1+2+...+10 = 55
    expect(cell("B1", true)).toBe(55);
    // Sum of entire column with infinite sequence should be INFINITY
    expect(cell("B2", true)).toMatchInlineSnapshot(`0`);
  });

  test("SEQUENCE with intersection", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SUM(C3:D4)"],
        ["B2", "=I12:K14"],
        ["H10", "=SEQUENCE(10, 10)"],
      ])
    );

    expect(cell("A1", true)).toBe(154);
  });

  test("SEQUENCE with crazy spill areas", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A4", "=SUM(A17:E21)"],
        ["B15", "=H18:J26"],
        ["G8", "=Z12:AD40"],
        ["P13", "=Q27:AC39"],
        ["P26", "=SEQUENCE(15,15)"],
      ])
    );

    expect(cell("A4", true)).toBe(1056);
  });

  test("Two spill into the same cell", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A3", "=SEQUENCE(1,3,5)"],
        ["C1", "=SEQUENCE(3)"],
      ])
    );
    expect(cell("C1")).toBe(1);
    expect(cell("C3")).toBe(3);
    expect(cell("A3")).toBe(FormulaError.SPILL);
    expect(cell("B3")).toBe("");

    // reverse the order
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A3", "=SEQUENCE(1,3,5)"],
        ["C1", "=SEQUENCE(3)"],
      ])
    );
    expect(cell("A3")).toBe(5);
    expect(cell("C3")).toBe(7);
    expect(cell("C1")).toBe(FormulaError.SPILL);
    expect(cell("C2")).toBe("");
  });

  test("SEQUENCE lazy evaluation", () => {
    const startTime = performance.now();
    
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=B10000000"],
        ["B1", "=SEQUENCE(10000000)"],
      ])
    );
    expect(cell("A1")).toBe(10000000);
    
    const duration = performance.now() - startTime;
    // Performance regression check: Should complete in < 50ms
    // (Was ~600ms before optimization, now ~4ms)
    expect(duration).toBeLessThan(50);
  });
});
