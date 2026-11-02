import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import {
  FormulaError,
  type SerializedCellValue,
  type SpreadsheetRange,
} from "../../../core/types";
import { parseCellReference } from "../../../core/utils";
import { MAX_LET_EXPANSION_NODES } from "./let-expansion";

/** A LET whose names each use the previous name twice: it expands to 2^n - 1 nodes. */
function doublingChain(pairs: number, calculation = `n_${pairs - 1}`): string {
  const parts = ["n_0", "A1"];
  for (let index = 1; index < pairs; index++) {
    parts.push(`n_${index}`, `n_${index - 1}+n_${index - 1}`);
  }
  return `=LET(${parts.join(", ")}, ${calculation})`;
}

describe("LET function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const address = (ref: string) => ({
    workbookName,
    sheetName,
    ...parseCellReference(ref),
  });

  const cell = (ref: string, target = engine) => target.getCellValue(address(ref));

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(address(ref), content);
  };

  const setContent = (content: Array<[string, SerializedCellValue]>) => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>(content)
    );
  };

  const formulaIn = (ref: string) =>
    engine.getSheetSerialized(sheetAddress).get(ref);

  const errorIn = (ref: string) => engine.getCellErrorDetails(address(ref));

  const range = (from: string, to: string): SpreadsheetRange => {
    const start = parseCellReference(from);
    const end = parseCellReference(to);
    return {
      start: { col: start.colIndex, row: start.rowIndex },
      end: {
        col: { type: "number", value: end.colIndex },
        row: { type: "number", value: end.rowIndex },
      },
    };
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("names and scope", () => {
    test("binds a name and returns the calculation", () => {
      setCellContent("A1", "=LET(x, 5, SUM(x, 1))");
      expect(cell("A1")).toBe(6);
    });

    test("does not share cached results between cells that bind different references", () => {
      // Regression: `x*2` must not be cached as one expression for both cells.
      setContent([
        ["A1", 1],
        ["A2", 10],
        ["B1", "=LET(x, A1, x*2)"],
        ["B2", "=LET(x, A2, x*2)"],
      ]);

      expect(cell("B1")).toBe(2);
      expect(cell("B2")).toBe(20);
    });

    test("binds names in order so a value can use earlier names", () => {
      setContent([
        ["A1", 4],
        ["B1", "=LET(a, A1, b, a*2, c, b+a, a & \"-\" & b & \"-\" & c)"],
        ["B2", "=LET(first, 5, second, SUM(first, 5), third, SUM(second, 5), SUM(second, third))"],
      ]);

      expect(cell("B1")).toBe("4-8-12");
      expect(cell("B2")).toBe(25);
    });

    test("matches names without regard to case", () => {
      setCellContent("A1", "=LET(Total, 3, TOTAL*2 + total)");
      expect(cell("A1")).toBe(9);
    });

    test("nested LETs see outer names and may redeclare them", () => {
      setContent([
        ["A1", "=LET(x, 3, LET(y, 4, x + y))"],
        ["A2", "=LET(x, 1, LET(x, x+1, x*10))"],
        ["A3", "=LET(x, 5, LET(x, 7, x*2) + x)"],
        ["A4", "=LET(x, 5, LET(y, 1, x, x+2, x*9) + x)"],
        ["A5", "=LET(a, 3, LET(b, 30, LET(c, 300, a + b + c)))"],
      ]);

      expect(cell("A1")).toBe(7);
      expect(cell("A2")).toBe(20);
      expect(cell("A3")).toBe(19);
      expect(cell("A4")).toBe(68);
      expect(cell("A5")).toBe(333);
    });

    test("a value does not see its own name", () => {
      setContent([
        ["A1", "=LET(x, 1, LET(x, x+1, x))"],
        // There is no defined name `x`, so the value's `x` is undefined.
        ["A2", "=LET(x, x+1, x)"],
      ]);

      expect(cell("A1")).toBe(2);
      expect(cell("A2")).toBe(FormulaError.NAME);
    });

    test("names cannot be used before they are declared", () => {
      setCellContent("A1", "=LET(a, b, b, 1, a)");
      expect(cell("A1")).toBe(FormulaError.NAME);
    });

    test("names are local to the LET", () => {
      setCellContent("A1", "=LET(x, 2, x) + x");
      expect(cell("A1")).toBe(FormulaError.NAME);
    });

    test("a bound value is not captured by a nested name of the same spelling", () => {
      engine.addNamedExpression({ expressionName: "y", expression: "100" });
      setContent([
        ["A1", "=LET(x, y, LET(y, 2, x+y))"],
        ["A2", "=LET(x, y*2, y, 1, x+y)"],
      ]);

      // x is bound to the defined name y (100), not to the inner y.
      expect(cell("A1")).toBe(102);
      expect(cell("A2")).toBe(201);
    });

    test("LET names hide defined names of the same name", () => {
      engine.addNamedExpression({ expressionName: "Rate", expression: "0.5" });
      setContent([
        ["A1", "=Rate*10"],
        ["A2", "=LET(Rate, 2, Rate*10)"],
        ["A3", "=LET(rate, 3, RATE*10)"],
        ["A4", "=LET(other, 3, Rate*10)"],
      ]);

      expect(cell("A1")).toBe(5);
      expect(cell("A2")).toBe(20);
      expect(cell("A3")).toBe(30);
      expect(cell("A4")).toBe(5);
    });

    test("names may share a function's name", () => {
      setCellContent("A1", "=LET(SQRT, 4, SUM(SQRT, SQRT))");
      expect(cell("A1")).toBe(8);
    });
  });

  describe("references", () => {
    test("a name bound to a reference stays a reference", () => {
      setContent([
        ["A1", 1],
        ["A2", 10],
        ["A3", 100],
        ["B1", "=LET(c, A2, ROW(c))"],
        ["B2", "=LET(r, A1:A3, SUM(r))"],
        ["B3", "=LET(r, A1:A3, ROW(r))"],
        ["B6", "=LET(c, C4, COLUMN(c))"],
      ]);

      expect(cell("B1")).toBe(2);
      expect(cell("B2")).toBe(111);
      // ROW of a bound range spills the range's row numbers.
      expect([cell("B3"), cell("B4"), cell("B5")]).toEqual([1, 2, 3]);
      expect(cell("B6")).toBe(3);
    });

    test("ROW() and COLUMN() in a value use the cell holding the formula", () => {
      setContent([
        ["E1", "=LET(r, ROW(), r*100)"],
        ["E2", "=LET(r, ROW(), r*100)"],
        ["F5", "=LET(c, COLUMN(), r, ROW(), c & \",\" & r)"],
        ["G5", "=LET(c, COLUMN(), r, ROW(), c & \",\" & r)"],
      ]);

      expect(cell("E1")).toBe(100);
      expect(cell("E2")).toBe(200);
      expect(cell("F5")).toBe("6,5");
      expect(cell("G5")).toBe("7,5");
    });

    test("recalculates after a referenced cell changes", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", "=LET(x, A1, x*2)"],
        ["B2", "=LET(r, A1:A3, total, SUM(r), total / COUNT(r))"],
      ]);
      expect(cell("B1")).toBe(2);
      expect(cell("B2")).toBe(2);

      setCellContent("A1", 7);
      expect(cell("B1")).toBe(14);
      expect(cell("B2")).toBe(4);

      setCellContent("A3", 6);
      expect(cell("B2")).toBe(5);
    });
  });

  describe("spills", () => {
    test("a bound spill can be used, transformed and aggregated", () => {
      setContent([
        ["A1", "=LET(s, SEQUENCE(3), s*2)"],
        ["B1", "=LET(s, SEQUENCE(4), SUM(s))"],
        ["C1", "=LET(n, 3, SEQUENCE(n, 2))"],
        ["E1", "=LET(s, SEQUENCE(2), s)"],
      ]);

      expect([cell("A1"), cell("A2"), cell("A3")]).toEqual([2, 4, 6]);
      expect(cell("A4")).toBe("");
      expect(cell("B1")).toBe(10);
      expect([cell("C1"), cell("D1"), cell("C3"), cell("D3")]).toEqual([
        1, 2, 5, 6,
      ]);
      expect([cell("E1"), cell("E2")]).toEqual([1, 2]);
    });

    test("array literals can use names", () => {
      setCellContent("A1", "=LET(x, 5, {x, x+1; x*2, x*3})");

      expect([cell("A1"), cell("B1"), cell("A2"), cell("B2")]).toEqual([
        5, 6, 10, 15,
      ]);
    });

    test("works with UNIQUE and recalculates its spill", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["A4", "c"],
        ["A5", "b"],
        ["C1", "=LET(u, UNIQUE(A1:A5), u)"],
        ["D1", "=LET(data, A1:A5, UNIQUE(data))"],
        ["E1", "=LET(n, SEQUENCE(3), UNIQUE(n))"],
      ]);

      expect([cell("C1"), cell("C2"), cell("C3"), cell("C4")]).toEqual([
        "a",
        "b",
        "c",
        "",
      ]);
      expect([cell("D1"), cell("D2"), cell("D3")]).toEqual(["a", "b", "c"]);
      expect([cell("E1"), cell("E2"), cell("E3")]).toEqual([1, 2, 3]);

      setCellContent("A3", "d");
      expect([cell("C1"), cell("C2"), cell("C3"), cell("C4")]).toEqual([
        "a",
        "b",
        "d",
        "c",
      ]);
    });
  });

  describe("named expressions", () => {
    test("a named expression can contain a LET", () => {
      engine.addNamedExpression({
        expressionName: "Tripled",
        expression: "LET(a, 2, a*3)",
      });
      setContent([
        ["A1", "=Tripled"],
        // The named expression's `a` is its own; the cell's `a` is not visible to it.
        ["A2", "=LET(a, 10, Tripled + a)"],
      ]);

      expect(cell("A1")).toBe(6);
      expect(cell("A2")).toBe(16);
    });

    test("a LET can use a named expression", () => {
      engine.addNamedExpression({ expressionName: "Tax", expression: "0.25" });
      setCellContent("A1", "=LET(price, 80, price * (1 + Tax))");
      expect(cell("A1")).toBe(100);

      engine.updateNamedExpression({ expressionName: "Tax", expression: "0.5" });
      expect(cell("A1")).toBe(120);
    });

    test("a named expression does not see the caller's LET names", () => {
      engine.addNamedExpression({ expressionName: "y", expression: "100" });
      engine.addNamedExpression({ expressionName: "Doubled", expression: "y*2" });
      setCellContent("A1", "=LET(y, 5, Doubled + y)");
      expect(cell("A1")).toBe(205);
    });
  });

  describe("unused names", () => {
    test("an error in an unused value does not matter", () => {
      setCellContent("A1", "=LET(unused, CEILING(5), 5)");
      expect(cell("A1")).toBe(5);
    });

    test("an unused value is never evaluated", () => {
      // Evaluating these values would make each cell depend on itself.
      setContent([
        ["A1", "=LET(unused, A1, 5)"],
        ["A2", "=LET(x, A2, LET(x, 1, x))"],
      ]);
      expect(cell("A1")).toBe(5);
      expect(cell("A2")).toBe(1);
    });

    test("an error in a used value propagates", () => {
      setCellContent("A1", "=LET(x, CEILING(5), x+1)");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("IF only evaluates the branch it takes", () => {
      setContent([
        ["A1", 0],
        ["B1", "=LET(x, CEILING(5), IF(A1 > 0, x, \"skipped\"))"],
      ]);
      expect(cell("B1")).toBe("skipped");

      setCellContent("A1", 1);
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });
  });

  describe("invalid calls", () => {
    test.each([
      ["=LET()"],
      ["=LET(x)"],
      ["=LET(x, 5)"],
      ["=LET(x, 1, y, 2)"],
      ["=LET(x, 1, )"],
    ])("%s is #VALUE!", (formula) => {
      setCellContent("A1", formula);
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test.each([
      ["=LET(x1, 5, x1)", "x1 is a cell reference"],
      ["=LET(TestSheet!x, 5, 1)", "qualified"],
      ["=LET(1, 5, 1)", "not a name"],
      ['=LET("x", 5, 1)', "not a name"],
      ["=LET(, 5, 1)", "missing"],
      ["=LET(x, 1, X, 2, x)", "repeated name"],
      ["=LET(a, 1, LET(x, 1, x, 2, x))", "repeated name in a nested LET"],
    ])("%s is #NAME? (%s)", (formula) => {
      setCellContent("A1", formula);
      expect(cell("A1")).toBe(FormulaError.NAME);
    });

    test("explains what is wrong", () => {
      setContent([
        ["A1", "=LET(x1, 5, x1)"],
        ["A2", "=LET(x, 1, X, 2, x)"],
        ["A3", "=LET(x, 5)"],
      ]);

      expect(errorIn("A1")?.message).toContain("X1");
      expect(errorIn("A1")?.message).toContain("cell reference");
      expect(errorIn("A2")?.message).toContain("more than once");
      expect(errorIn("A3")?.message).toContain("odd number");
    });

    test("an omitted value is 0", () => {
      setContent([
        ["A1", "=LET(a, , b, , c, , a + b + c)"],
        ["A2", "=LET(x, 20, y, , x + y)"],
        ["A3", "=LET(x, , x)"],
      ]);

      expect(cell("A1")).toBe(0);
      expect(cell("A2")).toBe(20);
      expect(cell("A3")).toBe(0);
    });

    test("IFERROR catches an invalid LET", () => {
      setCellContent("A1", '=IFERROR(LET(x, 1, x, 2, x), "invalid")');
      expect(cell("A1")).toBe("invalid");
    });
  });

  describe("expansion size limit", () => {
    test("allows expansions up to the limit", () => {
      // 13 doubling pairs expand to 2^13 - 1 = 8,191 nodes.
      expect(2 ** 13 - 1).toBeLessThanOrEqual(MAX_LET_EXPANSION_NODES);
      setContent([
        ["A1", 1],
        ["B1", doublingChain(13)],
      ]);

      expect(cell("B1")).toBe(4096);
      setCellContent("A1", 2);
      expect(cell("B1")).toBe(8192);
    });

    test("returns #CALC! past the limit", () => {
      // 14 pairs expand to 16,383 nodes; Excel's 126 pairs to about 2^126.
      expect(2 ** 14 - 1).toBeGreaterThan(MAX_LET_EXPANSION_NODES);
      setContent([
        ["A1", 1],
        ["B1", doublingChain(14)],
        ["B2", doublingChain(126)],
      ]);

      expect(cell("B1")).toBe(FormulaError.CALC);
      expect(errorIn("B1")?.message).toContain("16,383");
      expect(cell("B2")).toBe(FormulaError.CALC);
      expect(errorIn("B2")?.message).toContain("limit");
    });

    test("only counts the names that are used", () => {
      setContent([
        ["A1", 3],
        ["B1", doublingChain(126, "n_0 + n_2")],
      ]);
      // n_0 + n_2 = A1 + 4 * A1
      expect(cell("B1")).toBe(15);
    });
  });

  describe("copying and persistence", () => {
    test("autofill adjusts references in values but never names", () => {
      setContent([
        ["A1", 1],
        ["A2", 10],
        ["A3", 100],
        ["B1", "=LET(x, A1, ab, $A$1, x*2 + ab)"],
      ]);

      engine.autoFill(sheetAddress, range("B1", "B1"), [range("B2", "B3")], "down");

      expect(formulaIn("B2")).toBe("=LET(x, A2, ab, $A$1, x*2+ab)");
      expect(formulaIn("B3")).toBe("=LET(x, A3, ab, $A$1, x*2+ab)");
      expect([cell("B1"), cell("B2"), cell("B3")]).toEqual([3, 21, 201]);
    });

    test("pasting adjusts references in values but never names", () => {
      setContent([
        ["A1", 1],
        ["C5", 7],
        ["B1", "=LET(x, A1, x*2)"],
      ]);

      engine.pasteCells([address("B1")], address("D5"), {
        include: ["content"],
        type: "formula",
      });

      expect(formulaIn("D5")).toBe("=LET(x, C5, x*2)");
      expect(cell("D5")).toBe(14);
    });

    test("survives serializeEngine and resetToSerializedEngine", () => {
      setContent([
        ["A1", 3],
        ["B1", "=LET(t, A1*2, u, t+1, IF(u>5, u*10, t))"],
        ["C1", "=LET(s, SEQUENCE(2), s+A1)"],
        ["D1", "=LET(r, ROW(), r + A1)"],
        ["D2", "=LET(r, ROW(), r + A1)"],
      ]);
      const read = (target: FormulaEngine) =>
        ["B1", "C1", "C2", "D1", "D2"].map((ref) => cell(ref, target));
      expect(read(engine)).toEqual([70, 4, 5, 4, 5]);

      for (const includeEvaluationState of [true, false]) {
        const restored = FormulaEngine.buildEmpty();
        restored.resetToSerializedEngine(
          engine.serializeEngine({ includeEvaluationState })
        );
        expect(read(restored)).toEqual([70, 4, 5, 4, 5]);

        restored.setCellContent(address("A1"), 1);
        expect(read(restored)).toEqual([2, 2, 3, 2, 3]);
      }
    });
  });
});
