import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../core/engine";
import type { TableManager } from "../core/managers/table-manager";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type SerializedCellValue,
} from "../core/types";
import { parseCellReference } from "../core/utils";
import { createArrayNode, createValueNode } from "../parser/ast";
import { astToString } from "../parser/formatter";
import { parseFormula } from "../parser/parser";
import { AstEvaluationNode } from "./dependency-nodes/ast-evaluation-node";
import { EvaluationContext } from "./evaluation-context";
import type { FormulaEvaluator } from "./formula-evaluator";

describe("FormulaEvaluator", () => {
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

  describe("omitted arguments", () => {
    test("evaluate like a reference to an empty cell", () => {
      setContent([
        ["A1", '=CONCATENATE("a",,"b")'],
        ["A2", '=CONCATENATE("a",Z1,"b")'],
        ["A3", "=SUM(1,,2)"],
        ["A4", "=SUM(1,Z1,2)"],
        ["A5", "=CONCATENATE(,)"],
        ["A6", '=EXACT("",)'],
      ]);

      expect(cell("A1")).toBe("ab");
      expect(cell("A2")).toBe("ab");
      expect(cell("A3")).toBe(3);
      expect(cell("A4")).toBe(3);
      expect(cell("A5")).toBe("");
      expect(cell("A6")).toBe(true);
    });

    test("differ from Excel where this engine's blank differs from 0", () => {
      setContent([
        // Excel counts the omitted argument as a zero: 3 and 1.
        ["A1", "=COUNT(1,,2)"],
        ["A2", "=AVERAGE(1,,2)"],
      ]);

      expect(cell("A1")).toBe(2);
      expect(cell("A2")).toBe(1.5);
    });

    test("read as FALSE in logical functions, like Excel", () => {
      setContent([
        ["A1", "=AND(TRUE,)"],
        ["A2", "=OR(FALSE,)"],
        ["A3", '=IF(,"yes","no")'],
      ]);

      expect(cell("A1")).toBe(false);
      expect(cell("A2")).toBe(false);
      expect(cell("A3")).toBe("no");
    });

    test("read as empty text in TEXTJOIN, like Excel", () => {
      setContent([
        ["A1", '=TEXTJOIN("-",FALSE,"a",,"b")'],
        ["A2", '=TEXTJOIN("-",TRUE,"a",,"b")'],
      ]);

      expect(cell("A1")).toBe("a--b");
      expect(cell("A2")).toBe("a-b");
    });

    test("give IF's omitted values 0, like Excel", () => {
      setContent([
        ["A1", '=IF(TRUE,,"x")'],
        ["A2", '=IF(FALSE,"x",)'],
        // Without the comma value_if_false is missing, not omitted.
        ["A3", '=IF(FALSE,"x")'],
      ]);

      expect(cell("A1")).toBe(0);
      expect(cell("A2")).toBe(0);
      expect(cell("A3")).toBe(false);
    });

    test("recalculate with the other arguments", () => {
      setContent([
        ["A1", "x"],
        ["B1", "y"],
        ["C1", "=CONCATENATE(A1,,B1)"],
        ["C2", "=SUM(A2,,B2)"],
      ]);

      expect(cell("C1")).toBe("xy");
      expect(cell("C2")).toBe(0);

      setCellContent("B1", "z");
      setCellContent("A2", 4);
      setCellContent("B2", 5);

      expect(cell("C1")).toBe("xz");
      expect(cell("C2")).toBe(9);
    });

    test("survive warm-cache serialization", () => {
      setContent([
        ["A1", 2],
        ["B1", "=SUM(A1,,3)"],
        ["B2", "=SEQUENCE(,3,A1)"],
      ]);
      expect(cell("B1")).toBe(5);
      expect(cell("D2")).toBe(4);

      const hydrated = FormulaEngine.buildEmpty();
      hydrated.resetToSerializedEngine(engine.serializeEngine());
      engine = hydrated;

      expect(cell("B1")).toBe(5);
      expect(cell("B2")).toBe(2);
      expect(cell("D2")).toBe(4);

      setCellContent("A1", 10);
      expect(cell("B1")).toBe(13);
      expect(cell("D2")).toBe(12);
    });

    test("keep their meaning when the engine normalizes the stored formula", () => {
      setCellContent("A1", "=SUM( 1,, 2 )");
      setCellContent("A2", "=SEQUENCE( , 2 )");

      const stored = engine.getSheetSerialized(sheetAddress);
      const storedBody = (ref: string) => {
        const content = stored.get(ref);
        if (typeof content !== "string" || !content.startsWith("=")) {
          throw new Error(`Expected a formula in ${ref}`);
        }
        return astToString(parseFormula(content.slice(1)));
      };
      expect(storedBody("A1")).toBe("SUM(1,,2)");
      expect(storedBody("A2")).toBe("SEQUENCE(,2)");
      expect(cell("A1")).toBe(3);
      expect(cell("B2")).toBe(2);
    });

    test("an empty formula is an error, not an omitted argument", () => {
      setCellContent("A1", "=");
      expect(cell("A1")).toBe(FormulaError.ERROR);
    });
  });

  describe("array literals", () => {
    test("spill their elements row by row", () => {
      setContent([["A1", '={1,2;"three",TRUE}']]);

      expect(cell("A1")).toBe(1);
      expect(cell("B1")).toBe(2);
      expect(cell("A2")).toBe("three");
      expect(cell("B2")).toBe(true);
      expect(cell("C1")).toBe("");
      expect(cell("A3")).toBe("");
    });

    test("a 1x1 literal is its single value", () => {
      setContent([
        ["A1", "={5}"],
        ["A2", "=SUM({5})"],
        ["A3", "={5}+1"],
      ]);

      expect(cell("A1")).toBe(5);
      expect(cell("B1")).toBe("");
      expect(cell("A2")).toBe(5);
      expect(cell("A3")).toBe(6);
    });

    test("feed aggregate functions", () => {
      setContent([
        ["A1", "=SUM({1,2;3,4})"],
        ["A2", '=COUNT({1,"a",TRUE,2})'],
        ["A3", "=MAX({1,5;3,2})"],
        ["A4", "=MIN({4,2;3,5})"],
        ["A5", "=AVERAGE({1,2,3,4})"],
      ]);

      expect(cell("A1")).toBe(10);
      expect(cell("A2")).toBe(2);
      expect(cell("A3")).toBe(5);
      expect(cell("A4")).toBe(2);
      expect(cell("A5")).toBe(2.5);
    });

    test("are read by ranges that overlap their spill", () => {
      setContent([
        ["A1", "={1,2;3,4}"],
        ["D1", "=SUM(A1:B2)"],
        ["D2", "=SUM(B1:B2)"],
        ["D3", "=SUM(A2:B9)"],
      ]);

      expect(cell("D1")).toBe(10);
      expect(cell("D2")).toBe(6);
      expect(cell("D3")).toBe(7);
    });

    test("support lookups and operators", () => {
      setContent([
        ["A1", "=INDEX({10,20;30,40},2,1)"],
        ["A2", "=MATCH(30,{10;20;30},0)"],
        ["A3", "=SUM({1,2}*10)"],
        ["B1", "={1,2}*10"],
      ]);

      expect(cell("A1")).toBe(30);
      expect(cell("A2")).toBe(3);
      expect(cell("A3")).toBe(30);
      expect(cell("B1")).toBe(10);
      expect(cell("C1")).toBe(20);
    });

    test("keep an error as the value of its own element", () => {
      setContent([["A1", '={1,1/"x",3}']]);

      expect(cell("A1")).toBe(1);
      expect(cell("B1")).toBe(FormulaError.VALUE);
      expect(cell("C1")).toBe(3);
    });

    test("track the cells their elements reference", () => {
      setContent([
        ["A5", 1],
        ["B5", 2],
        ["A1", "={A5,B5}"],
        ["A2", "=SUM({A5,B5})"],
      ]);

      expect(cell("A1")).toBe(1);
      expect(cell("B1")).toBe(2);
      expect(cell("A2")).toBe(3);

      setCellContent("B5", 10);

      expect(cell("B1")).toBe(10);
      expect(cell("A2")).toBe(11);
    });

    test("do not nest", () => {
      setContent([["A1", "={1,{2,3}}"]]);

      expect(cell("A1")).toBe(1);
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("{} is #CALC!, Excel's error for an empty array", () => {
      setContent([["A1", "={}"]]);
      expect(cell("A1")).toBe(FormulaError.CALC);
    });

    test("a ragged literal is a syntax error", () => {
      setContent([["A1", "={1,2;3}"]]);
      expect(cell("A1")).toBe(FormulaError.ERROR);
    });

    test("a ragged tree built in code pads short rows with #N/A", () => {
      // The parser rejects ragged literals, so build the tree directly and
      // evaluate it with the engine's own evaluator.
      const internals = engine as unknown as {
        evaluationManager: { formulaEvaluator: FormulaEvaluator };
        tableManager: TableManager;
      };
      const evaluator = internals.evaluationManager.formulaEvaluator;
      const origin: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      };
      const context = new EvaluationContext(
        internals.tableManager,
        new AstEvaluationNode(parseFormula("0"), {}),
        origin
      );
      const number = (value: number) =>
        createValueNode({ type: "number", value });
      const ragged = createArrayNode([[number(1), number(2)], [number(3)]]);

      const result = evaluator.evaluateArray(ragged, context);
      if (result.type !== "spilled-values") {
        throw new Error(`Expected a spilled result, got ${result.type}`);
      }
      expect(result.spillArea(origin)).toEqual({
        start: { col: 0, row: 0 },
        end: {
          col: { type: "number", value: 1 },
          row: { type: "number", value: 1 },
        },
      });
      expect(result.evaluate({ x: 1, y: 1 }, context)).toMatchObject({
        type: "error",
        err: FormulaError.NA,
      });

      const cells = result.evaluateAllCells.call(evaluator, {
        context,
        evaluate: result.evaluate,
        origin,
        lookupOrder: "row-major",
      });
      if (cells.type !== "values") {
        throw new Error(`Expected values, got ${cells.type}`);
      }
      expect(
        cells.values.map(({ relativePos, result }) => [
          relativePos,
          result.type === "value" ? result.result : result.type === "error" ? result.err : result.type,
        ])
      ).toEqual([
        [{ x: 0, y: 0 }, { type: "number", value: 1 }],
        [{ x: 1, y: 0 }, { type: "number", value: 2 }],
        [{ x: 0, y: 1 }, { type: "number", value: 3 }],
        [{ x: 1, y: 1 }, FormulaError.NA],
      ]);
    });

    test("read every cell in the requested lookup order", () => {
      const internals = engine as unknown as {
        evaluationManager: { formulaEvaluator: FormulaEvaluator };
        tableManager: TableManager;
      };
      const evaluator = internals.evaluationManager.formulaEvaluator;
      const origin: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 2,
        rowIndex: 3,
      };
      const context = new EvaluationContext(
        internals.tableManager,
        new AstEvaluationNode(parseFormula("0"), {}),
        origin
      );
      const literal = parseFormula("{1,2;3,4}");
      if (literal.type !== "array") throw new Error("Expected an array");
      const result = evaluator.evaluateArray(literal, context);
      if (result.type !== "spilled-values") {
        throw new Error(`Expected a spilled result, got ${result.type}`);
      }

      const read = (
        lookupOrder: "row-major" | "col-major",
        intersection?: Parameters<
          typeof result.evaluateAllCells
        >[0]["intersection"]
      ) => {
        const cells = result.evaluateAllCells.call(evaluator, {
          context,
          evaluate: result.evaluate,
          origin,
          lookupOrder,
          intersection,
        });
        if (cells.type !== "values") throw new Error("Expected values");
        return cells.values.map(({ relativePos, result }) => [
          relativePos.x,
          relativePos.y,
          result.type === "value" ? result.result : undefined,
        ]);
      };
      const n = (value: number): CellValue => ({ type: "number", value });

      expect(read("row-major")).toEqual([
        [0, 0, n(1)],
        [1, 0, n(2)],
        [0, 1, n(3)],
        [1, 1, n(4)],
      ]);
      expect(read("col-major")).toEqual([
        [0, 0, n(1)],
        [0, 1, n(3)],
        [1, 0, n(2)],
        [1, 1, n(4)],
      ]);
      // Intersections are in sheet coordinates; positions stay relative to the
      // origin (C4).
      expect(
        read("row-major", {
          start: { col: 3, row: 0 },
          end: {
            col: { type: "infinity", sign: "positive" },
            row: { type: "infinity", sign: "positive" },
          },
        })
      ).toEqual([
        [1, 0, n(2)],
        [1, 1, n(4)],
      ]);
      expect(
        read("row-major", {
          start: { col: 10, row: 10 },
          end: {
            col: { type: "number", value: 12 },
            row: { type: "number", value: 12 },
          },
        })
      ).toEqual([]);
    });
  });

  describe("unary operators over arrays", () => {
    test("apply the operator once when an aggregate reads every cell", () => {
      setContent([
        ["A1", "=SUM(-SEQUENCE(3))"],
        // UNIQUE reads cells through the evaluate callback it is given.
        ["A2", "=SUM(-UNIQUE({1;2;3}))"],
        ["A3", "=SUM(-{1,2;3,4})"],
        ["A4", "=SUM(--SEQUENCE(3))"],
        ["A5", "=SUM(SEQUENCE(4)%)"],
      ]);

      expect(cell("A1")).toBe(-6);
      expect(cell("A2")).toBe(-6);
      expect(cell("A3")).toBe(-10);
      expect(cell("A4")).toBe(6);
      expect(cell("A5")).toBe(0.1);
    });

    test("spill and feed ranges that read the spill", () => {
      setContent([
        ["A1", "=-SEQUENCE(3)"],
        ["B1", "=SUM(A1:A3)"],
        ["B2", "=SUM(A2:A9)"],
      ]);

      expect(cell("A1")).toBe(-1);
      expect(cell("A3")).toBe(-3);
      expect(cell("B1")).toBe(-6);
      expect(cell("B2")).toBe(-5);
    });
  });
});
