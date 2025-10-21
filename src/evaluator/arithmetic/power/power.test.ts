import { describe, expect, test } from "bun:test";
import { power } from "./power";
import { FormulaError, type CellAddress } from "src/core/types";
import { EvaluationContext } from "src/evaluator/evaluation-context";
import { TableManager } from "src/core/managers/table-manager";
import { WorkbookManager } from "src/core/managers/workbook-manager";
import { CellValueNode } from "src/evaluator/dependency-nodes/cell-value-node";

const errAddress: CellAddress = {
  sheetName: "Sheet1",
  workbookName: "Workbook1",
  colIndex: 0,
  rowIndex: 0,
};
const workbookManager = new WorkbookManager();
const tableManager = new TableManager(workbookManager);
const dependencyNode = new CellValueNode("cell-value:Workbook1:Sheet1:A1");
const ctx = new EvaluationContext(tableManager, dependencyNode, errAddress);

describe("power function", () => {
  test("basic number exponentiation", () => {
    expect(
      power({ type: "number", value: 2 }, { type: "number", value: 3 }, ctx)
    ).toEqual({
      type: "number",
      value: 8,
    });
  });

  test("square root (fractional exponent)", () => {
    expect(
      power({ type: "number", value: 9 }, { type: "number", value: 0.5 }, ctx)
    ).toEqual({
      type: "number",
      value: 3,
    });
  });

  test("power of zero", () => {
    expect(
      power({ type: "number", value: 5 }, { type: "number", value: 0 }, ctx)
    ).toEqual({
      type: "number",
      value: 1,
    });
  });

  test("zero to positive power", () => {
    expect(
      power({ type: "number", value: 0 }, { type: "number", value: 3 }, ctx)
    ).toEqual({
      type: "number",
      value: 0,
    });
  });

  test("zero to zero", () => {
    expect(
      power({ type: "number", value: 0 }, { type: "number", value: 0 }, ctx)
    ).toEqual({
      type: "number",
      value: 1,
    });
  });

  test("zero to negative power", () => {
    expect(
      power({ type: "number", value: 0 }, { type: "number", value: -2 }, ctx)
    ).toEqual({
      type: "infinity",
      sign: "positive",
    });
  });

  test("negative base to even integer power", () => {
    expect(
      power({ type: "number", value: -3 }, { type: "number", value: 2 }, ctx)
    ).toEqual({
      type: "number",
      value: 9,
    });
  });

  test("negative base to odd integer power", () => {
    expect(
      power({ type: "number", value: -2 }, { type: "number", value: 3 }, ctx)
    ).toEqual({
      type: "number",
      value: -8,
    });
  });

  test("negative base to non-integer power (error)", () => {
    expect(
      power({ type: "number", value: -4 }, { type: "number", value: 0.5 }, ctx)
    ).toEqual({
      type: "error",
      err: FormulaError.NUM,
      message: "Cannot raise negative number to non-integer power",
      errAddress: ctx.dependencyNode,
    });
  });

  test("one to any power", () => {
    expect(
      power({ type: "number", value: 1 }, { type: "number", value: 100 }, ctx)
    ).toEqual({
      type: "number",
      value: 1,
    });
  });

  test("negative exponent", () => {
    expect(
      power({ type: "number", value: 2 }, { type: "number", value: -3 }, ctx)
    ).toEqual({
      type: "number",
      value: 0.125,
    });
  });

  describe("infinity handling", () => {
    test("positive infinity to positive power", () => {
      expect(
        power(
          { type: "infinity", sign: "positive" },
          { type: "number", value: 2 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative infinity to even integer power", () => {
      expect(
        power(
          { type: "infinity", sign: "negative" },
          { type: "number", value: 2 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative infinity to odd integer power", () => {
      expect(
        power(
          { type: "infinity", sign: "negative" },
          { type: "number", value: 3 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("negative infinity to non-integer power (error)", () => {
      expect(
        power(
          { type: "infinity", sign: "negative" },
          { type: "number", value: 2.5 },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot raise negative infinity to non-integer power",
        errAddress: ctx.dependencyNode,
      });
    });

    test("infinity to zero power", () => {
      expect(
        power(
          { type: "infinity", sign: "positive" },
          { type: "number", value: 0 },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 1,
      });
    });

    test("infinity to negative power", () => {
      expect(
        power(
          { type: "infinity", sign: "positive" },
          { type: "number", value: -2 },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 0,
      });
    });

    test("number > 1 to positive infinity", () => {
      expect(
        power(
          { type: "number", value: 2 },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("number > 1 to negative infinity", () => {
      expect(
        power(
          { type: "number", value: 3 },
          { type: "infinity", sign: "negative" },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 0,
      });
    });

    test("0 < number < 1 to positive infinity", () => {
      expect(
        power(
          { type: "number", value: 0.5 },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 0,
      });
    });

    test("0 < number < 1 to negative infinity", () => {
      expect(
        power(
          { type: "number", value: 0.8 },
          { type: "infinity", sign: "negative" },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("1 to positive infinity", () => {
      expect(
        power(
          { type: "number", value: 1 },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 1,
      });
    });

    test("1 to negative infinity", () => {
      expect(
        power(
          { type: "number", value: 1 },
          { type: "infinity", sign: "negative" },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 1,
      });
    });

    test("-1 to positive infinity", () => {
      expect(
        power(
          { type: "number", value: -1 },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 1,
      });
    });

    test("infinity to infinity", () => {
      expect(
        power(
          { type: "infinity", sign: "positive" },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative infinity to infinity", () => {
      expect(
        power(
          { type: "infinity", sign: "negative" },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });
  });

  describe("overflow to infinity", () => {
    test("large base to large exponent", () => {
      expect(
        power(
          { type: "number", value: Number.MAX_VALUE },
          { type: "number", value: 2 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative overflow", () => {
      expect(
        power(
          { type: "number", value: -Number.MAX_VALUE },
          { type: "number", value: 3 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });
  });

  describe("error cases", () => {
    test("number ^ string", () => {
      expect(
        power(
          { type: "number", value: 5 },
          { type: "string", value: "hello" },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot exponentiate number and string",
        errAddress: ctx.dependencyNode,
      });
    });

    test("string ^ number", () => {
      expect(
        power(
          { type: "string", value: "world" },
          { type: "number", value: 2 },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot exponentiate string and number",
        errAddress: ctx.dependencyNode,
      });
    });

    test("boolean ^ number", () => {
      expect(
        power(
          { type: "boolean", value: true },
          { type: "number", value: 3 },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot exponentiate boolean and number",
        errAddress: ctx.dependencyNode,
      });
    });

    test("infinity ^ string", () => {
      expect(
        power(
          { type: "infinity", sign: "positive" },
          { type: "string", value: "text" },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot exponentiate infinity and string",
        errAddress: ctx.dependencyNode,
      });
    });

    test("infinity ^ boolean", () => {
      expect(
        power(
          { type: "infinity", sign: "negative" },
          { type: "boolean", value: false },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot exponentiate infinity and boolean",
        errAddress: ctx.dependencyNode,
      });
    });
  });

  describe("edge cases", () => {
    test("very small base", () => {
      expect(
        power(
          { type: "number", value: Number.MIN_VALUE },
          { type: "number", value: 2 },
          ctx
        )
      ).toEqual({
        type: "number",
        value: Math.pow(Number.MIN_VALUE, 2),
      });
    });

    test("NaN handling", () => {
      expect(
        power({ type: "number", value: NaN }, { type: "number", value: 2 }, ctx)
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });

    test("exponent is NaN", () => {
      expect(
        power({ type: "number", value: 5 }, { type: "number", value: NaN }, ctx)
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });
  });
});
