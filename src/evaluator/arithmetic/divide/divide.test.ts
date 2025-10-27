import { describe, expect, test } from "bun:test";
import { divide } from "./divide";
import { FormulaError, type CellAddress } from "../../../core/types";
import { EvaluationContext } from "../../evaluation-context";
import { CellValueNode } from "../../dependency-nodes/cell-value-node";
import { WorkbookManager } from "../../../core/managers/workbook-manager";
import { TableManager } from "../../../core/managers/table-manager";

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

describe("divide function", () => {
  test("basic number division", () => {
    expect(
      divide({ type: "number", value: 12 }, { type: "number", value: 3 }, ctx)
    ).toEqual({
      type: "number",
      value: 4,
    });
  });

  test("division with decimal result", () => {
    expect(
      divide({ type: "number", value: 7 }, { type: "number", value: 2 }, ctx)
    ).toEqual({
      type: "number",
      value: 3.5,
    });
  });

  test("negative dividend", () => {
    expect(
      divide({ type: "number", value: -15 }, { type: "number", value: 3 }, ctx)
    ).toEqual({
      type: "number",
      value: -5,
    });
  });

  test("negative divisor", () => {
    expect(
      divide({ type: "number", value: 20 }, { type: "number", value: -4 }, ctx)
    ).toEqual({
      type: "number",
      value: -5,
    });
  });

  test("negative by negative", () => {
    expect(
      divide({ type: "number", value: -18 }, { type: "number", value: -6 }, ctx)
    ).toEqual({
      type: "number",
      value: 3,
    });
  });

  test("division by one", () => {
    expect(
      divide({ type: "number", value: 42 }, { type: "number", value: 1 }, ctx)
    ).toEqual({
      type: "number",
      value: 42,
    });
  });

  test("zero dividend", () => {
    expect(
      divide({ type: "number", value: 0 }, { type: "number", value: 5 }, ctx)
    ).toEqual({
      type: "number",
      value: 0,
    });
  });

  describe("division by zero", () => {
    test("positive number / zero", () => {
      expect(
        divide({ type: "number", value: 5 }, { type: "number", value: 0 }, ctx)
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative number / zero", () => {
      expect(
        divide({ type: "number", value: -3 }, { type: "number", value: 0 }, ctx)
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("zero / zero (indeterminate)", () => {
      expect(
        divide({ type: "number", value: 0 }, { type: "number", value: 0 }, ctx)
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "0 / 0 is undefined",
        errAddress: ctx.dependencyNode,
      });
    });
  });

  describe("infinity handling", () => {
    test("positive infinity / positive number", () => {
      expect(
        divide(
          { type: "infinity", sign: "positive" },
          { type: "number", value: 5 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("positive infinity / negative number", () => {
      expect(
        divide(
          { type: "infinity", sign: "positive" },
          { type: "number", value: -3 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("negative infinity / positive number", () => {
      expect(
        divide(
          { type: "infinity", sign: "negative" },
          { type: "number", value: 7 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("negative infinity / negative number", () => {
      expect(
        divide(
          { type: "infinity", sign: "negative" },
          { type: "number", value: -2 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("positive number / positive infinity", () => {
      expect(
        divide(
          { type: "number", value: 100 },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 0,
      });
    });

    test("negative number / positive infinity", () => {
      expect(
        divide(
          { type: "number", value: -50 },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 0,
      });
    });

    test("number / negative infinity", () => {
      expect(
        divide(
          { type: "number", value: 25 },
          { type: "infinity", sign: "negative" },
          ctx
        )
      ).toEqual({
        type: "number",
        value: 0,
      });
    });

    test("infinity / infinity (indeterminate)", () => {
      expect(
        divide(
          { type: "infinity", sign: "positive" },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot divide infinity by infinity",
        errAddress: ctx.dependencyNode,
      });
    });

    test("positive infinity / negative infinity (indeterminate)", () => {
      expect(
        divide(
          { type: "infinity", sign: "positive" },
          { type: "infinity", sign: "negative" },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot divide infinity by infinity",
        errAddress: ctx.dependencyNode,
      });
    });

    test("infinity / zero (error)", () => {
      expect(
        divide(
          { type: "infinity", sign: "positive" },
          { type: "number", value: 0 },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot divide infinity by zero",
        errAddress: ctx.dependencyNode,
      });
    });
  });

  describe("overflow to infinity", () => {
    test("large dividend overflow", () => {
      expect(
        divide(
          { type: "number", value: Number.MAX_VALUE },
          { type: "number", value: Number.MIN_VALUE },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("very small divisor causing overflow", () => {
      expect(
        divide(
          { type: "number", value: 1 },
          { type: "number", value: Number.MIN_VALUE },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });
  });

  describe("error cases", () => {
    test("number / string", () => {
      expect(
        divide(
          { type: "number", value: 5 },
          { type: "string", value: "hello" },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot divide number and string",
        errAddress: ctx.dependencyNode,
      });
    });

    test("string / number", () => {
      expect(
        divide(
          { type: "string", value: "world" },
          { type: "number", value: 10 },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot divide string and number",
        errAddress: ctx.dependencyNode,
      });
    });

    test("boolean / number", () => {
      expect(
        divide(
          { type: "boolean", value: true },
          { type: "number", value: 5 },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot divide boolean and number",
        errAddress: ctx.dependencyNode,
      });
    });

    test("infinity / string", () => {
      expect(
        divide(
          { type: "infinity", sign: "positive" },
          { type: "string", value: "text" },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot divide infinity and string",
        errAddress: ctx.dependencyNode,
      });
    });

    test("infinity / boolean", () => {
      expect(
        divide(
          { type: "infinity", sign: "negative" },
          { type: "boolean", value: true },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot divide infinity and boolean",
        errAddress: ctx.dependencyNode,
      });
    });
  });

  describe("edge cases", () => {
    test("very small numbers", () => {
      expect(
        divide(
          { type: "number", value: Number.MIN_VALUE },
          { type: "number", value: 2 },
          ctx
        )
      ).toEqual({
        type: "number",
        value: Number.MIN_VALUE / 2,
      });
    });

    test("NaN handling", () => {
      expect(
        divide(
          { type: "number", value: NaN },
          { type: "number", value: 5 },
          ctx
        )
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });

    test("dividing by NaN", () => {
      expect(
        divide(
          { type: "number", value: 10 },
          { type: "number", value: NaN },
          ctx
        )
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });
  });
});
