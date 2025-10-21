import { describe, expect, test } from "bun:test";
import { multiply } from "./multiply";
import { FormulaError } from "src/core/types";
import { type CellAddress } from "src/core/types";
import { EvaluationContext } from "src/evaluator/evaluation-context";
import { WorkbookManager } from "src/core/managers/workbook-manager";
import { TableManager } from "src/core/managers/table-manager";
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

describe("multiply function", () => {
  test("basic number multiplication", () => {
    expect(
      multiply({ type: "number", value: 4 }, { type: "number", value: 3 }, ctx)
    ).toEqual({
      type: "number",
      value: 12,
    });
  });

  test("negative number multiplication", () => {
    expect(
      multiply({ type: "number", value: -5 }, { type: "number", value: 3 }, ctx)
    ).toEqual({
      type: "number",
      value: -15,
    });
  });

  test("negative by negative", () => {
    expect(
      multiply(
        { type: "number", value: -4 },
        { type: "number", value: -6 },
        ctx
      )
    ).toEqual({
      type: "number",
      value: 24,
    });
  });

  test("multiplication by zero", () => {
    expect(
      multiply({ type: "number", value: 42 }, { type: "number", value: 0 }, ctx)
    ).toEqual({
      type: "number",
      value: 0,
    });
  });

  test("multiplication by one", () => {
    expect(
      multiply({ type: "number", value: 7 }, { type: "number", value: 1 }, ctx)
    ).toEqual({
      type: "number",
      value: 7,
    });
  });

  test("decimal multiplication", () => {
    expect(
      multiply(
        { type: "number", value: 2.5 },
        { type: "number", value: 4 },
        ctx
      )
    ).toEqual({
      type: "number",
      value: 10,
    });
  });

  describe("infinity handling", () => {
    test("positive infinity * positive number", () => {
      expect(
        multiply(
          { type: "infinity", sign: "positive" },
          { type: "number", value: 5 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("positive infinity * negative number", () => {
      expect(
        multiply(
          { type: "infinity", sign: "positive" },
          { type: "number", value: -3 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("negative infinity * positive number", () => {
      expect(
        multiply(
          { type: "infinity", sign: "negative" },
          { type: "number", value: 7 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("negative infinity * negative number", () => {
      expect(
        multiply(
          { type: "infinity", sign: "negative" },
          { type: "number", value: -2 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("positive number * positive infinity", () => {
      expect(
        multiply(
          { type: "number", value: 4 },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative number * positive infinity", () => {
      expect(
        multiply(
          { type: "number", value: -6 },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("positive infinity * positive infinity", () => {
      expect(
        multiply(
          { type: "infinity", sign: "positive" },
          { type: "infinity", sign: "positive" },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("positive infinity * negative infinity", () => {
      expect(
        multiply(
          { type: "infinity", sign: "positive" },
          { type: "infinity", sign: "negative" },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("negative infinity * negative infinity", () => {
      expect(
        multiply(
          { type: "infinity", sign: "negative" },
          { type: "infinity", sign: "negative" },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("infinity * zero (indeterminate)", () => {
      expect(
        multiply(
          { type: "infinity", sign: "positive" },
          { type: "number", value: 0 },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot multiply infinity by zero",
        errAddress: ctx.dependencyNode,
      });
    });

    test("zero * infinity (indeterminate)", () => {
      expect(
        multiply(
          { type: "number", value: 0 },
          { type: "infinity", sign: "negative" },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot multiply infinity by zero",
        errAddress: ctx.dependencyNode,
      });
    });
  });

  describe("overflow to infinity", () => {
    test("large positive numbers overflow", () => {
      expect(
        multiply(
          { type: "number", value: Number.MAX_VALUE },
          { type: "number", value: 2 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("large negative result overflow", () => {
      expect(
        multiply(
          { type: "number", value: Number.MAX_VALUE },
          { type: "number", value: -2 },
          ctx
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });
  });

  describe("error cases", () => {
    test("number * string", () => {
      expect(
        multiply(
          { type: "number", value: 5 },
          { type: "string", value: "hello" },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply number and string",
        errAddress: ctx.dependencyNode,
      });
    });

    test("string * number", () => {
      expect(
        multiply(
          { type: "string", value: "world" },
          { type: "number", value: 10 },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply string and number",
        errAddress: ctx.dependencyNode,
      });
    });

    test("boolean * number", () => {
      expect(
        multiply(
          { type: "boolean", value: true },
          { type: "number", value: 5 },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply boolean and number",
        errAddress: ctx.dependencyNode,
      });
    });

    test("infinity * string", () => {
      expect(
        multiply(
          { type: "infinity", sign: "positive" },
          { type: "string", value: "text" },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply infinity and string",
        errAddress: ctx.dependencyNode,
      });
    });

    test("infinity * boolean", () => {
      expect(
        multiply(
          { type: "infinity", sign: "negative" },
          { type: "boolean", value: false },
          ctx
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply infinity and boolean",
        errAddress: ctx.dependencyNode,
      });
    });
  });

  describe("edge cases", () => {
    test("very small numbers", () => {
      expect(
        multiply(
          { type: "number", value: Number.MIN_VALUE },
          { type: "number", value: 2 },
          ctx
        )
      ).toEqual({
        type: "number",
        value: Number.MIN_VALUE * 2,
      });
    });

    test("NaN handling", () => {
      expect(
        multiply(
          { type: "number", value: NaN },
          { type: "number", value: 5 },
          ctx
        )
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });

    test("multiplying by NaN", () => {
      expect(
        multiply(
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
