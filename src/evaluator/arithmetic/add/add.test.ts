import { describe, expect, test } from "bun:test";
import { add } from "./add";
import { FormulaError, type CellAddress } from "src/core/types";

const errAddress: CellAddress = {
  sheetName: "Sheet1",
  workbookName: "Workbook1",
  colIndex: 1,
  rowIndex: 1,
};
describe("add function", () => {
  test("basic number addition", () => {
    expect(
      add(
        { type: "number", value: 1 },
        { type: "number", value: 2 },
        errAddress
      )
    ).toEqual({
      type: "number",
      value: 3,
    });
  });

  test("negative number addition", () => {
    expect(
      add({ type: "number", value: -5 }, { type: "number", value: 3 }, errAddress)
    ).toEqual({
      type: "number",
      value: -2,
    });
  });

  test("zero addition", () => {
    expect(
      add({ type: "number", value: 0 }, { type: "number", value: 42 }, errAddress)
    ).toEqual({
      type: "number",
      value: 42,
    });
  });

  test("decimal addition", () => {
    expect(
      add({ type: "number", value: 1.5 }, { type: "number", value: 2.7 }, errAddress)
    ).toEqual({
      type: "number",
      value: 4.2,
    });
  });

  describe("infinity handling", () => {
    test("positive infinity + number", () => {
      expect(
        add(
          { type: "infinity", sign: "positive" },
          { type: "number", value: 100 },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("number + positive infinity", () => {
      expect(
        add(
          { type: "number", value: -50 },
          { type: "infinity", sign: "positive" },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative infinity + number", () => {
      expect(
        add(
          { type: "infinity", sign: "negative" },
          { type: "number", value: 1000 },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("number + negative infinity", () => {
      expect(
        add(
          { type: "number", value: 999 },
          { type: "infinity", sign: "negative" },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("positive infinity + positive infinity", () => {
      expect(
        add(
          { type: "infinity", sign: "positive" },
          { type: "infinity", sign: "positive" },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative infinity + negative infinity", () => {
      expect(
        add(
          { type: "infinity", sign: "negative" },
          { type: "infinity", sign: "negative" },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("positive infinity + negative infinity (undefined)", () => {
      expect(
        add(
          { type: "infinity", sign: "positive" },
          { type: "infinity", sign: "negative" },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot add positive and negative infinity",
        errAddress: errAddress,
      });
    });

    test("negative infinity + positive infinity (undefined)", () => {
      expect(
        add(
          { type: "infinity", sign: "negative" },
          { type: "infinity", sign: "positive" },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot add positive and negative infinity",
        errAddress: errAddress,
      });
    });
  });

  describe("overflow to infinity", () => {
    test("large positive numbers overflow to positive infinity", () => {
      expect(
        add(
          { type: "number", value: Number.MAX_VALUE },
          { type: "number", value: Number.MAX_VALUE },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("large negative numbers overflow to negative infinity", () => {
      expect(
        add(
          { type: "number", value: -Number.MAX_VALUE },
          { type: "number", value: -Number.MAX_VALUE },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });
  });

  describe("boolean error cases", () => {
    test("number + true should error", () => {
      expect(
        add({ type: "number", value: 5 }, { type: "boolean", value: true }, errAddress)
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add number and boolean",
        errAddress: errAddress,
      });
    });

    test("number + false should error", () => {
      expect(
        add({ type: "number", value: 10 }, { type: "boolean", value: false }, errAddress)
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add number and boolean",
        errAddress: errAddress,
      });
    });

    test("true + number should error", () => {
      expect(
        add({ type: "boolean", value: true }, { type: "number", value: 7 }, errAddress)
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and number",
        errAddress: errAddress,
      });
    });

    test("false + number should error", () => {
      expect(
        add({ type: "boolean", value: false }, { type: "number", value: 3 }, errAddress)
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and number",
        errAddress: errAddress,
      });
    });

    test("true + true should error", () => {
      expect(
        add({ type: "boolean", value: true }, { type: "boolean", value: true }, errAddress)
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and boolean",
        errAddress: errAddress,
      });
    });

    test("true + false should error", () => {
      expect(
        add({ type: "boolean", value: true }, { type: "boolean", value: false }, errAddress)
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and boolean",
        errAddress: errAddress,
        });
    });

    test("false + false should error", () => {
      expect(
        add(
          { type: "boolean", value: false },
          { type: "boolean", value: false },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and boolean",
        errAddress: errAddress,
      });
    });

    test("infinity + boolean should error", () => {
      expect(
        add(
          { type: "infinity", sign: "positive" },
          { type: "boolean", value: true },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add infinity and boolean",
        errAddress: errAddress,
      });
    });
  });

  describe("error cases", () => {
    test("number + string", () => {
      expect(
        add({ type: "number", value: 5 }, { type: "string", value: "hello" }, errAddress)
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add number and string",
        errAddress: errAddress,
      });
    });

    test("string + number", () => {
      expect(
        add({ type: "string", value: "world" }, { type: "number", value: 10 }, errAddress)
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add string and number",
        errAddress: errAddress,
      });
    });

    test("string + string", () => {
      expect(
        add(
          { type: "string", value: "hello" },
          { type: "string", value: "world" },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add string and string",
        errAddress: errAddress,
      });
    });

    test("boolean + string", () => {
      expect(
        add({ type: "boolean", value: true }, { type: "string", value: "test" }, errAddress)
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and string",
        errAddress: errAddress,
        });
    });

    test("infinity + string", () => {
      expect(
        add(
          { type: "infinity", sign: "positive" },
          { type: "string", value: "text" },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add infinity and string",
        errAddress: errAddress,
      });
    });
  });

  describe("edge cases", () => {
    test("very small numbers", () => {
      expect(
        add(
          { type: "number", value: Number.MIN_VALUE },
          { type: "number", value: Number.MIN_VALUE },
          errAddress
        )
      ).toEqual({
        type: "number",
        value: Number.MIN_VALUE * 2,
      });
    });

    test("positive and negative zero", () => {
      expect(
        add({ type: "number", value: 0 }, { type: "number", value: -0 }, errAddress)
      ).toEqual({
        type: "number",
        value: 0,
      });
    });

    test("NaN handling", () => {
      expect(
        add({ type: "number", value: NaN }, { type: "number", value: 5 }, errAddress)
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });

    test("adding to NaN", () => {
      expect(
        add({ type: "number", value: 10 }, { type: "number", value: NaN }, errAddress)
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });
  });
});
