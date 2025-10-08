import { describe, expect, test } from "bun:test";
import { subtract } from "./subtract";
import { FormulaError } from "src/core/types";
import { type CellAddress } from "src/core/types";
const errAddress: CellAddress = {
  sheetName: "Sheet1",
  workbookName: "Workbook1",
  colIndex: 1,
  rowIndex: 1,
};

describe("subtract function", () => {
  test("basic number subtraction", () => {
    expect(
      subtract(
        { type: "number", value: 5 },
        { type: "number", value: 3 },
        errAddress
      )
    ).toEqual({
      type: "number",
      value: 2,
    });
  });

  test("negative result", () => {
    expect(
      subtract(
        { type: "number", value: 3 },
        { type: "number", value: 8 },
        errAddress
      )
    ).toEqual({
      type: "number",
      value: -5,
    });
  });

  test("zero subtraction", () => {
    expect(
      subtract(
        { type: "number", value: 42 },
        { type: "number", value: 0 },
        errAddress
      )
    ).toEqual({
      type: "number",
      value: 42,
    });
  });

  test("subtracting from zero", () => {
    expect(
      subtract(
        { type: "number", value: 0 },
        { type: "number", value: 7 },
        errAddress
      )
    ).toEqual({
      type: "number",
      value: -7,
    });
  });

  test("decimal subtraction", () => {
    const result = subtract(
      { type: "number", value: 5.7 },
      { type: "number", value: 2.3 },
      errAddress
    );
    expect(result.type).toBe("number");
    if (result.type === "number") {
      expect(result.value).toBeCloseTo(3.4, 10);
    }
  });

  describe("infinity handling", () => {
    test("positive infinity - number", () => {
      expect(
        subtract(
          { type: "infinity", sign: "positive" },
          { type: "number", value: 100 },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative infinity - number", () => {
      expect(
        subtract(
          { type: "infinity", sign: "negative" },
          { type: "number", value: 50 },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("number - positive infinity", () => {
      expect(
        subtract(
          { type: "number", value: 100 },
          { type: "infinity", sign: "positive" },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });

    test("number - negative infinity", () => {
      expect(
        subtract(
          { type: "number", value: 50 },
          { type: "infinity", sign: "negative" },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("positive infinity - positive infinity (undefined)", () => {
      expect(
        subtract(
          { type: "infinity", sign: "positive" },
          { type: "infinity", sign: "positive" },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot subtract infinity from same-signed infinity",
        errAddress: errAddress,
      });
    });

    test("negative infinity - negative infinity (undefined)", () => {
      expect(
        subtract(
          { type: "infinity", sign: "negative" },
          { type: "infinity", sign: "negative" },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot subtract infinity from same-signed infinity",
        errAddress: errAddress,
      });
    });

    test("positive infinity - negative infinity", () => {
      expect(
        subtract(
          { type: "infinity", sign: "positive" },
          { type: "infinity", sign: "negative" },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("negative infinity - positive infinity", () => {
      expect(
        subtract(
          { type: "infinity", sign: "negative" },
          { type: "infinity", sign: "positive" },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });
  });

  describe("overflow to infinity", () => {
    test("large positive result overflow", () => {
      expect(
        subtract(
          { type: "number", value: Number.MAX_VALUE },
          { type: "number", value: -Number.MAX_VALUE },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "positive",
      });
    });

    test("large negative result overflow", () => {
      expect(
        subtract(
          { type: "number", value: -Number.MAX_VALUE },
          { type: "number", value: Number.MAX_VALUE },
          errAddress
        )
      ).toEqual({
        type: "infinity",
        sign: "negative",
      });
    });
  });

  describe("error cases", () => {
    test("number - string", () => {
      expect(
        subtract(
          { type: "number", value: 5 },
          { type: "string", value: "hello" },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot subtract number and string",
        errAddress: errAddress,
      });
    });

    test("string - number", () => {
      expect(
        subtract(
          { type: "string", value: "world" },
          { type: "number", value: 10 },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot subtract string and number",
        errAddress: errAddress,
      });
    });

    test("boolean - number", () => {
      expect(
        subtract(
          { type: "boolean", value: true },
          { type: "number", value: 5 },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot subtract boolean and number",
        errAddress: errAddress,
      });
    });

    test("number - boolean", () => {
      expect(
        subtract(
          { type: "number", value: 10 },
          { type: "boolean", value: false },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot subtract number and boolean",
        errAddress: errAddress,
      });
    });

    test("infinity - string", () => {
      expect(
        subtract(
          { type: "infinity", sign: "positive" },
          { type: "string", value: "text" },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot subtract infinity and string",
        errAddress: errAddress,
      });
    });

    test("infinity - boolean", () => {
      expect(
        subtract(
          { type: "infinity", sign: "negative" },
          { type: "boolean", value: true },
          errAddress
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot subtract infinity and boolean",
        errAddress: errAddress,
      });
    });
  });

  describe("edge cases", () => {
    test("very small numbers", () => {
      expect(
        subtract(
          { type: "number", value: Number.MIN_VALUE },
          { type: "number", value: Number.MIN_VALUE },
          errAddress
        )
      ).toEqual({
        type: "number",
        value: 0,
      });
    });

    test("NaN handling", () => {
      expect(
        subtract(
          { type: "number", value: NaN },
          { type: "number", value: 5 },
          errAddress
        )
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });

    test("subtracting from NaN", () => {
      expect(
        subtract(
          { type: "number", value: 10 },
          { type: "number", value: NaN },
          errAddress
        )
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });
  });
});
