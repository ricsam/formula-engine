import { describe, expect, test } from "bun:test";
import { multiply } from "./multiply";
import { FormulaError } from "src/core/types";

describe("multiply function", () => {
  test("basic number multiplication", () => {
    expect(
      multiply({ type: "number", value: 4 }, { type: "number", value: 3 })
    ).toEqual({
      type: "number",
      value: 12,
    });
  });

  test("negative number multiplication", () => {
    expect(
      multiply({ type: "number", value: -5 }, { type: "number", value: 3 })
    ).toEqual({
      type: "number",
      value: -15,
    });
  });

  test("negative by negative", () => {
    expect(
      multiply({ type: "number", value: -4 }, { type: "number", value: -6 })
    ).toEqual({
      type: "number",
      value: 24,
    });
  });

  test("multiplication by zero", () => {
    expect(
      multiply({ type: "number", value: 42 }, { type: "number", value: 0 })
    ).toEqual({
      type: "number",
      value: 0,
    });
  });

  test("multiplication by one", () => {
    expect(
      multiply({ type: "number", value: 7 }, { type: "number", value: 1 })
    ).toEqual({
      type: "number",
      value: 7,
    });
  });

  test("decimal multiplication", () => {
    expect(
      multiply({ type: "number", value: 2.5 }, { type: "number", value: 4 })
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
          { type: "number", value: 5 }
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
          { type: "number", value: -3 }
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
          { type: "number", value: 7 }
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
          { type: "number", value: -2 }
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
          { type: "infinity", sign: "positive" }
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
          { type: "infinity", sign: "positive" }
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
          { type: "infinity", sign: "positive" }
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
          { type: "infinity", sign: "negative" }
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
          { type: "infinity", sign: "negative" }
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
          { type: "number", value: 0 }
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot multiply infinity by zero",
      });
    });

    test("zero * infinity (indeterminate)", () => {
      expect(
        multiply(
          { type: "number", value: 0 },
          { type: "infinity", sign: "negative" }
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot multiply infinity by zero",
      });
    });
  });

  describe("overflow to infinity", () => {
    test("large positive numbers overflow", () => {
      expect(
        multiply(
          { type: "number", value: Number.MAX_VALUE },
          { type: "number", value: 2 }
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
          { type: "number", value: -2 }
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
        multiply({ type: "number", value: 5 }, { type: "string", value: "hello" })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply number and string",
      });
    });

    test("string * number", () => {
      expect(
        multiply({ type: "string", value: "world" }, { type: "number", value: 10 })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply string and number",
      });
    });

    test("boolean * number", () => {
      expect(
        multiply({ type: "boolean", value: true }, { type: "number", value: 5 })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply boolean and number",
      });
    });

    test("infinity * string", () => {
      expect(
        multiply(
          { type: "infinity", sign: "positive" },
          { type: "string", value: "text" }
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply infinity and string",
      });
    });

    test("infinity * boolean", () => {
      expect(
        multiply(
          { type: "infinity", sign: "negative" },
          { type: "boolean", value: false }
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot multiply infinity and boolean",
      });
    });
  });

  describe("edge cases", () => {
    test("very small numbers", () => {
      expect(
        multiply(
          { type: "number", value: Number.MIN_VALUE },
          { type: "number", value: 2 }
        )
      ).toEqual({
        type: "number",
        value: Number.MIN_VALUE * 2,
      });
    });

    test("NaN handling", () => {
      expect(
        multiply({ type: "number", value: NaN }, { type: "number", value: 5 })
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });

    test("multiplying by NaN", () => {
      expect(
        multiply({ type: "number", value: 10 }, { type: "number", value: NaN })
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });
  });
});