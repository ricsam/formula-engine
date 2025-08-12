import { describe, expect, test } from "bun:test";
import { add } from "./add";
import { FormulaError } from "src/core/types";

describe("add function", () => {
  test("basic number addition", () => {
    expect(
      add({ type: "number", value: 1 }, { type: "number", value: 2 })
    ).toEqual({
      type: "number",
      value: 3,
    });
  });

  test("negative number addition", () => {
    expect(
      add({ type: "number", value: -5 }, { type: "number", value: 3 })
    ).toEqual({
      type: "number",
      value: -2,
    });
  });

  test("zero addition", () => {
    expect(
      add({ type: "number", value: 0 }, { type: "number", value: 42 })
    ).toEqual({
      type: "number",
      value: 42,
    });
  });

  test("decimal addition", () => {
    expect(
      add({ type: "number", value: 1.5 }, { type: "number", value: 2.7 })
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
          { type: "number", value: 100 }
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
          { type: "infinity", sign: "positive" }
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
          { type: "number", value: 1000 }
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
          { type: "infinity", sign: "negative" }
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
          { type: "infinity", sign: "positive" }
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
          { type: "infinity", sign: "negative" }
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
          { type: "infinity", sign: "negative" }
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot add positive and negative infinity",
      });
    });

    test("negative infinity + positive infinity (undefined)", () => {
      expect(
        add(
          { type: "infinity", sign: "negative" },
          { type: "infinity", sign: "positive" }
        )
      ).toEqual({
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot add positive and negative infinity",
      });
    });
  });

  describe("overflow to infinity", () => {
    test("large positive numbers overflow to positive infinity", () => {
      expect(
        add(
          { type: "number", value: Number.MAX_VALUE },
          { type: "number", value: Number.MAX_VALUE }
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
          { type: "number", value: -Number.MAX_VALUE }
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
        add({ type: "number", value: 5 }, { type: "boolean", value: true })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add number and boolean",
      });
    });

    test("number + false should error", () => {
      expect(
        add({ type: "number", value: 10 }, { type: "boolean", value: false })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add number and boolean",
      });
    });

    test("true + number should error", () => {
      expect(
        add({ type: "boolean", value: true }, { type: "number", value: 7 })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and number",
      });
    });

    test("false + number should error", () => {
      expect(
        add({ type: "boolean", value: false }, { type: "number", value: 3 })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and number",
      });
    });

    test("true + true should error", () => {
      expect(
        add({ type: "boolean", value: true }, { type: "boolean", value: true })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and boolean",
      });
    });

    test("true + false should error", () => {
      expect(
        add({ type: "boolean", value: true }, { type: "boolean", value: false })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and boolean",
      });
    });

    test("false + false should error", () => {
      expect(
        add({ type: "boolean", value: false }, { type: "boolean", value: false })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and boolean",
      });
    });

    test("infinity + boolean should error", () => {
      expect(
        add(
          { type: "infinity", sign: "positive" },
          { type: "boolean", value: true }
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add infinity and boolean",
      });
    });
  });

  describe("error cases", () => {
    test("number + string", () => {
      expect(
        add({ type: "number", value: 5 }, { type: "string", value: "hello" })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add number and string",
      });
    });

    test("string + number", () => {
      expect(
        add({ type: "string", value: "world" }, { type: "number", value: 10 })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add string and number",
      });
    });

    test("string + string", () => {
      expect(
        add({ type: "string", value: "hello" }, { type: "string", value: "world" })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add string and string",
      });
    });

    test("boolean + string", () => {
      expect(
        add({ type: "boolean", value: true }, { type: "string", value: "test" })
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add boolean and string",
      });
    });

    test("infinity + string", () => {
      expect(
        add(
          { type: "infinity", sign: "positive" },
          { type: "string", value: "text" }
        )
      ).toEqual({
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot add infinity and string",
      });
    });
  });

  describe("edge cases", () => {
    test("very small numbers", () => {
      expect(
        add(
          { type: "number", value: Number.MIN_VALUE },
          { type: "number", value: Number.MIN_VALUE }
        )
      ).toEqual({
        type: "number",
        value: Number.MIN_VALUE * 2,
      });
    });

    test("positive and negative zero", () => {
      expect(
        add({ type: "number", value: 0 }, { type: "number", value: -0 })
      ).toEqual({
        type: "number",
        value: 0,
      });
    });

    test("NaN handling", () => {
      expect(
        add({ type: "number", value: NaN }, { type: "number", value: 5 })
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });

    test("adding to NaN", () => {
      expect(
        add({ type: "number", value: 10 }, { type: "number", value: NaN })
      ).toEqual({
        type: "number",
        value: NaN,
      });
    });
  });
});