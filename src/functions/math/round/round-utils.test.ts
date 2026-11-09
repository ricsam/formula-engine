import { describe, expect, test } from "bun:test";
import { roundDecimal, toSignificantDigits } from "./round-utils";

describe("toSignificantDigits", () => {
  test("returns 15 significant digits and the exponent of the first", () => {
    expect(toSignificantDigits(2.675)).toEqual({
      digits: "267500000000000",
      exponent: 0,
    });
    expect(toSignificantDigits(0.0125)).toEqual({
      digits: "125000000000000",
      exponent: -2,
    });
    expect(toSignificantDigits(0.1 + 0.2)).toEqual({
      digits: "300000000000000",
      exponent: -1,
    });
    expect(toSignificantDigits(1e21)).toEqual({
      digits: "100000000000000",
      exponent: 21,
    });
  });

  test("rounds the 16th digit half up from the exact binary value", () => {
    // 1/3 continues with 3s; 2/3 with 6s
    expect(toSignificantDigits(1 / 3).digits).toBe("333333333333333");
    expect(toSignificantDigits(2 / 3).digits).toBe("666666666666667");
  });

  test("rounds an exact tie in the 16th digit toward zero, as Excel does", () => {
    expect(toSignificantDigits(1000000000000005)).toEqual({
      digits: "100000000000000",
      exponent: 15,
    });
    expect(toSignificantDigits(123456789012345.5)).toEqual({
      digits: "123456789012345",
      exponent: 14,
    });
    expect(toSignificantDigits(99999999999999.25)).toEqual({
      digits: "999999999999992",
      exponent: 13,
    });
    // Rounding up would carry into a new digit
    expect(toSignificantDigits(999999999999999.5)).toEqual({
      digits: "999999999999999",
      exponent: 14,
    });
  });
});

describe("roundDecimal", () => {
  test("rounds in each mode", () => {
    expect(roundDecimal(2.5, 0, "half-away-from-zero")).toBe(3);
    expect(roundDecimal(-2.5, 0, "half-away-from-zero")).toBe(-3);
    expect(roundDecimal(2.1, 0, "away-from-zero")).toBe(3);
    expect(roundDecimal(-2.1, 0, "away-from-zero")).toBe(-3);
    expect(roundDecimal(2.9, 0, "toward-zero")).toBe(2);
    expect(roundDecimal(-2.9, 0, "toward-zero")).toBe(-2);
  });

  test("works on the 15-digit decimal value", () => {
    expect(roundDecimal(2.675, 2, "half-away-from-zero")).toBe(2.68);
    expect(roundDecimal(1.005, 2, "half-away-from-zero")).toBe(1.01);
    expect(roundDecimal(0.1 + 0.2, 1, "away-from-zero")).toBe(0.3);
    expect(roundDecimal(4.35 * 100, 0, "toward-zero")).toBe(435);
  });

  test("rounds positions above the first significant digit", () => {
    expect(roundDecimal(0.004, 0, "half-away-from-zero")).toBe(0);
    expect(roundDecimal(0.004, 0, "away-from-zero")).toBe(1);
    expect(roundDecimal(0.6, 0, "half-away-from-zero")).toBe(1);
    expect(roundDecimal(49, -2, "half-away-from-zero")).toBe(0);
    expect(roundDecimal(50, -2, "half-away-from-zero")).toBe(100);
    expect(roundDecimal(5, -3, "away-from-zero")).toBe(1000);
  });

  test("rounds numbers with more than 15 significant digits", () => {
    // 16 significant digits: the rounding position is the 15th
    expect(roundDecimal(1234567890123.456, 2, "half-away-from-zero")).toBe(
      1234567890123.46
    );
    expect(roundDecimal(1234567890123.459, 2, "toward-zero")).toBe(
      1234567890123.46
    );
    expect(roundDecimal(1 / 3, 15, "half-away-from-zero")).toBe(
      0.333333333333333
    );
    // Binary noise beyond the 15th digit is dropped in every mode
    expect(roundDecimal(0.1 + 0.2, 15, "away-from-zero")).toBe(0.3);
    expect(roundDecimal(0.1 + 0.2, 16, "half-away-from-zero")).toBe(0.3);
  });

  test("keeps a value with no more digits than are kept", () => {
    expect(roundDecimal(1 / 3, 16, "half-away-from-zero")).toBe(1 / 3);
    expect(roundDecimal(1 / 3, 20, "away-from-zero")).toBe(1 / 3);
    expect(roundDecimal(123456789012345, 0, "half-away-from-zero")).toBe(
      123456789012345
    );
    expect(roundDecimal(2.5, 1000, "toward-zero")).toBe(2.5);
  });

  test("never returns negative zero", () => {
    expect(roundDecimal(-0.4, 0, "half-away-from-zero")).toBe(0);
    expect(roundDecimal(-0.9, 0, "toward-zero")).toBe(0);
    expect(roundDecimal(-4, -1, "half-away-from-zero")).toBe(0);
  });

  test("returns undefined when the result is too large for a number", () => {
    expect(roundDecimal(5, -400, "away-from-zero")).toBeUndefined();
    expect(roundDecimal(1.7e308, -308, "half-away-from-zero")).toBeUndefined();
    expect(roundDecimal(5, -400, "half-away-from-zero")).toBe(0);
    expect(roundDecimal(Infinity, 0, "half-away-from-zero")).toBeUndefined();
  });
});
