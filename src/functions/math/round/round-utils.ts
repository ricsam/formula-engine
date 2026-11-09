/**
 * Decimal rounding shared by ROUND, ROUNDUP, ROUNDDOWN and TEXT.
 *
 * JavaScript numbers are binary, so `2.675` is stored as 2.67499999999999982...
 * and `Math.round(2.675 * 100) / 100` is 2.67. Excel rounds the number it
 * shows, its decimal representation with 15 significant digits, so
 * `ROUND(2.675, 2)` is 2.68 and `ROUNDUP(0.1 + 0.2, 1)` is 0.3, not 0.4. These
 * helpers do the same: they round the 15-digit decimal representation and
 * build the result from decimal text, so it is the double closest to the
 * rounded decimal value.
 */

/**
 * How a rounded-off part is resolved:
 * - `"half-away-from-zero"`: ROUND, 2.5 -> 3 and -2.5 -> -3
 * - `"away-from-zero"`: ROUNDUP, 2.1 -> 3 and -2.1 -> -3
 * - `"toward-zero"`: ROUNDDOWN, 2.9 -> 2 and -2.9 -> -2
 */
export type DecimalRoundingMode =
  | "half-away-from-zero"
  | "away-from-zero"
  | "toward-zero";

/** Significant digits Excel keeps, and that the rounding works on. */
const SIGNIFICANT_DIGITS = 15;

/**
 * Rounding to more decimals than this, or to fewer than its negative, gives
 * the same result as any larger magnitude: every finite double has fewer
 * than 400 decimal places and is smaller than 10^400.
 */
const MAX_ROUNDING_DIGITS = 400;

/** The significant digits of `magnitude.toExponential(fractionDigits)`. */
function exponentialParts(
  magnitude: number,
  fractionDigits?: number
): { digits: string; exponent: number } {
  const [mantissa, exponentText] = magnitude
    .toExponential(fractionDigits)
    .split("e");
  return { digits: mantissa!.replace(".", ""), exponent: Number(exponentText) };
}

/**
 * Whether the exact value of `magnitude` ends with a 5 in its 16th
 * significant digit, as 1000000000000005 and 123456789012345.5 do. Such a
 * value lies exactly halfway between two 15-digit decimals.
 */
function isTieInSixteenthDigit(magnitude: number): boolean {
  // 41 significant digits of the exact binary value: an expansion that ends
  // at the 16th digit shows only zeros after it.
  const { digits } = exponentialParts(magnitude, 40);
  return digits[15] === "5" && !/[1-9]/.test(digits.slice(16));
}

/**
 * The decimal representation of a finite, non-zero magnitude with 15
 * significant digits, as Excel shows it: `|value|` is
 * `0.<digits> * 10^(exponent + 1)`, so 2.675 is
 * `{ digits: "267500000000000", exponent: 0 }` and 0.0125 is
 * `{ digits: "125000000000000", exponent: -2 }`.
 *
 * The digits are rounded half up from the exact binary value, except that a
 * value exactly halfway between two 15-digit decimals rounds toward zero, as
 * in Excel: 1000000000000005 is 1000000000000000.
 */
export function toSignificantDigits(magnitude: number): {
  digits: string;
  exponent: number;
} {
  if (isTieInSixteenthDigit(magnitude)) {
    // Exactly 16 digits; dropping the 5 rounds toward zero.
    const { digits, exponent } = exponentialParts(magnitude, 15);
    return { digits: digits.slice(0, SIGNIFICANT_DIGITS), exponent };
  }
  return exponentialParts(magnitude, SIGNIFICANT_DIGITS - 1);
}

/**
 * Rounds `value` to `digits` decimal places; a negative `digits` rounds to
 * tens, hundreds and so on: `roundDecimal(1234.5, -2, mode)` works on the
 * hundreds. `digits` is truncated to a whole number.
 *
 * Returns `undefined` when the rounded value is too large for a number, as
 * `roundDecimal(5, -400, "away-from-zero")` would be. The result is never
 * negative zero.
 *
 * When none of the 15 significant digits is rounded off, a value that has no
 * more digits than `digits` keeps is returned unchanged
 * (`roundDecimal(1/3, 20, mode)` is 1/3), and any other value becomes its
 * 15-digit decimal, which is what Excel shows: 0.1 + 0.2, stored as
 * 0.30000000000000004, rounds to 0.3 at 15 or 16 decimals.
 */
export function roundDecimal(
  value: number,
  digits: number,
  mode: DecimalRoundingMode
): number | undefined {
  if (!Number.isFinite(value)) {
    return undefined;
  }
  if (value === 0) {
    return 0;
  }
  const places = Math.max(
    -MAX_ROUNDING_DIGITS,
    Math.min(MAX_ROUNDING_DIGITS, Math.trunc(digits))
  );
  const magnitude = Math.abs(value);
  const { digits: significant, exponent } = toSignificantDigits(magnitude);
  const sign = value < 0 ? -1 : 1;

  // The number of significant digits that stay.
  const kept = exponent + 1 + places;
  if (kept >= SIGNIFICANT_DIGITS) {
    // The shortest digits that identify the number, such as the 16 of 1/3.
    if (exponentialParts(magnitude).digits.length <= kept) {
      return value;
    }
    return (
      sign * Number(`${significant}e${exponent - (SIGNIFICANT_DIGITS - 1)}`)
    );
  }

  // `prefix` holds the kept digits (none when the rounding position lies above
  // the first significant digit) and `rest` the digits rounded off.
  const prefix = kept > 0 ? significant.slice(0, kept) : "";
  const rest = kept > 0 ? significant.slice(kept) : significant;
  // Digits that are skipped entirely, when the rounding position lies more than
  // one place above the first significant digit, are zeros.
  const firstDropped = kept >= 0 ? rest[0]! : "0";

  let roundUp: boolean;
  switch (mode) {
    case "half-away-from-zero":
      roundUp = firstDropped >= "5";
      break;
    case "away-from-zero":
      roundUp = /[1-9]/.test(rest);
      break;
    case "toward-zero":
      roundUp = false;
      break;
  }

  // At most 15 digits, which is below 2^53 and therefore exact.
  const units = Number(prefix === "" ? "0" : prefix) + (roundUp ? 1 : 0);
  if (units === 0) {
    return 0;
  }
  const rounded = Number(`${units}e${-places}`);
  if (!Number.isFinite(rounded)) {
    return undefined;
  }
  return sign * rounded;
}
