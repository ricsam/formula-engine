import { createRoundingFunction } from "./round";

/**
 * ROUNDUP(number, num_digits)
 *
 * Rounds a number away from zero to `num_digits` decimal places:
 * `ROUNDUP(3.2, 0)` is 4, `ROUNDUP(-3.2, 0)` is -4 and `ROUNDUP(1234.1, -2)` is
 * 1300. Numbers that already have at most `num_digits` decimals are unchanged.
 *
 * The digits are those of the 15-significant-digit decimal representation, as
 * in Excel, so `ROUNDUP(0.1 + 0.2, 1)` is 0.3 even though the binary sum is
 * slightly above it. Arguments, arrays, errors and omitted arguments follow
 * ROUND (see `round.ts`).
 */
export const ROUNDUP = createRoundingFunction("ROUNDUP", "away-from-zero");
