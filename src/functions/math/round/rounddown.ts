import { createRoundingFunction } from "./round";

/**
 * ROUNDDOWN(number, num_digits)
 *
 * Rounds a number toward zero to `num_digits` decimal places:
 * `ROUNDDOWN(3.9, 0)` is 3, `ROUNDDOWN(-3.9, 0)` is -3 and
 * `ROUNDDOWN(1299, -2)` is 1200. Unlike INT, negative numbers move toward zero.
 *
 * The digits are those of the 15-significant-digit decimal representation, as
 * in Excel, so `ROUNDDOWN(4.35 * 100, 0)` is 435 even though the binary product
 * is slightly below it. Arguments, arrays, errors and omitted arguments follow
 * ROUND (see `round.ts`).
 */
export const ROUNDDOWN = createRoundingFunction("ROUNDDOWN", "toward-zero");
