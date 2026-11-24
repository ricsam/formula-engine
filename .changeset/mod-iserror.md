---
"@ricsam/formula-engine": minor
---

Add `MOD` and `ISERROR`, for 61 functions in all.

`MOD(number, divisor)` returns the remainder with the sign of the divisor, as
in Excel: `MOD(-3, 2)` is 1 and `MOD(3, -2)` is -1. A divisor of 0 is
`#DIV/0!`, and so is an omitted one; an omitted number is 0. A number that is
the divisor times a whole number in this engine's arithmetic has no remainder,
so `MOD(0.5, 0.1)` is 0 because `0.1*5` is 0.5; any other remainder is exact
for the stored binary values, so `MOD(0.3, 0.1)` is 0.09999999999999998, just
as `0.1*3=0.3` is `FALSE`. Following the engine's strict typing, text, logical
values and empty cells are `#VALUE!` where Excel converts them. `INFINITY` in
either argument and a quotient of 2^53 or more, where a number can no longer
tell consecutive multiples apart, are `#NUM!`. Array arguments broadcast, so
`MOD(SEQUENCE(6), 3)` spills 1, 2, 0, 1, 2, 0.

`ISERROR(value)` is `TRUE` for every error value, exactly where `IFERROR`
returns its fallback, and never propagates an error. It tests arrays element
by element, and aggregates read every cell of a bounded result, so
`SUM(IF(ISERROR(A1:A9), 1, 0))` counts the errors. A division by zero is
`INFINITY` in this engine, not an error, so `ISERROR(1/0)` is `FALSE` where
Excel is `TRUE`.
