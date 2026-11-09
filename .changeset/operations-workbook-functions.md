---
"@ricsam/formula-engine": minor
---

Add the functions the Operations workbooks evaluated to `#NAME?`: `IFS`,
`FILTER`, `TRANSPOSE`, `STDEV.S` (alias `STDEV`), `ROUND`, `ROUNDUP`,
`ROUNDDOWN`, `ISNUMBER`, `ISBLANK`, `SUMPRODUCT`, `VLOOKUP` and `TEXT`, for 59
functions in all. `GETPIVOTDATA` is not added: the engine has no pivot tables.

`IFS(test1, value1, ...)` returns the value of the first true condition and
`#N/A` when none is. Conditions are tested like `IF`'s and only the chosen
value is evaluated; an array condition chooses cell by cell and broadcasts like
`IF`. An odd number of arguments is `#VALUE!`.

`FILTER(array, include, [if_empty])` returns the rows of `array` selected by a
column `include`, or its columns for a row `include`. Nothing selected is
`if_empty`, or `#CALC!` without it; an error in `include` is the result. Whole
columns such as `FILTER(B:B, A:A="x")` work. Booleans are not numbers in this
engine, so Excel's `(A2:A9="x")*(B2:B9>5)` is `#VALUE!`; combine conditions
with `IF(A2:A9="x", B2:B9>5)`. `TRANSPOSE(array)` swaps rows and columns lazily
and spills.

`ROUND`, `ROUNDUP` and `ROUNDDOWN` round the decimal value Excel shows, with 15
significant digits, so `ROUND(2.675, 2)` is 2.68, `ROUNDUP(0.1+0.2, 1)` is 0.3
and `ROUNDDOWN(4.35*100, 0)` is 435. A negative `num_digits` rounds to tens,
hundreds and so on. Like `INT`, they accept only numbers, and array arguments
broadcast.

`STDEV.S` is the sample standard deviation of the numbers in its arguments,
ranges and spills, such as `STDEV.S(FILTER(...))`. Errors propagate and fewer
than two numbers is `#DIV/0!`. `SUMPRODUCT` adds the products of same-sized
arrays: a different size is `#VALUE!`, a cell that is not a number counts as 0,
and an error anywhere is the result. The `--(...)` idiom is `#VALUE!` here;
choose values with `IF` instead.

`VLOOKUP` searches the first column of a table like `MATCH`: exact for `FALSE`
or `0` (text ignores case, no wildcards) and the largest value less than or
equal to the lookup value for `TRUE`, the default. `col_index_num` below 1 is
`#VALUE!` and beyond the table `#REF!`, and a blank result cell is `""`. Array
lookup values and column numbers spill.

`ISNUMBER` and `ISBLANK` test a value without propagating errors and test
arrays element by element. Aggregates read every cell of a bounded result, so
`SUM(IF(ISBLANK(A1:A9), 1, 0))` counts the empty cells. `ISBLANK` is `TRUE` for
everything that reads as `""`, including a formula returning `""` and `""`
itself, where Excel is `TRUE` only for an empty cell.

`TEXT(value, format_text)` formats numbers, dates and times with Excel number
format codes: digit placeholders, thousands separators and scaling,
percentages, scientific notation, fractions, `General`, literal text, up to
four sections with conditions, currency tags, and date and time codes on
Excel's 1900 date system, including elapsed time. Numbers are rounded as Excel
shows them (`TEXT(2.675,"0.00")` is "2.68"). Following the engine's strict
typing, text is never converted to a number, an empty cell formats as `""`, and
`format_text` must be text. Array arguments broadcast.
