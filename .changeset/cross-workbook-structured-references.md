---
"@ricsam/formula-engine": minor
---

Reference tables in other workbooks as `[Workbook]Table[Column]`, the way
`[Workbook]Sheet!A1` references a cell there:
`=INDEX([Batch 1]Results[Yield], 2)` or
`=SUM([Batch 1]Results[Yield])+SUM([Batch 2]Results[Yield])`. Every structured
reference form works: `[Workbook]Table[@Column]` (the row of the formula cell),
column spans such as `[Workbook]Table[[Sample]:[Yield]]`, and the `#Headers`,
`#Data` and `#All` selectors with or without columns. These formulas were
syntax errors that showed `#ERROR!`. As before, a reference without the
workbook names a table in the formula's own workbook.

The references are tracked like cross-workbook cell references. They
recalculate when the table's data or size changes, a reference such as
`INDEX([Batch 1]Results[Yield], 1)` is `#REF!` while the workbook, table or
column is missing and resolves again once it is added, and renaming the
workbook, the table or a column header rewrites them. Formula analysis
resolves them to the other workbook's table and marks all of their tokens
`external`, like other references into another workbook.

Breaking: the previous spelling `[Workbook]!Table[Column]` is now a syntax
error whose message names the new spelling, and formulas are stored as
`[Workbook]Table[Column]`. Remove the `!` from stored formulas that use it. A
warm snapshot saved with such a reference restores cold, with a
`[FormulaEngine snapshot]` warning.
