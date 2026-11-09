---
"@ricsam/formula-engine": patch
---

Scope table and sheet renames to their workbook. Renaming a table or sheet
rewrote every reference to that name in every workbook, so with two batch
workbooks that each have a `Results` table or a `Data` sheet, renaming one
broke `=SUM(Results[Yield])` and `=Data!A1` in the other, and even references
qualified with the other workbook's name. Only references that resolve to the
renamed workbook's table or sheet change now. An unqualified reference in a
global named expression, which can resolve in any workbook, is still renamed.
