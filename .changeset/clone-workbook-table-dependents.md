---
"@ricsam/formula-engine": patch
---

Recalculate references to a workbook's tables when `cloneWorkbook` creates it.
A reference written before the clone, such as `=INDEX([Copy]Results[Yield], 1)`,
stayed `#REF!`, and formulas in the clone that refer to its own tables could
reuse that stale result.
