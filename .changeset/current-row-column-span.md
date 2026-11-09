---
"@ricsam/formula-engine": patch
---

Fix unqualified current-row column spans such as `[@[pep_concentration]:[gate]]`,
which read only their first column. Formatting wrote the span as
`[@pep_concentration]`, and since the engine normalizes formula content when it
stores it, `=TEXTJOIN(", ", TRUE, [@[pep_concentration]:[gate]])` returned
`"10uM"` instead of `"10uM, RFP+"` and the stored formula lost the rest of the
span. `SUM([@[a]:[b]])` and other functions over such spans were affected the
same way. Table-qualified spans such as `Table1[@[a]:[b]]` already worked.
