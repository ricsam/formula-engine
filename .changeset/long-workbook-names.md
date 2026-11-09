---
"@ricsam/formula-engine": patch
---

Parse workbook names of any length. A workbook qualifier whose name had more
than eight words, or a whole-sheet alias such as `[Book]Sheet1` followed by
many more tokens, as in `SUM([Book]Sheet1, 1, 2, 3, 4, 5)`, was a syntax error.
