---
"@ricsam/formula-engine": patch
---

Write `#REF!` for a reference that a paste or fill moves off the sheet, and
evaluate error literals as their error. Pasting `=A1+1` one row higher stored
`=A-2+1`-style text that showed `#ERROR!`; it now stores `=#REF!+1`, which is
`#REF!` like in Excel. An error literal such as `=#N/A` or `=LEN(#REF!)` used to
show `#ERROR!` and now shows that error.
