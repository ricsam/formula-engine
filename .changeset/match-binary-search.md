---
"@ricsam/formula-engine": patch
---

Search like Excel in `MATCH` with match type `1` (the default) and `-1`. Excel
does a binary search there without checking the order, and so does `MATCH`
now, so unsorted data returns the position Excel returns instead of `#N/A`.
Workbooks that leave out the `,0` of an exact match, such as
`=INDEX(MAPcondition[MAP Payload],MATCH([@MAP ID],MAPcondition[MAP ID]),0)`
over unsorted IDs, now give Excel's results: with `MAP-994`, `MAP-10977` and
`MAP-320` in `A1:A3`, `=MATCH("MAP-994",A1:A3)` is `3`, as in Excel. Sorted
data finds the same positions as before. Blank cells, errors and values of
another type are stepped over the way Excel steps over them, so
`=MATCH(9.99E+307,B:B)` finds the last number in column B whatever the order.
An unbounded range such as `B:B` is searched up to its last non-empty cell.
